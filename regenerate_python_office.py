#!/usr/bin/env python3
"""Safely attempt self-contained Python generators from restored chats.

Recognises two families of generators:

1. **Office generators** — scripts using openpyxl, python-docx, python-pptx
   that declare absolute Claude output paths (``/mnt/user-data/outputs/*.xlsx``
   etc.).  Absolute paths are redirected to a ``regenerated`` folder.

2. **Media generators** — scripts using matplotlib, PIL/Pillow, cv2/opencv
   that write PNG/GIF/SVG/HTML to relative paths (``fig.savefig("out.png")``,
   ``cv2.imwrite("out.png")``, ``Path("page.html").write_text(...)``).
   These are run with ``cwd=regenerated`` so outputs land in the right place.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from artifact_timestamps import restore_artifact_mtime
from winenv import ensure_tools_on_path, to_posix_path

ensure_tools_on_path()


OFFICE_RE = re.compile(r"['\"](/(?:mnt/user-data/outputs|home/claude)/[^'\"]+\.(?:xlsx|docx|pptx))['\"]", re.I)
SAFE_LIBRARY_RE = re.compile(r"(?:openpyxl|from\s+docx\s+import|from\s+pptx\s+import)")
UNSAFE_RE = re.compile(r"(?:\bsubprocess\b|\bos\.system\b|\brequests\b|\burllib\b|\bsocket\b|\brmtree\b|\bPopen\b)")

# Media generator detection
MEDIA_LIBRARY_RE = re.compile(r"(?:import\s+matplotlib|from\s+matplotlib|import\s+PIL|from\s+PIL|import\s+cv2|from\s+cv2)")
MEDIA_OUTPUT_RE = re.compile(r"""['"]([\w./-]+\.(?:png|gif|svg|html?|jpg|jpeg|webp|bmp))['"]""", re.I)
# Patterns that indicate the script writes files (not just mentions a filename in a comment)
MEDIA_WRITE_RE = re.compile(r"(?:savefig|imwrite|write_text|\.save\s*\(|open\s*\([^)]+['\"]w)", re.I)
# HTML builder scripts: read images via base64/open and write HTML
HTML_BUILDER_RE = re.compile(r"(?:write_text\s*\(\s*['\"]?[^\)]*\.html|Path\s*\(\s*['\"][^'\"]*\.html)", re.I)


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else "restored_export/conversations").resolve()
    report: list[dict[str, object]] = []
    seen: set[tuple[Path, str]] = set()
    candidates = sorted(
        path
        for pattern in ("*/files/**/*.py", "*/inline_scripts/**/*.py")
        for path in root.glob(pattern)
        if path.is_file()
    )
    for source in candidates:
        text = source.read_text(encoding="utf-8", errors="replace")
        outputs = OFFICE_RE.findall(text)
        if not outputs or not SAFE_LIBRARY_RE.search(text):
            continue
        conv_dir = next(parent for parent in source.parents if parent.parent == root)
        regenerated = conv_dir / "regenerated"
        entry: dict[str, object] = {
            "source": source.relative_to(conv_dir).as_posix(),
            "declared_outputs": outputs,
            "status": "pending",
            "created": [],
        }
        if UNSAFE_RE.search(text):
            entry["status"] = "skipped-unsafe"
            report.append(entry)
            continue
        # Do not run several recovered copies of an identical generator.
        digest = hashlib.sha256(text.encode("utf-8")).hexdigest()
        marker = (conv_dir, digest)
        if marker in seen:
            entry["status"] = "skipped-duplicate"
            report.append(entry)
            continue
        seen.add(marker)
        regenerated.mkdir(parents=True, exist_ok=True)
        regenerated_posix = to_posix_path(regenerated)
        patched = text
        for original in outputs:
            patched = patched.replace(original, regenerated_posix + "/" + Path(original).name)
        # One early spreadsheet generator saved to /home/claude without using
        # the standard outputs directory; the general regexp above catches it.
        before = {path.name for path in regenerated.iterdir() if path.is_file()}
        with tempfile.TemporaryDirectory(prefix="claude-office-") as temp_dir:
            runner = Path(temp_dir) / source.name
            runner.write_text(patched, encoding="utf-8")
            try:
                result = subprocess.run(
                    [sys.executable, str(runner)],
                    cwd=source.parent,
                    text=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    timeout=90,
                    check=False,
                    env={**os.environ, "PYTHONPATH": str(source.parent) + os.pathsep + os.environ.get("PYTHONPATH", "")},
                )
                entry["exit_code"] = result.returncode
                entry["log"] = result.stdout[-8000:]
                entry["status"] = "ok" if result.returncode == 0 else "failed"
            except subprocess.TimeoutExpired as exc:
                entry["status"] = "timeout"
                entry["log"] = str(exc)
        after = {path.name for path in regenerated.iterdir() if path.is_file()}
        entry["created"] = sorted(after - before)
        timestamped: dict[str, object] = {}
        if entry["status"] == "ok":
            output_names = {Path(original).name for original in outputs} | (after - before)
            for output_name in sorted(output_names):
                output_path = regenerated / output_name
                if match := restore_artifact_mtime(conv_dir, output_path):
                    timestamped[output_name] = match
        entry["restored_mtimes"] = timestamped
        report.append(entry)

    # ── Pass 2: media generators (matplotlib/PIL/cv2 → PNG/GIF/HTML) ──
    # These scripts write to relative paths (fig.savefig("out.png")) rather
    # than absolute Claude paths.  Run them with cwd=regenerated so outputs
    # land in the right place.  Skip scripts already handled in pass 1.
    office_sources = {entry["source"] for entry in report}
    media_candidates = sorted(
        path
        for pattern in ("*/files/**/*.py", "*/inline_scripts/**/*.py")
        for path in root.glob(pattern)
        if path.is_file()
    )
    for source in media_candidates:
        text = source.read_text(encoding="utf-8", errors="replace")
        is_media = MEDIA_LIBRARY_RE.search(text) and MEDIA_WRITE_RE.search(text)
        is_html_builder = HTML_BUILDER_RE.search(text) and re.search(r"(?:base64|b64|read_bytes)", text)
        if not is_media and not is_html_builder:
            continue
        conv_dir = next(parent for parent in source.parents if parent.parent == root)
        rel_source = source.relative_to(conv_dir).as_posix()
        # Skip if already processed as an office generator
        if rel_source in office_sources:
            continue
        # Detect declared output filenames
        media_outputs = sorted({
            match for match in MEDIA_OUTPUT_RE.findall(text)
            if not match.startswith(("/", "http")) and "." in match
        })
        entry: dict[str, object] = {
            "source": rel_source,
            "declared_outputs": media_outputs,
            "status": "pending",
            "created": [],
            "category": "media",
        }
        if UNSAFE_RE.search(text):
            entry["status"] = "skipped-unsafe"
            report.append(entry)
            continue
        digest = hashlib.sha256(text.encode("utf-8")).hexdigest()
        marker = (conv_dir, digest)
        if marker in seen:
            entry["status"] = "skipped-duplicate"
            report.append(entry)
            continue
        seen.add(marker)
        regenerated = conv_dir / "regenerated"
        regenerated.mkdir(parents=True, exist_ok=True)
        before = {path.name for path in regenerated.iterdir() if path.is_file()}
        with tempfile.TemporaryDirectory(prefix="claude-media-") as temp_dir:
            runner = Path(temp_dir) / source.name
            # Patch absolute Claude paths if any, same as office generators
            patched = text
            for original in OFFICE_RE.findall(text):
                patched = patched.replace(original, to_posix_path(regenerated) + "/" + Path(original).name)
            runner.write_text(patched, encoding="utf-8")
            try:
                result = subprocess.run(
                    [sys.executable, str(runner)],
                    cwd=str(regenerated),  # run from regenerated/ so relative outputs land here
                    text=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    timeout=120,
                    check=False,
                    env={**os.environ, "PYTHONPATH": str(source.parent) + os.pathsep + os.environ.get("PYTHONPATH", "")},
                )
                entry["exit_code"] = result.returncode
                entry["log"] = result.stdout[-8000:]
                entry["status"] = "ok" if result.returncode == 0 else "failed"
            except subprocess.TimeoutExpired as exc:
                entry["status"] = "timeout"
                entry["log"] = str(exc)
        after = {path.name for path in regenerated.iterdir() if path.is_file()}
        entry["created"] = sorted(after - before)
        timestamped: dict[str, object] = {}
        if entry["status"] == "ok":
            output_names = set(media_outputs) | (after - before)
            for output_name in sorted(output_names):
                output_path = regenerated / output_name
                if match := restore_artifact_mtime(conv_dir, output_path):
                    timestamped[output_name] = match
        entry["restored_mtimes"] = timestamped
        report.append(entry)

    report_path = root.parent / "python_regeneration_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    summary = {
        "candidates": len(report),
        "ok": sum(item["status"] == "ok" for item in report),
        "failed": sum(item["status"] == "failed" for item in report),
        "skipped": sum(str(item["status"]).startswith("skipped") for item in report),
        "created_files": sum(len(item["created"]) for item in report),
        "report": str(report_path),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if not summary["failed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
