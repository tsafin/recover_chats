#!/usr/bin/env python3
"""Safely attempt self-contained Python Office generators from restored chats.

Only scripts using known local Office libraries are considered.  Scripts that
reference networking, subprocesses, shell execution, or destructive filesystem
operations are skipped.  Absolute Claude output paths are redirected to a
`regenerated` folder inside the corresponding conversation.
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


OFFICE_RE = re.compile(r"['\"](/(?:mnt/user-data/outputs|home/claude)/[^'\"]+\.(?:xlsx|docx|pptx))['\"]", re.I)
SAFE_LIBRARY_RE = re.compile(r"(?:openpyxl|from\s+docx\s+import|from\s+pptx\s+import)")
UNSAFE_RE = re.compile(r"(?:\bsubprocess\b|\bos\.system\b|\brequests\b|\burllib\b|\bsocket\b|\brmtree\b|\bPopen\b)")


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else "restored_export/conversations").resolve()
    report: list[dict[str, object]] = []
    seen: set[tuple[Path, str]] = set()
    candidates = sorted(path for path in root.glob("*/files/**/*.py") if path.is_file())
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
        patched = text
        for original in outputs:
            patched = patched.replace(original, str(regenerated / Path(original).name))
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
