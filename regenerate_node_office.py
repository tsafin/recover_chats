#!/usr/bin/env python3
"""Attempt recovered Node.js PPTX/DOCX generators with local dependencies."""

from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
import sys
from pathlib import Path


LIB_RE = re.compile(r"require\(['\"](?:pptxgenjs|docx)['\"]\)")
OUTPUT_RE = re.compile(r"(['\"])(/[^'\"]+\.(?:pptx|docx|xlsx))\1", re.I)
DIRECT_WRITE_RE = re.compile(r"(writeFile\s*\(\s*\{\s*fileName\s*:\s*)(['\"])([^'\"]+\.(?:pptx|docx|xlsx))\2", re.I)
UNSAFE_RE = re.compile(r"(?:child_process|\bexecSync\b|\bspawn\b|\brmSync\b|\brmdirSync\b)")


def main() -> int:
    conversations = Path(sys.argv[1] if len(sys.argv) > 1 else "restored_export/conversations").resolve()
    runtime = Path(sys.argv[2] if len(sys.argv) > 2 else "restored_export/runtime/node_modules").resolve()
    report: list[dict[str, object]] = []
    seen: set[tuple[Path, str]] = set()
    candidates = sorted(
        path
        for pattern in ("*/files/**/*.js", "*/inline_scripts/**/*.js")
        for path in conversations.glob(pattern)
        if path.is_file()
    )
    for source in candidates:
        text = source.read_text(encoding="utf-8", errors="replace")
        if not LIB_RE.search(text) or "writeFile" not in text:
            continue
        conv_dir = next(parent for parent in source.parents if parent.parent == conversations)
        regenerated = conv_dir / "regenerated"
        regenerated.mkdir(parents=True, exist_ok=True)
        entry: dict[str, object] = {"source": source.relative_to(conv_dir).as_posix(), "status": "pending", "created": []}
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

        declared: list[str] = []
        def replace_absolute(match: re.Match[str]) -> str:
            original = match.group(2)
            declared.append(original)
            return match.group(1) + str(regenerated / Path(original).name) + match.group(1)

        patched = OUTPUT_RE.sub(replace_absolute, text)
        def replace_direct(match: re.Match[str]) -> str:
            original = match.group(3)
            declared.append(original)
            return match.group(1) + match.group(2) + str(regenerated / Path(original).name) + match.group(2)

        patched = DIRECT_WRITE_RE.sub(replace_direct, patched)
        entry["declared_outputs"] = sorted(set(declared))
        before = {path.name for path in regenerated.iterdir() if path.is_file()}
        env = {**os.environ, "NODE_PATH": str(runtime) + os.pathsep + os.environ.get("NODE_PATH", "")}
        try:
            result = subprocess.run(
                ["node"],
                input=patched,
                cwd=source.parent,
                env=env,
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                timeout=120,
                check=False,
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

    report_path = conversations.parent / "node_regeneration_report.json"
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
