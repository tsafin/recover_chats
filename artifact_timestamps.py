#!/usr/bin/env python3
"""Restore regenerated artifact mtimes from the original conversation log."""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Any


def _basename(value: object) -> str:
    return PurePosixPath(str(value or "").replace("\\", "/")).name.casefold()


def _epoch_ns(value: str) -> int:
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    return int(parsed.timestamp()) * 1_000_000_000 + parsed.microsecond * 1_000


def find_artifact_timestamp(conv_dir: Path, output_name: str) -> dict[str, str] | None:
    """Return the latest log timestamp for the final version of output_name.

    Exact presentation records are preferred.  Some generators leave useful
    intermediate Office files that were never presented; for those, fall back
    to the latest message containing the exact basename in its recorded JSON.
    """
    conversation_path = conv_dir / "conversation.json"
    if not conversation_path.exists():
        return None
    conversation = json.loads(conversation_path.read_text(encoding="utf-8"))
    messages = conversation.get("chat_messages") or []
    message_by_uuid = {str(message.get("uuid")): message for message in messages}
    wanted = output_name.casefold()
    candidates: list[dict[str, str]] = []

    presented_path = conv_dir / "presented_files.json"
    if presented_path.exists():
        for record in json.loads(presented_path.read_text(encoding="utf-8")):
            if _basename(record.get("file_path")) != wanted:
                continue
            message_uuid = str(record.get("message_uuid") or "")
            timestamp = str((message_by_uuid.get(message_uuid) or {}).get("created_at") or "")
            if timestamp:
                candidates.append({"timestamp": timestamp, "source": "presented-file", "message_uuid": message_uuid})

    if not candidates:
        for message in messages:
            # JSON matching covers command text, tool inputs and tool results,
            # while requiring the complete basename avoids loose stem matches.
            if wanted not in json.dumps(message, ensure_ascii=False).casefold():
                continue
            timestamp = str(message.get("created_at") or "")
            if timestamp:
                candidates.append({
                    "timestamp": timestamp,
                    "source": "message-mention",
                    "message_uuid": str(message.get("uuid") or ""),
                })

    if not candidates:
        return None
    dated: list[tuple[int, dict[str, str]]] = []
    for candidate in candidates:
        try:
            dated.append((_epoch_ns(candidate["timestamp"]), candidate))
        except ValueError:
            continue
    return max(dated, key=lambda item: item[0])[1] if dated else None


def restore_artifact_mtime(conv_dir: Path, output_path: Path) -> dict[str, str] | None:
    """Set output_path mtime from its log record, preserving its current atime."""
    match = find_artifact_timestamp(conv_dir, output_path.name)
    if match is None or not output_path.is_file():
        return None
    stat = output_path.stat()
    os.utime(output_path, ns=(stat.st_atime_ns, _epoch_ns(match["timestamp"])))
    return match


def backfill_tree(conversations: Path) -> dict[str, Any]:
    restored: list[dict[str, str]] = []
    unmatched: list[str] = []
    for output in sorted(conversations.glob("*/regenerated/*")):
        if not output.is_file():
            continue
        match = restore_artifact_mtime(output.parent.parent, output)
        relative = output.relative_to(conversations).as_posix()
        if match:
            restored.append({"file": relative, **match})
        else:
            unmatched.append(relative)
    return {"restored": restored, "unmatched": unmatched}


def main() -> int:
    conversations = Path(sys.argv[1] if len(sys.argv) > 1 else "restored_export/conversations").resolve()
    result = backfill_tree(conversations)
    print(json.dumps({"restored": len(result["restored"]), "unmatched": result["unmatched"]}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
