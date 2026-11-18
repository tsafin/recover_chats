"""Atomic persistent state for resumable pipeline runs."""

from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


def atomic_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    except Exception:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass
        raise


@dataclass(slots=True)
class PipelineState:
    output: Path
    version: int = 1
    stages: dict[str, dict[str, Any]] = field(default_factory=dict)
    tasks: dict[str, dict[str, Any]] = field(default_factory=dict)

    @property
    def directory(self) -> Path:
        return self.output / ".pipeline"

    @property
    def path(self) -> Path:
        return self.directory / "state.json"

    @property
    def tasks_path(self) -> Path:
        return self.directory / "tasks.json"

    @classmethod
    def load(cls, output: Path) -> "PipelineState":
        state = cls(output.resolve())
        if state.path.exists():
            raw = json.loads(state.path.read_text(encoding="utf-8"))
            state.version = int(raw.get("version") or 1)
            state.stages = dict(raw.get("stages") or {})
        if state.tasks_path.exists():
            raw_tasks = json.loads(state.tasks_path.read_text(encoding="utf-8"))
            state.tasks = dict(raw_tasks.get("tasks") or {})
        return state

    def save(self) -> None:
        atomic_json(self.path, {"version": self.version, "stages": self.stages})
        atomic_json(self.tasks_path, {"version": self.version, "tasks": self.tasks})

    def record_stage(self, name: str, value: dict[str, Any]) -> None:
        self.stages[name] = value
        self.save()

    def record_task(self, key: str, value: dict[str, Any]) -> None:
        self.tasks[key] = value
        self.save()
