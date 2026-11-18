"""Typed public contracts shared by restoration pipeline stages."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


@dataclass(slots=True)
class PipelineConfig:
    source: Path
    output: Path
    node_modules: Path | None = None
    skip_code_execution: bool = False
    strict_generators: bool = False
    max_passes: int = 8
    from_stage: str | None = None
    only_stage: str | None = None

    def normalized(self) -> "PipelineConfig":
        source = self.source.resolve()
        output = self.output.resolve()
        node_modules = (self.node_modules or output / "runtime" / "node_modules").resolve()
        return PipelineConfig(
            source=source,
            output=output,
            node_modules=node_modules,
            skip_code_execution=self.skip_code_execution,
            strict_generators=self.strict_generators,
            max_passes=max(1, self.max_passes),
            from_stage=self.from_stage,
            only_stage=self.only_stage,
        )

    @property
    def conversations(self) -> Path:
        return self.output / "conversations"

    def to_dict(self) -> dict[str, Any]:
        value = asdict(self)
        for key in ("source", "output", "node_modules"):
            if value[key] is not None:
                value[key] = str(value[key])
        return value


@dataclass(slots=True)
class StageResult:
    name: str
    status: str = "ok"
    exit_code: int = 0
    started_at: str = field(default_factory=utc_now)
    finished_at: str = ""
    summary: dict[str, Any] = field(default_factory=dict)
    created: list[str] = field(default_factory=list)
    updated: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    report: str | None = None
    fatal: bool = False

    def finish(self) -> "StageResult":
        if not self.finished_at:
            self.finished_at = utc_now()
        return self

    def to_dict(self) -> dict[str, Any]:
        return asdict(self.finish())


@dataclass(slots=True)
class PipelineResult:
    config: PipelineConfig
    stages: list[StageResult] = field(default_factory=list)
    started_at: str = field(default_factory=utc_now)
    finished_at: str = ""

    @property
    def exit_code(self) -> int:
        return 1 if any(stage.fatal and stage.exit_code for stage in self.stages) else 0

    def to_dict(self) -> dict[str, Any]:
        if not self.finished_at:
            self.finished_at = utc_now()
        return {
            "config": self.config.to_dict(),
            "started_at": self.started_at,
            "finished_at": self.finished_at,
            "exit_code": self.exit_code,
            "stages": [stage.to_dict() for stage in self.stages],
        }
