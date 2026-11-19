"""High-level orchestration for the complete restoration workflow."""

from __future__ import annotations

import traceback
from collections.abc import Callable
from pathlib import Path
from typing import Any

from artifact_timestamps import restore_timestamps
from check_restoration import validate_restoration
from regenerate_known_office import regenerate_known
from regenerate_recipe_docx import regenerate_recipes
from regenerate_node_office import regenerate_node
from regenerate_python_office import regenerate_python
from restore_export import restore_export

from .model import PipelineConfig, PipelineResult, StageResult, utc_now
from .state import PipelineState, atomic_json


STAGES = ("restore", "python", "node", "recipes", "known", "timestamps", "render", "validate")


def _selected(config: PipelineConfig, name: str) -> bool:
    if config.only_stage:
        return name == config.only_stage
    if config.from_stage:
        return STAGES.index(name) >= STAGES.index(config.from_stage)
    return True


def _created(report: list[dict[str, Any]]) -> list[str]:
    values: list[str] = []
    for entry in report:
        conversation = str(entry.get("conversation") or "")
        for value in entry.get("created") or []:
            values.append(f"{conversation}/{value}" if conversation else str(value))
        if entry.get("status") == "ok" and entry.get("output"):
            values.append(f"{conversation}/{entry['output']}" if conversation else str(entry["output"]))
    return sorted(set(values))


def run_pipeline(config: PipelineConfig) -> PipelineResult:
    config = config.normalized()
    config.output.mkdir(parents=True, exist_ok=True)
    state = PipelineState.load(config.output)
    result = PipelineResult(config)

    def execute(name: str, action: Callable[[], StageResult]) -> StageResult | None:
        if not _selected(config, name):
            return None
        started_at = utc_now()
        try:
            stage = action()
        except Exception as exc:
            stage = StageResult(
                name=name,
                status="failed",
                exit_code=1,
                fatal=name in {"restore", "render", "validate"},
                summary={"error": f"{type(exc).__name__}: {exc}", "traceback": traceback.format_exc()},
            )
        stage.started_at = started_at
        stage.finished_at = utc_now()
        result.stages.append(stage)
        state.record_stage(name, stage.to_dict())
        return stage

    restore_stage = execute("restore", lambda: StageResult("restore", summary=restore_export(config.source, config.output), report=str(config.output / "SUMMARY.json")))
    if restore_stage and restore_stage.fatal and restore_stage.exit_code:
        atomic_json(config.output / "pipeline_report.json", result.to_dict())
        return result

    if not config.skip_code_execution:
        def python_stage() -> StageResult:
            report, summary = regenerate_python(config.conversations, config.max_passes)
            failed = int(summary.get("failed") or 0) + int(summary.get("blocked") or 0)
            return StageResult(
                "python",
                status="partial" if failed else "ok",
                exit_code=1 if failed else 0,
                fatal=config.strict_generators and bool(failed),
                summary=summary,
                created=_created(report),
                report=str(summary.get("report") or ""),
            )

        execute("python", python_stage)

        def node_stage() -> StageResult:
            assert config.node_modules is not None
            if not config.node_modules.is_dir():
                return StageResult("node", status="skipped", warnings=[f"node_modules not found: {config.node_modules}"])
            report, summary = regenerate_node(config.conversations, config.node_modules)
            failed = int(summary.get("failed") or 0)
            return StageResult("node", status="partial" if failed else "ok", exit_code=1 if failed else 0, fatal=config.strict_generators and bool(failed), summary=summary, created=_created(report), report=str(summary.get("report") or ""))

        execute("node", node_stage)

        def local_stage(name: str, action: Callable) -> StageResult:
            report, summary = action(config.conversations)
            failed = bool(summary["failed"])
            return StageResult(
                name, status="partial" if failed else "ok" if report else "skipped",
                exit_code=int(failed), fatal=config.strict_generators and failed,
                summary=summary, created=_created(report), report=str(summary["report"]),
            )

        execute("recipes", lambda: local_stage("recipes", regenerate_recipes))
        execute("known", lambda: local_stage("known", regenerate_known))
    else:
        for name in ("python", "node", "recipes", "known"):
            if _selected(config, name):
                stage = StageResult(name, status="skipped", warnings=["recovered code execution disabled"]).finish()
                result.stages.append(stage)
                state.record_stage(name, stage.to_dict())

    execute("timestamps", lambda: StageResult("timestamps", summary=restore_timestamps(config.conversations)))
    render_stage = execute("render", lambda: StageResult("render", summary=restore_export(config.source, config.output), report=str(config.output / "SUMMARY.json")))
    if render_stage and render_stage.fatal and render_stage.exit_code:
        atomic_json(config.output / "pipeline_report.json", result.to_dict())
        return result

    def validation_stage() -> StageResult:
        report = validate_restoration(config.output)
        failed = bool(report["broken_links"] or report["corrupt_office_files"])
        return StageResult("validate", status="failed" if failed else "ok", exit_code=1 if failed else 0, fatal=True, summary=report, report=str(config.output / "validation_report.json"))

    execute("validate", validation_stage)
    atomic_json(config.output / "pipeline_report.json", result.to_dict())
    return result
