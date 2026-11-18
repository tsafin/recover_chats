#!/usr/bin/env python3
"""Run the complete restartable Claude export restoration pipeline."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from pipeline.model import PipelineConfig
from pipeline.orchestrator import STAGES, run_pipeline


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path, default=Path("restored_export"))
    parser.add_argument("--node-modules", type=Path)
    parser.add_argument("--skip-code-execution", action="store_true")
    parser.add_argument("--strict-generators", action="store_true")
    parser.add_argument("--max-passes", type=int, default=8)
    parser.add_argument("--from-stage", choices=STAGES)
    parser.add_argument("--only-stage", choices=STAGES)
    args = parser.parse_args()
    config = PipelineConfig(
        source=args.source,
        output=args.output,
        node_modules=args.node_modules,
        skip_code_execution=args.skip_code_execution,
        strict_generators=args.strict_generators,
        max_passes=args.max_passes,
        from_stage=args.from_stage,
        only_stage=args.only_stage,
    )
    result = run_pipeline(config)
    print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2))
    return result.exit_code


if __name__ == "__main__":
    raise SystemExit(main())
