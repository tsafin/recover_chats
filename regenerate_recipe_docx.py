"""Convert locally selected Markdown and TeX inputs to DOCX."""

import argparse
import json
from pathlib import Path

from regenerate_known_office import regenerate_local


def regenerate_recipes(root: Path) -> tuple[list[dict], dict]:
    return regenerate_local(root, docx=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("conversations", type=Path, nargs="?",
                        default=Path("restored_export/conversations"))
    args = parser.parse_args()
    _, summary = regenerate_recipes(args.conversations)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return int(bool(summary["failed"]))


if __name__ == "__main__":
    raise SystemExit(main())
