#!/usr/bin/env python3
"""Check local HTML links and OOXML ZIP integrity in a restored export."""

from __future__ import annotations

import json
import sys
import zipfile
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote, urlsplit


class Links(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.hrefs: list[tuple[str, bool]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "a":
            attributes = dict(attrs)
            for name, value in attrs:
                if name == "href" and value:
                    self.hrefs.append((value, attributes.get("data-original-reference") == "true"))


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else "restored_export").resolve()
    broken: list[dict[str, str]] = []
    links = 0
    html_files = list(root.glob("INDEX*.html")) + [root / "raw" / "artifacts" / "INDEX.html"]
    html_files.extend(root.glob("conversations/*/chat*.html"))
    html_files = [path for path in html_files if path.exists()]
    for html_file in html_files:
        parser = Links()
        parser.feed(html_file.read_text(encoding="utf-8", errors="replace"))
        for raw_href, original_reference in parser.hrefs:
            if original_reference:
                continue
            try:
                parsed = urlsplit(raw_href)
            except ValueError:
                broken.append({"html": str(html_file.relative_to(root)), "href": raw_href, "reason": "invalid URL"})
                continue
            if parsed.scheme or raw_href.startswith("#"):
                continue
            links += 1
            target = (html_file.parent / unquote(parsed.path)).resolve()
            try:
                target.relative_to(root)
            except ValueError:
                broken.append({"html": str(html_file.relative_to(root)), "href": raw_href, "reason": "outside root"})
                continue
            if not target.exists():
                broken.append({"html": str(html_file.relative_to(root)), "href": raw_href, "reason": "missing"})

    office = [path for path in root.glob("conversations/*/regenerated/*") if path.suffix.lower() in {".pptx", ".docx", ".xlsx"}]
    corrupt: list[dict[str, str]] = []
    for path in office:
        try:
            with zipfile.ZipFile(path) as archive:
                bad = archive.testzip()
                if bad:
                    corrupt.append({"file": str(path.relative_to(root)), "reason": f"CRC error: {bad}"})
        except Exception as exc:
            corrupt.append({"file": str(path.relative_to(root)), "reason": f"{type(exc).__name__}: {exc}"})
    result = {
        "html_files": len(html_files),
        "local_links": links,
        "broken_links": broken,
        "office_files": len(office),
        "corrupt_office_files": corrupt,
    }
    (root / "validation_report.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: len(value) if isinstance(value, list) else value for key, value in result.items()}, ensure_ascii=False, indent=2))
    return 1 if broken or corrupt else 0


if __name__ == "__main__":
    raise SystemExit(main())
