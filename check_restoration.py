#!/usr/bin/env python3
"""Check local HTML links and OOXML ZIP integrity in a restored export."""

from __future__ import annotations

import json
import sys
import zipfile
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlsplit


class Links(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.hrefs: list[tuple[str, bool]] = []
        self.message_payloads: list[str] = []
        self._in_message_data = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "script" and dict(attrs).get("id") == "message-data":
            self._in_message_data = True
        if tag == "a":
            attributes = dict(attrs)
            for name, value in attrs:
                if name == "href" and value:
                    self.hrefs.append((value, attributes.get("data-original-reference") == "true"))

    def handle_data(self, data: str) -> None:
        if self._in_message_data:
            self.message_payloads.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag == "script":
            self._in_message_data = False


def validate_restoration(root: Path) -> dict[str, object]:
    """Validate links and OOXML containers, returning the full report."""
    root = root.resolve()
    broken: list[dict[str, str]] = []
    links = 0
    html_files = list(root.glob("INDEX*.html")) + [root / "raw" / "artifacts" / "INDEX.html"]
    data_files = sorted(root.glob("conversations/*/chat-data.js"))
    html_files.extend(root.glob("viewer*.html"))
    if not data_files:
        html_files.extend(root.glob("conversations/*/chat*.html"))
    html_files = [path for path in html_files if path.exists()]
    def pages():
        for path in html_files:
            yield path, path.read_text(encoding="utf-8", errors="replace")
        prefix = "window.restoredChatData = "
        for path in data_files:
            source = path.read_text(encoding="utf-8").strip()
            if not source.startswith(prefix) or not source.endswith(";"):
                raise ValueError(f"Invalid chat data: {path}")
            for markup in json.loads(source[len(prefix):-1])["pages"].values():
                yield path, markup

    for html_file, markup in pages():
        parser = Links()
        parser.feed(markup)
        for payload in parser.message_payloads:
            for markup, _height in json.loads(payload):
                parser.feed(markup)
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
            elif target.parent == root and target.name.startswith("viewer") and target.suffix == ".html":
                folders = parse_qs(parsed.query).get("chat", [])
                if (len(folders) != 1 or folders[0] in {".", ".."}
                        or "/" in folders[0] or "\\" in folders[0]
                        or not (root / "conversations" / folders[0] / "chat-data.js").is_file()):
                    broken.append({"html": str(html_file.relative_to(root)), "href": raw_href,
                                   "reason": "missing or invalid chat data"})

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
        "chat_data_files": len(data_files),
        "local_links": links,
        "broken_links": broken,
        "office_files": len(office),
        "corrupt_office_files": corrupt,
    }
    (root / "validation_report.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return result


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else "restored_export")
    result = validate_restoration(root)
    print(json.dumps({key: len(value) if isinstance(value, list) else value for key, value in result.items()}, ensure_ascii=False, indent=2))
    return 1 if result["broken_links"] or result["corrupt_office_files"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
