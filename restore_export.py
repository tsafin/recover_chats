#!/usr/bin/env python3
"""Restore a Claude export into per-conversation folders.

The Claude export contains conversation JSON and metadata for presented files,
but generally does not contain the bytes of those presented files.  This script
recovers everything that is actually present: chats, attachments with extracted
text, Write/Edit/create_file state, shell commands, heredoc-created sources,
projects, memories, and artifacts.  It also writes a manifest of presented
files so missing binary results are explicit rather than silently ignored.
"""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import os
import re
import shutil
import sys
import zipfile
from collections import Counter
from html.parser import HTMLParser
from pathlib import Path, PurePosixPath
from typing import Any, Iterable
from urllib.parse import quote, unquote

from markdown_it import MarkdownIt


OFFICE_EXTENSIONS = {".pptx", ".docx", ".xlsx", ".pdf", ".odt", ".ods", ".odp"}
COMMAND_TOOLS = {"Bash", "bash_tool"}
WRITE_TOOLS = {"Write", "create_file"}
EDIT_TOOLS = {"Edit", "str_replace"}
PRESENT_TOOLS = {"present_files", "SendUserFile", "mcp__Claude_Docs__export"}

# ── i18n ──────────────────────────────────────────────────────────────

LOCALES_DIR = Path(__file__).resolve().parent / "locales"
DEFAULT_LANG = "ru"


def load_locales() -> dict[str, dict[str, Any]]:
    """Load all locale JSON files from the locales/ directory."""
    locales: dict[str, dict[str, Any]] = {}
    if not LOCALES_DIR.is_dir():
        return locales
    for path in sorted(LOCALES_DIR.glob("*.json")):
        data = json.loads(path.read_text(encoding="utf-8"))
        code = data.get("lang_code") or path.stem
        locales[code] = data
    return locales


LOCALES = load_locales()
AVAILABLE_LANGS = list(LOCALES.keys())


def L(key: str, lang: str | None = None) -> str:
    """Look up a dotted i18n key like 'index.brand' in the given locale.

    Falls back to DEFAULT_LANG, then to the key itself.
    """
    lang = lang or DEFAULT_LANG
    locale = LOCALES.get(lang) or LOCALES.get(DEFAULT_LANG) or {}
    parts = key.split(".")
    value: Any = locale
    for part in parts:
        if isinstance(value, dict):
            value = value.get(part)
        else:
            value = None
            break
    if value is None and lang != DEFAULT_LANG:
        return L(key, DEFAULT_LANG)
    return str(value) if value is not None else key


def lang_suffix(lang: str) -> str:
    """Return '' for the default language, '_<lang>' otherwise."""
    return "" if lang == DEFAULT_LANG else f"_{lang}"


def lang_switcher_html(current_lang: str, relative_root: str = "") -> str:
    """Render a compact language switcher dropdown.

    *relative_root* is the path prefix from the HTML file to the output root
    (e.g. '../..' for a conversation page, '' for INDEX.html).
    """
    if len(AVAILABLE_LANGS) < 2:
        return ""
    options = []
    for code in AVAILABLE_LANGS:
        name = LOCALES[code].get("lang_name", code)
        selected = " selected" if code == current_lang else ""
        options.append(f'<option value="{code}"{selected}>{html.escape(name)}</option>')
    return (
        f'<select class="lang-switch" onchange="'
        f"document.cookie='restore_lang='+this.value+';path=/;max-age=31536000';"
        f"sessionStorage.setItem('restore_chat_scroll',window.scrollY);"
        f"var d=document.location.pathname.replace(/(_(en|zh|ru))\\.html/,'.html');"
        f"var base=d.replace(/\\.html$/,'');"
        f"document.location.href=base+('{{suffix}}'==='{{default}}'?'':'_'+this.value)+'.html'"
        f'">{"".join(options)}</select>'
    ).replace("{{suffix}}", "").replace("{{default}}", DEFAULT_LANG)


def json_dump(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def safe_component(value: Any, limit: int = 100) -> str:
    text = str(value or "untitled").replace("\x00", "")
    text = re.sub(r"[\\/:*?\"<>|\x00-\x1f]", "_", text)
    text = re.sub(r"\s+", " ", text).strip(" .") or "untitled"
    if len(text) <= limit:
        return text
    digest = hashlib.sha1(text.encode("utf-8")).hexdigest()[:8]
    return f"{text[: limit - 9].rstrip()}_{digest}"


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    index = 2
    while True:
        candidate = path.with_name(f"{stem}_{index}{suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def unpack_archives(source: Path, raw_dir: Path) -> list[dict[str, Any]]:
    raw_dir.mkdir(parents=True, exist_ok=True)
    report: list[dict[str, Any]] = []
    for archive in sorted(source.glob("*.zip")):
        entry = {"archive": archive.name, "members": 0, "bytes": 0, "error": None}
        try:
            with zipfile.ZipFile(archive) as zf:
                bad = zf.testzip()
                if bad:
                    raise zipfile.BadZipFile(f"CRC error in {bad}")
                for info in zf.infolist():
                    member = PurePosixPath(info.filename)
                    if member.is_absolute() or ".." in member.parts:
                        raise ValueError(f"unsafe ZIP member: {info.filename}")
                zf.extractall(raw_dir)
                entry["members"] = len(zf.infolist())
                entry["bytes"] = sum(item.file_size for item in zf.infolist())
        except Exception as exc:  # keep a complete audit report
            entry["error"] = f"{type(exc).__name__}: {exc}"
        report.append(entry)
    return report


def path_key(raw_path: str, cwd: str | None = None) -> str:
    raw = os.path.expandvars(str(raw_path).strip().strip("'\""))
    if not raw.startswith("/") and cwd:
        raw = f"{cwd.rstrip('/')}/{raw}"
    parts: list[str] = []
    for part in PurePosixPath(raw).parts:
        if part in {"", "/", "."}:
            continue
        if part == "..":
            if parts:
                parts.pop()
            continue
        parts.append(part)
    return "/".join(parts) or "unnamed"


def local_parts(raw_path: str) -> tuple[str, ...]:
    key = path_key(raw_path)
    prefixes = (
        ("mnt/user-data/outputs/", "outputs/"),
        ("mnt/user-data/uploads/", "uploads/"),
        ("mnt/data/", "uploads/"),
    )
    for prefix, replacement in prefixes:
        if key.startswith(prefix):
            key = replacement + key[len(prefix) :]
            break
    else:
        marker = "/scratchpad/"
        if marker in f"/{key}":
            key = "scratchpad/" + key.split(marker, 1)[1]
        else:
            key = "filesystem/" + key
    return tuple(safe_component(part, 120) for part in key.split("/") if part)


def write_vfs(files_dir: Path, raw_path: str, content: str) -> Path:
    target = files_dir.joinpath(*local_parts(raw_path))
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(content, encoding="utf-8")
    return target


def iter_blocks(conversation: dict[str, Any]) -> Iterable[tuple[dict[str, Any], dict[str, Any]]]:
    for message in conversation.get("chat_messages") or []:
        for block in message.get("content") or []:
            if isinstance(block, dict):
                yield message, block


def initial_cwd(command: str) -> str | None:
    match = re.search(r"(?:^|[;&|]\s*)cd\s+(?:--\s+)?(['\"]?)([^;&|\n]+?)\1\s*(?:&&|;|$)", command)
    return match.group(2).strip() if match else None


HEREDOC_RE = re.compile(
    r"(?P<head>(?:cat|tee|python(?:3)?|node|bash|sh)\b[^\n]*?)"
    r"<<-?\s*(?P<quote>['\"]?)(?P<delim>[A-Za-z_][A-Za-z0-9_]*)\2[^\n]*\n",
    re.MULTILINE,
)


def heredocs(command: str) -> Iterable[dict[str, Any]]:
    """Yield heredoc bodies and cat/tee target paths without executing commands."""
    cursor = 0
    while True:
        match = HEREDOC_RE.search(command, cursor)
        if not match:
            return
        delimiter = match.group("delim")
        end_re = re.compile(rf"(?m)^[\t ]*{re.escape(delimiter)}[\t ]*(?:\r?\n|$)")
        end = end_re.search(command, match.end())
        if not end:
            return
        body = command[match.end() : end.start()]
        head = match.group("head")
        full_header = command[command.rfind("\n", 0, match.start()) + 1 : match.end()].rstrip()
        target = None
        append = False
        if re.search(r"\b(?:cat|tee)\b", head):
            path_match = re.search(r"(?:>>|>)\s*([^\s;&|]+)", full_header)
            if path_match:
                target = path_match.group(1).strip("'\"")
                append = ">>" in path_match.group(0)
            elif head.lstrip().startswith("tee"):
                words = re.findall(r"(?:[^\s'\"]|\"[^\"]*\"|'[^']*')+", head)
                candidates = [w for w in words[1:] if not w.startswith("-")]
                if candidates:
                    target = candidates[-1].strip("'\"")
                    append = "-a" in words
        yield {
            "header": full_header,
            "body": body,
            "target": target,
            "append": append,
            "interpreter": head.strip().split()[0],
        }
        cursor = end.end()


def render_chat_markdown(conversation: dict[str, Any]) -> str:
    out = [f"# {conversation.get('name') or 'Untitled'}", ""]
    out.append(f"- UUID: `{conversation.get('uuid', '')}`")
    out.append(f"- Created: {conversation.get('created_at', '')}")
    out.append(f"- Updated: {conversation.get('updated_at', '')}")
    if conversation.get("summary"):
        out.extend(["", "## Summary", "", str(conversation["summary"])])
    for index, message in enumerate(conversation.get("chat_messages") or [], 1):
        sender = "User" if message.get("sender") == "human" else "Claude"
        out.extend(["", f"## {index}. {sender}", "", f"_{message.get('created_at', '')}_", ""])
        blocks = message.get("content") or []
        if not blocks and message.get("text"):
            out.append(str(message["text"]))
        for block in blocks:
            if not isinstance(block, dict):
                continue
            kind = block.get("type")
            if kind == "text":
                out.append(str(block.get("text") or ""))
            elif kind == "thinking":
                out.extend(["", "<details><summary>Thinking</summary>", "", str(block.get("thinking") or ""), "", "</details>"])
            elif kind == "tool_use":
                payload = json.dumps(block.get("input") or {}, ensure_ascii=False, indent=2)
                out.extend(["", f"<details><summary>Tool: {block.get('name', 'unknown')}</summary>", "", "```json", payload, "```", "", "</details>"])
            elif kind == "tool_result":
                out.extend(["", f"<details><summary>Tool result: {block.get('name', 'unknown')}</summary>", ""])
                for item in block.get("content") or []:
                    if isinstance(item, dict) and item.get("type") == "text":
                        out.extend(["```text", str(item.get("text") or ""), "```"])
                    elif isinstance(item, dict) and item.get("type") == "local_resource":
                        out.append(f"- Presented file: `{item.get('file_path', '')}` ({item.get('mime_type', '')}, UUID `{item.get('uuid', '')}`)")
                out.extend(["", "</details>"])
    return "\n".join(out).rstrip() + "\n"


def href(path: str) -> str:
    return quote(path.replace(os.sep, "/"), safe="/#?=&%")


def linked_text(value: Any, resource_links: dict[str, str]) -> str:
    """Escape message text and link URLs and known Claude filesystem paths."""
    text = str(value or "")
    tokens: list[tuple[int, int, str]] = []
    for match in re.finditer(r"https?://[^\s<>()\"',}\]]+", text):
        tokens.append((match.start(), match.end(), f'<a href="{html.escape(match.group(0), quote=True)}">{html.escape(match.group(0))}</a>'))
    for original, local in resource_links.items():
        start = 0
        while original and (pos := text.find(original, start)) >= 0:
            tokens.append((pos, pos + len(original), f'<a class="file-link" href="{href(local)}">{html.escape(original)}</a>'))
            start = pos + len(original)
    tokens.sort(key=lambda item: (item[0], -(item[1] - item[0])))
    out: list[str] = []
    cursor = 0
    for start, end, replacement in tokens:
        if start < cursor:
            continue
        out.append(html.escape(text[cursor:start]))
        out.append(replacement)
        cursor = end
    out.append(html.escape(text[cursor:]))
    return "".join(out)


MARKDOWN = MarkdownIt(
    "commonmark",
    {"html": False, "breaks": True, "linkify": False, "typographer": False},
).enable(["table", "strikethrough"])


class LinkRenderedMarkdown(HTMLParser):
    """Post-process safe Markdown HTML and attach recovered local files."""

    def __init__(self, resource_links: dict[str, str], lang: str = DEFAULT_LANG) -> None:
        super().__init__(convert_charrefs=True)
        self.resource_links = resource_links
        self._lang = lang
        self.output: list[str] = []
        self.literal_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        rendered_attrs = []
        mapped_link = False
        unresolved_relative = False
        for name, value in attrs:
            if name == "href" and value:
                decoded = unquote(value)
                local = self.resource_links.get(decoded) or self.resource_links.get(decoded.lstrip("/"))
                if local:
                    value = href(local)
                    mapped_link = True
                elif not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:", decoded) and not decoded.startswith("#"):
                    unresolved_relative = True
            if value is None:
                rendered_attrs.append(f" {name}")
            else:
                rendered_attrs.append(f' {name}="{html.escape(value, quote=True)}"')
        if tag == "a" and mapped_link:
            rendered_attrs.append(' class="file-link"')
        elif tag == "a" and unresolved_relative:
            rendered_attrs.append(f' class="source-reference" data-original-reference="true" title="{html.escape(L("chat.source_reference_title", self._lang))}"')
        self.output.append(f"<{tag}{''.join(rendered_attrs)}>")
        if tag in {"a", "code", "pre"}:
            self.literal_depth += 1

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        rendered_attrs = "".join(
            f' {name}="{html.escape(value, quote=True)}"' if value is not None else f" {name}"
            for name, value in attrs
        )
        self.output.append(f"<{tag}{rendered_attrs}>")

    def handle_endtag(self, tag: str) -> None:
        self.output.append(f"</{tag}>")
        if tag in {"a", "code", "pre"}:
            self.literal_depth = max(0, self.literal_depth - 1)

    def handle_data(self, data: str) -> None:
        if self.literal_depth:
            self.output.append(html.escape(data))
        else:
            self.output.append(linked_text(data, self.resource_links))


def render_markdown(value: Any, resource_links: dict[str, str], lang: str = DEFAULT_LANG) -> str:
    parser = LinkRenderedMarkdown(resource_links, lang)
    parser.feed(MARKDOWN.render(str(value or "")))
    parser.close()
    return "".join(parser.output)


def render_chat_html(
    conversation: dict[str, Any],
    recovered: list[dict[str, Any]],
    presented: list[dict[str, Any]],
    attachment_records: list[dict[str, Any]],
    conv_dir: Path,
    lang: str = DEFAULT_LANG,
) -> str:
    t = lambda key: L(f"chat.{key}", lang)
    title = str(conversation.get("name") or t("untitled"))
    resource_links: dict[str, str] = {}
    basename_links: dict[str, str] = {}
    for item in recovered:
        local = str(item["local_path"])
        original = str(item["original_path"])
        resource_links[original] = local
        resource_links[original.lstrip("/")] = local
        basename_links.setdefault(Path(original).name, local)
    # Markdown files often refer to sibling sources with short relative paths.
    # Add only unambiguous path suffixes to the resolver.
    suffix_candidates: dict[str, set[str]] = {}
    for item in recovered:
        original_parts = PurePosixPath(str(item["original_path"])).parts
        local = str(item["local_path"])
        for length in range(1, min(5, len(original_parts)) + 1):
            suffix = "/".join(original_parts[-length:])
            suffix_candidates.setdefault(suffix, set()).add(local)
            suffix_candidates.setdefault("./" + suffix, set()).add(local)
    for candidate, locals_ in suffix_candidates.items():
        if len(locals_) == 1:
            resource_links.setdefault(candidate, next(iter(locals_)))
    regenerated = sorted((conv_dir / "regenerated").glob("*")) if (conv_dir / "regenerated").exists() else []
    for path in regenerated:
        if path.is_file():
            basename_links[path.name] = path.relative_to(conv_dir).as_posix()
    for item in presented:
        raw_path = str(item.get("file_path") or "")
        if Path(raw_path).name in basename_links:
            resource_links[raw_path] = basename_links[Path(raw_path).name]

    attachments_by_message: dict[str, list[dict[str, Any]]] = {}
    for record in attachment_records:
        attachments_by_message.setdefault(str(record.get("message_uuid")), []).append(record)

    message_times = {
        str(message.get("uuid")): str(message.get("created_at") or "")
        for message in conversation.get("chat_messages") or []
    }
    version_counts: Counter[str] = Counter()
    artifact_rows: list[dict[str, Any]] = []
    for item in presented:
        raw_path = str(item.get("file_path") or item.get("name") or "unknown")
        filename = Path(raw_path).name or raw_path
        version_key = filename.casefold()
        version_counts[version_key] += 1
        local = resource_links.get(raw_path, "")
        suffix = Path(local or filename).suffix.lower()
        artifact_rows.append({
            "name": filename,
            "date": message_times.get(str(item.get("message_uuid")), ""),
            "version": version_counts[version_key],
            "local": local,
            "raw_path": raw_path,
            "message_uuid": str(item.get("message_uuid") or ""),
            "previewable": suffix in {".pdf", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".html", ".htm", ".txt", ".md", ".json", ".csv"},
        })
    listed_locals = {str(item["local"]) for item in artifact_rows if item["local"]}
    messages = conversation.get("chat_messages") or []
    for path in regenerated:
        if not path.is_file():
            continue
        local = path.relative_to(conv_dir).as_posix()
        if local in listed_locals:
            continue
        mention_date = ""
        for message in messages:
            if path.name in json.dumps(message, ensure_ascii=False):
                mention_date = str(message.get("created_at") or mention_date)
        version_key = path.name.casefold()
        version_counts[version_key] += 1
        artifact_rows.append({
            "name": path.name,
            "date": mention_date or str(conversation.get("created_at") or ""),
            "version": version_counts[version_key],
            "local": local,
            "raw_path": "",
            "message_uuid": "",
            "previewable": path.suffix.lower() in {".pdf", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".html", ".htm", ".txt", ".md", ".json", ".csv"},
        })

    def artifact_dom_key(name: str) -> str:
        return hashlib.sha1(name.casefold().encode("utf-8")).hexdigest()[:12]

    def artifact_icon(name: str) -> str:
        suffix = Path(name).suffix.lower()
        icon_class = "presentation" if suffix in {".ppt", ".pptx"} else "sheet" if suffix in {".xls", ".xlsx", ".csv"} else "document"
        if icon_class == "presentation":
            drawing = '<rect x="4" y="4" width="16" height="11" rx="2"/><path d="M12 15v5m-4 0h8M8 8h8"/>'
        elif icon_class == "sheet":
            drawing = '<rect x="5" y="3" width="14" height="18" rx="2"/><path d="M5 9h14M9.5 9v12M14.5 9v12M5 15h14"/>'
        else:
            drawing = '<path d="M6 3h8l4 4v14H6z"/><path d="M14 3v5h5M9 13h6M9 17h6"/>'
        return f'<span class="artifact-file-icon {icon_class}"><svg viewBox="0 0 24 24" aria-hidden="true">{drawing}</svg></span>'

    def markdown_payload(local: str, name: str) -> str:
        if Path(name).suffix.lower() not in {".md", ".markdown"}:
            return ""
        try:
            md_content = (conv_dir / local).read_text(encoding="utf-8", errors="replace")
        except Exception:
            md_content = ""
        md_json = (
            json.dumps(md_content, ensure_ascii=False)
            .replace("&", "\\u0026")
            .replace("<", "\\u003c")
            .replace(">", "\\u003e")
            .replace("\u2028", "\\u2028")
            .replace("\u2029", "\\u2029")
        )
        return f'<script type="application/json" class="md-data" data-src="{href(local)}">{md_json}</script>'

    def json_text_payload(value: str, css_class: str) -> str:
        payload = (
            json.dumps(value, ensure_ascii=False)
            .replace("&", "\\u0026")
            .replace("<", "\\u003c")
            .replace(">", "\\u003e")
            .replace("\u2028", "\\u2028")
            .replace("\u2029", "\\u2029")
        )
        return f'<script type="application/json" class="{css_class}">{payload}</script>'

    def visible_size(size: int) -> str:
        if size < 1024:
            return f"{size} B"
        if size < 1024 * 1024:
            return f"{size / 1024:.2f}".rstrip("0").rstrip(".") + " KB"
        return f"{size / (1024 * 1024):.2f}".rstrip("0").rstrip(".") + " MB"

    def attachment_card(item: dict[str, Any], date: str) -> str:
        local = str(item["local_path"])
        original_name = str(item.get("name") or "")
        pasted = not original_name
        title_text = t("pasted_content") if pasted else original_name
        file_type = str(item.get("file_type") or Path(original_name).suffix.lstrip(".") or "txt").upper()
        badge = t("pasted_badge") if pasted else file_type
        content = str(item.get("content") or "")
        snippet = content.strip().replace("\r\n", "\n").replace("\r", "\n")[:420]
        line_count = int(item.get("lines") or len(content.splitlines()))
        size = int(item.get("bytes") or 0)
        meta = f'{visible_size(size)} · {line_count} {t("lines_label")}'
        visible_date = date.replace("T", " ")[:16] if date else ""
        return "".join([
            '<div class="input-attachment-wrap">',
            f'<button type="button" class="input-attachment preview-button" data-text="1" data-src="{href(local)}" data-title="{html.escape(title_text, quote=True)}" data-meta="{html.escape(meta, quote=True)}" aria-label="{html.escape(t("button_preview"), quote=True)}">',
            f'<span class="input-attachment-snippet">{html.escape(snippet)}</span>',
            f'<span class="input-attachment-badge">{html.escape(badge)}</span>',
            '</button>',
            '<div class="input-attachment-footer">',
            f'<span>{html.escape(visible_date)}</span><span class="input-attachment-actions">',
            f'<button type="button" class="attachment-copy" title="{html.escape(t("button_copy_attachment"), quote=True)}" aria-label="{html.escape(t("button_copy_attachment"), quote=True)}"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="8" y="8" width="11" height="12" rx="2"/><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h2"/></svg></button>',
            f'<a href="{href(local)}" download title="{html.escape(t("button_download_attachment"), quote=True)}" aria-label="{html.escape(t("button_download_attachment"), quote=True)}"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12m-5-5 5 5 5-5M5 20h14"/></svg></a>',
            '</span></div>',
            json_text_payload(content, "attachment-data"),
            '</div>',
        ])

    def artifact_card(item: dict[str, Any], *, inline: bool = False) -> str:
        name = str(item["name"])
        local = str(item["local"])
        date = str(item["date"])
        visible_date = date.replace("T", " ")[:16] if date else t("date_unknown")
        key = artifact_dom_key(name)
        classes = "artifact-card inline-artifact" if inline else "artifact-card"
        bits = [
            f'<article class="{classes}" data-artifact-key="{key}" data-date="{html.escape(date, quote=True)}" data-name="{html.escape(name.casefold(), quote=True)}" data-version="{item["version"]}">',
            artifact_icon(name),
            '<div class="artifact-card-copy">',
            f'<div class="artifact-name">{html.escape(name)}</div>',
            f'<div class="artifact-meta">{html.escape(Path(name).suffix.lstrip(".").upper() or "FILE")} · {html.escape(visible_date)} · {t("version_label")} {item["version"]}</div>',
            '</div>',
        ]
        if local:
            is_md = Path(name).suffix.lower() in {".md", ".markdown"}
            bits.append('<div class="artifact-hover-actions">')
            if item["previewable"]:
                preview_data = ' data-md="1"' if is_md else f' data-src="{href(local)}"'
                bits.append(f'<button type="button" class="artifact-icon-button preview-button"{preview_data} data-title="{html.escape(name, quote=True)}" title="{t("button_preview")}" aria-label="{t("button_preview")}"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M2.5 12s3.5-6 9.5-6 9.5 6 9.5 6-3.5 6-9.5 6-9.5-6-9.5-6Z"/><circle cx="12" cy="12" r="2.5"/></svg></button>')
            bits.append(f'<a class="artifact-icon-button" href="{href(local)}" download title="{t("button_download")}" aria-label="{t("button_download")}"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12m-5-5 5 5 5-5M5 20h14"/></svg></a>')
            if inline:
                bits.append(f'<button type="button" class="artifact-icon-button locate-artifact" data-artifact-key="{key}" title="{t("button_show_in_list")}" aria-label="{t("button_show_in_list")}"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="5" cy="12" r="1"/><circle cx="12" cy="12" r="1"/><circle cx="19" cy="12" r="1"/></svg></button>')
            else:
                bits.append(f'<a class="artifact-icon-button" href="{href(local)}" target="_blank" title="{t("button_open")}" aria-label="{t("button_open")}"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 5h5v5M19 5l-8 8M18 13v6H5V6h6"/></svg></a>')
            bits.append('</div>')
            bits.append(markdown_payload(local, name))
        else:
            bits.append(f'<div class="artifact-missing" title="">⚠ {t("metadata_missing_warning")}</div>')
        bits.append('</article>')
        return "".join(bits)

    copy_title = html.escape(t("button_copy_chat"), quote=True)
    parts = [f"""<!doctype html>
<html lang="{L("html_lang", lang)}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<script>try{{document.documentElement.dataset.theme=localStorage.getItem('restore_theme')||'dark'}}catch(error){{document.documentElement.dataset.theme='dark'}}</script>
<style>
:root{{color-scheme:light;--bg:#f7f6f2;--surface:#fff;--surface-2:#f0eee8;--surface-3:#e8e5de;--ink:#292724;--muted:#77736c;--line:#dedbd2;--line-strong:#c9c5bb;--accent:#d97757;--accent-soft:#f6e8e1;--link:#356ea3;--user:#e9e7e1;--code:#252523;--code-ink:#f3f1eb;--shadow:0 1px 2px #1817140d}}
:root[data-theme="dark"]{{color-scheme:dark;--bg:#181817;--surface:#20201f;--surface-2:#292927;--surface-3:#343330;--ink:#eeeDE8;--muted:#a6a39c;--line:#363532;--line-strong:#4a4843;--accent:#d97757;--accent-soft:#3c2922;--link:#76a9dc;--user:#2b2a28;--code:#111110;--code-ink:#eceae4;--shadow:none}}
*{{box-sizing:border-box}}html{{background:var(--bg)}}body{{margin:0;background:var(--bg);color:var(--ink);font:14px/1.58 ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;transition:background .18s ease,color .18s ease}}
a{{color:var(--link)}}.layout{{display:grid;grid-template-columns:minmax(0,1fr) 8px var(--aside-w,340px);gap:0;height:100vh;overflow:hidden}}
main{{min-width:0;padding:0 clamp(18px,3vw,46px) 70px;overflow-y:auto;overflow-x:hidden}}.hero,.message,.panel{{background:transparent;border:0;box-shadow:none}}
.hero{{position:sticky;top:0;z-index:10;margin:0 calc(clamp(18px,3vw,46px)*-1) 24px;padding:8px clamp(18px,3vw,46px);border-bottom:1px solid var(--line);background:color-mix(in srgb,var(--bg) 94%,transparent);backdrop-filter:blur(14px)}}.hero-row{{display:flex;align-items:center;gap:12px;min-height:34px}}.hero h1{{min-width:0;flex:1;margin:0;font-size:14px;line-height:1.4;font-weight:650;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}.hero-copy{{position:relative;flex:0 0 34px;box-shadow:none!important}}.hero-copy .check-icon{{display:none}}.hero-copy.copied{{color:#68a27b!important}}.hero-copy.copied .copy-icon{{display:none}}.hero-copy.copied .check-icon{{display:block}}.meta{{max-height:0;opacity:0;overflow:hidden;color:var(--muted);font-size:10px;transition:max-height .18s ease,opacity .18s ease}}
.hero .chips{{margin:0;gap:5px;max-height:0;opacity:0;overflow:hidden;transition:max-height .18s ease,opacity .18s ease}}.hero:hover .meta,.hero:focus-within .meta{{max-height:34px;opacity:1;margin-top:5px}}.hero:hover .chips,.hero:focus-within .chips{{max-height:40px;opacity:1;margin-top:6px}}.conv-summary{{margin:6px 0 0;border:1px solid color-mix(in srgb,var(--accent) 30%,var(--line));border-radius:8px;background:color-mix(in srgb,var(--accent) 6%,var(--surface));overflow:hidden}}.conv-summary>summary{{cursor:pointer;padding:5px 10px;font-size:11px;font-weight:500;color:color-mix(in srgb,var(--accent) 70%,var(--ink));list-style:none}}.conv-summary>summary::-webkit-details-marker{{display:none}}.conv-summary>summary::before{{content:"▸";display:inline-block;margin-right:5px;transition:transform .15s ease;font-size:10px}}.conv-summary[open]>summary::before{{transform:rotate(90deg)}}.conv-summary>div{{padding:8px 12px 12px;border-top:1px solid color-mix(in srgb,var(--accent) 20%,var(--line));max-height:300px;overflow:auto;font-size:12px;line-height:1.5}}.conv-summary .markdown>{{font-size:12px;line-height:1.5}}.conv-summary .markdown h1,.conv-summary .markdown h2,.conv-summary .markdown h3{{font-size:13px;margin:8px 0 4px}}
.message{{max-width:860px;margin:0 auto 26px;padding:0}}.message.user{{width:fit-content;max-width:min(72%,700px);margin:12px 0 36px auto;padding:12px 16px;border-radius:13px;background:var(--user)}}.message.user.has-attachments{{width:min(72%,700px);padding:0;background:transparent}}.message.user.has-attachments .user-bubble{{width:fit-content;margin-left:auto;padding:12px 16px;border-radius:13px;background:var(--user)}}
.message.assistant{{background:transparent}}.message.user .head{{margin:0}}.message.user .role{{display:none}}.message.user .time{{display:none}}
.head{{display:flex;justify-content:space-between;gap:12px;margin-bottom:12px}}.role{{font-weight:750}}.time{{color:var(--muted);font-size:12px}}.text{{white-space:pre-wrap;overflow-wrap:anywhere}}.markdown{{overflow-wrap:anywhere}}.markdown>:first-child{{margin-top:0}}.markdown>:last-child{{margin-bottom:0}}.markdown h1,.markdown h2,.markdown h3{{line-height:1.25;margin:1.1em 0 .45em}}.markdown h1{{font-size:1.55em}}.markdown h2{{font-size:1.3em}}.markdown h3{{font-size:1.12em}}.markdown blockquote{{margin:10px 0;padding:3px 14px;border-left:4px solid #b9b4aa;color:#555;background:#f7f6f3}}.markdown table{{border-collapse:collapse;display:block;max-width:100%;overflow:auto}}.markdown th,.markdown td{{border:1px solid var(--line);padding:6px 10px}}.markdown th{{background:#f1efe9}}.markdown img{{max-width:100%}}.markdown ul,.markdown ol{{padding-left:25px}}.markdown pre{{white-space:pre-wrap;overflow-wrap:anywhere;background:var(--code);color:#e8eaed;font:12px/1.45 ui-monospace,SFMono-Regular,Consolas,monospace;padding:10px 12px;border-radius:8px;margin:8px 0}}.markdown code{{font:13px/1.4 ui-monospace,SFMono-Regular,Consolas,monospace;background:#f0eee8;padding:1px 4px;border-radius:3px}}.markdown pre code{{background:none;padding:0}}
details{{border:1px solid var(--line);border-radius:10px;margin:7px 0;background:var(--surface)}}summary{{cursor:pointer;padding:7px 10px;font-weight:400;font-size:12px;color:var(--muted)}}details[open]>summary{{border-bottom:1px solid var(--line)}}details>div,details>pre{{margin:0;padding:10px;border-top:1px solid var(--line)}}
details.tool-call{{background:transparent;border-color:var(--line)}}details.tool-call>summary{{font-size:11px;color:var(--muted)}}details.tool-call>pre{{font-size:11px}}
.tool-log{{max-width:860px;margin:4px auto 24px;border:0;background:transparent}}.tool-log>summary{{display:flex;align-items:center;gap:7px;padding:4px 0;list-style:none;font-size:13px;color:var(--muted)}}.tool-log>summary::-webkit-details-marker{{display:none}}.tool-log>summary::after{{content:"›";font-size:17px;line-height:1;transition:transform .15s ease}}.tool-log[open]>summary::after{{transform:rotate(90deg)}}.tool-log[open]{{border:1px solid var(--line);border-radius:12px;background:color-mix(in srgb,var(--surface) 45%,transparent)}}.tool-log[open]>summary{{padding:9px 12px;border-bottom:1px solid var(--line)}}.tool-log-body{{padding:0!important;border:0!important}}.tool-step{{margin:0;border:0;border-bottom:1px solid var(--line);border-radius:0;background:transparent}}.tool-step:last-child{{border-bottom:0}}.tool-step>summary{{padding:10px 12px;font-size:12px}}.tool-step[open]>summary{{color:var(--ink)}}.tool-step>pre,.tool-step>.tool-result-body{{max-height:340px;overflow:auto;border-top:1px solid var(--line)}}
pre{{white-space:pre-wrap;overflow-wrap:anywhere;background:var(--code);color:var(--code-ink);font:12px/1.45 ui-monospace,SFMono-Regular,Consolas,monospace}}
.markdown blockquote{{color:var(--muted);background:var(--surface-2);border-left-color:var(--line-strong)}}.markdown th{{background:var(--surface-2)}}.markdown pre{{color:var(--code-ink)}}.markdown code{{background:var(--surface-2)}}
.chips{{display:flex;flex-wrap:wrap;gap:7px;margin:8px 0}}.chip,.file-link{{display:inline-block;border:1px solid var(--line);border-radius:999px;padding:3px 9px;background:var(--surface);color:var(--muted);text-decoration:none;font-size:11px}}.chip:hover,.file-link:hover{{color:var(--ink);border-color:var(--line-strong)}}
.resource{{border:1px solid var(--line);border-radius:10px;padding:9px 11px;margin:8px 0;background:var(--surface)}}.missing{{border-color:#8d713d;background:color-mix(in srgb,#9a7429 12%,var(--surface))}}.source-reference{{color:var(--muted);text-decoration-style:dotted}}
.input-attachment-wrap{{width:132px;margin:2px 0 10px auto}}.input-attachment{{display:flex;width:132px;height:132px;flex-direction:column;justify-content:space-between;gap:7px;padding:11px;border:1px solid var(--line-strong);border-radius:11px;background:color-mix(in srgb,var(--surface) 72%,transparent);color:var(--ink);text-align:left;cursor:pointer;transition:border-color .13s ease,background .13s ease}}.input-attachment:hover,.input-attachment:focus-visible{{border-color:var(--muted);background:var(--surface)}}.input-attachment-snippet{{display:-webkit-box;min-height:0;overflow:hidden;-webkit-box-orient:vertical;-webkit-line-clamp:5;white-space:pre-wrap;overflow-wrap:anywhere;font:11px/1.32 ui-monospace,SFMono-Regular,Consolas,monospace;color:var(--ink)}}.input-attachment-badge{{align-self:flex-start;border-radius:999px;padding:2px 7px;background:var(--surface-2);color:var(--muted);font-size:10px;line-height:1.5}}.input-attachment-footer{{display:flex;align-items:center;justify-content:flex-end;gap:7px;padding:5px 2px 0;color:var(--muted);font-size:11px;white-space:nowrap}}.input-attachment-actions{{display:flex;align-items:center;gap:2px}}.input-attachment-actions button,.input-attachment-actions a{{display:grid;width:24px;height:24px;place-items:center;border:0;border-radius:6px;padding:0;background:transparent;color:var(--muted);cursor:pointer}}.input-attachment-actions button:hover,.input-attachment-actions a:hover{{background:var(--surface-2);color:var(--ink)}}.input-attachment-actions svg{{width:15px;height:15px;fill:none;stroke:currentColor;stroke-width:1.7;stroke-linecap:round;stroke-linejoin:round}}
.right-resizer{{position:sticky;top:0;height:100vh;cursor:col-resize;touch-action:none}}.right-resizer::after{{content:"";position:absolute;left:3px;top:12px;bottom:12px;width:1px;border-radius:2px;background:var(--line);opacity:.8}}.right-resizer:hover::after,.right-resizer.dragging::after{{width:2px;background:var(--accent);opacity:1}}body.resizing{{cursor:col-resize;user-select:none}}body.resizing-preview{{cursor:nwse-resize;user-select:none}}
aside{{height:100vh;overflow:auto;min-width:0;border-left:1px solid var(--line);background:var(--bg)}}.panel{{padding:13px 14px 30px}}.panel h2{{font-size:13px;margin:16px 0 8px}}.panel h2:first-child{{margin-top:0}}.panel ul{{padding-left:20px;margin:5px 0}}.panel li{{margin:5px 0;overflow-wrap:anywhere}}.count{{color:var(--muted);font-size:11px}}
.artifact-toolbar{{display:flex;gap:7px;align-items:center;margin:8px 0 11px}}.artifact-toolbar select{{min-width:0;flex:1;border:1px solid var(--line);border-radius:8px;background:var(--surface);color:var(--ink);padding:6px}}.artifact-list,.file-list{{display:grid;gap:9px}}.artifact-card,.file-card{{position:relative;border:1px solid var(--line);border-radius:12px;padding:11px;background:var(--surface);box-shadow:var(--shadow)}}.artifact-card{{display:grid;grid-template-columns:38px minmax(0,1fr);gap:10px;align-items:center;min-height:62px}}.inline-artifact{{max-width:568px;margin:14px 0 22px;background:var(--surface)}}.artifact-card:hover,.artifact-card:focus-within,.file-card:hover{{border-color:var(--line-strong)}}.artifact-card-copy{{min-width:0}}.artifact-file-icon{{width:36px;height:36px;display:grid;place-items:center;border-radius:7px;background:var(--surface-2);color:var(--muted)}}.artifact-file-icon.presentation{{color:#e47d4f;background:color-mix(in srgb,#c76136 18%,var(--surface))}}.artifact-file-icon.sheet{{color:#68a27b;background:color-mix(in srgb,#43865a 18%,var(--surface))}}.artifact-file-icon svg,.artifact-icon-button svg{{width:19px;height:19px;fill:none;stroke:currentColor;stroke-width:1.6;stroke-linecap:round;stroke-linejoin:round}}.artifact-name,.file-name{{font-weight:620;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}.artifact-meta,.file-meta{{color:var(--muted);font-size:11px;margin-top:3px}}.artifact-hover-actions{{position:absolute;z-index:3;right:9px;top:-17px;display:flex;gap:5px;opacity:0;transform:translateY(3px);pointer-events:none;transition:opacity .13s ease,transform .13s ease}}.artifact-card:hover .artifact-hover-actions,.artifact-card:focus-within .artifact-hover-actions{{opacity:1;transform:none;pointer-events:auto}}.artifact-icon-button{{width:34px;height:34px;display:grid;place-items:center;border:1px solid var(--line)!important;border-radius:10px!important;padding:0!important;background:var(--surface-2)!important;color:var(--ink)!important;text-decoration:none;cursor:pointer;box-shadow:0 2px 8px #0003}}.artifact-icon-button:hover{{background:var(--surface-3)!important;border-color:var(--line-strong)!important}}.locate-artifact circle{{fill:currentColor;stroke:none}}.file-card a{{display:inline-block;margin-top:7px;border:1px solid var(--line);border-radius:7px;padding:4px 8px;background:var(--surface-2);color:var(--ink);font-size:11px;text-decoration:none}}.artifact-missing{{grid-column:1/-1;color:#b8954f;font-size:11px}}.artifact-card.located{{animation:artifact-pulse 1.2s ease}}@keyframes artifact-pulse{{0%,100%{{box-shadow:var(--shadow)}}35%{{box-shadow:0 0 0 3px color-mix(in srgb,var(--accent) 55%,transparent)}}}}.preview-box{{display:none;grid-column:1/-1;margin-top:2px;padding-top:10px;border-top:1px solid var(--line)}}.preview-box.open{{display:block}}.preview-head{{display:flex;justify-content:space-between;gap:8px;align-items:center;margin-bottom:10px}}.preview-title{{overflow-wrap:anywhere}}.preview-resizer{{position:relative;height:55vh;min-height:120px;max-height:90vh;overflow:hidden;border:1px solid var(--line);border-radius:8px;background:var(--surface)}}.preview-box iframe{{width:100%;height:100%;border:0;background:white}}.preview-box .md-preview{{height:100%;overflow:auto;padding:14px 16px;background:var(--surface)}}.resize-handle{{position:absolute;right:0;bottom:0;width:18px;height:18px;cursor:nwse-resize;z-index:10;background:linear-gradient(135deg,transparent 50%,var(--line-strong) 50%);border-radius:0 0 8px 0;opacity:.6}}.resize-handle:hover{{opacity:1;background:linear-gradient(135deg,transparent 50%,var(--accent) 50%)}}.preview-close{{border:0;background:transparent;color:var(--muted);font-size:22px;cursor:pointer}}
.preview-box.standalone{{margin:0;padding:0;border:0}}.preview-meta{{display:none;color:var(--muted);font-size:11px;margin:0 0 10px}}.preview-meta:not(:empty){{display:block}}.preview-box.standalone .preview-resizer{{height:calc(100vh - 78px);max-height:none}}.preview-box .md-preview.plain{{padding:0}}.plain-text-preview{{min-height:100%;margin:0;padding:18px!important;border-radius:0;background:var(--surface);color:var(--ink);white-space:pre-wrap;overflow-wrap:anywhere;font:12px/1.52 ui-monospace,SFMono-Regular,Consolas,monospace}}.panel.attachment-preview-open{{height:100vh;padding-bottom:10px;overflow:hidden}}.panel.attachment-preview-open>:not(.preview-box){{display:none}}
.lang-switch{{position:absolute;top:10px;right:18px;border:1px solid var(--line);border-radius:8px;background:var(--surface);color:var(--ink);padding:4px 6px;font:inherit;font-size:11px}}
@media(hover:none){{.artifact-hover-actions{{position:static;grid-column:2;grid-row:2;justify-self:end;opacity:1;transform:none;pointer-events:auto;margin-top:5px}}}}@media(max-width:900px){{.layout{{grid-template-columns:1fr;height:auto;overflow:auto}}main{{padding:0 14px 50px;overflow:visible}}.hero{{margin:0 -14px 20px;padding:11px 14px}}.right-resizer{{display:none}}aside{{height:auto;order:-1;border:0;border-bottom:1px solid var(--line)}}.message.user{{max-width:88%}}.lang-switch{{right:10px}}}}
.summary-card{{position:relative}}
.summary-card .conv-summary>summary{{min-height:34px;padding:8px 108px 8px 10px}}
.summary-card>.artifact-hover-actions{{top:3px;right:5px}}
.summary-card .artifact-icon-button{{width:28px;height:28px}}
.summary-card .artifact-icon-button svg{{width:16px;height:16px}}
.summary-card:hover .artifact-hover-actions,.summary-card:focus-within .artifact-hover-actions{{opacity:1;transform:none;pointer-events:auto}}
#summary-copy.copied{{color:#68a27b!important}}
.message-slot{{display:flow-root;min-height:1px;scroll-margin-top:80px}}
@media(hover:none){{.summary-card>.artifact-hover-actions{{position:absolute;margin-top:0;justify-content:flex-end}}}}
</style><title>""", html.escape(title), "</title></head><body><div class=\"layout\"><main>"]
    parts.extend([
        '<section class="hero"><div class="hero-row">', f"<h1>{html.escape(title)}</h1>",
        f'<button type="button" class="artifact-icon-button hero-copy" id="copy-chat" title="{copy_title}" aria-label="{copy_title}"><svg class="copy-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="8" y="8" width="11" height="12" rx="2"/><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h2"/></svg><svg class="check-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m5 12 4 4L19 6"/></svg></button></div>',
    ])
    # Conversation summary — collapsible highlighted block
    summary_text = str(conversation.get("summary") or "").strip()
    if summary_text:
        summary_url = "data:text/markdown;charset=utf-8," + quote(summary_text + "\n", safe="")
        preview_label = html.escape(t("button_preview"), quote=True)
        download_label = html.escape(t("button_download"), quote=True)
        copy_label = html.escape(t("button_copy_attachment"), quote=True)
        parts.extend([
            '<div class="summary-card">',
            f'<details class="conv-summary" id="conv-summary"><summary>📖 {t("summary_label")}</summary><div class="markdown">{render_markdown(summary_text, resource_links, lang)}</div></details>',
            '<div class="artifact-hover-actions">',
            f'<button type="button" class="artifact-icon-button" id="summary-preview" aria-controls="conv-summary" aria-expanded="false" title="{preview_label}" aria-label="{preview_label}"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M2.5 12s3.5-6 9.5-6 9.5 6 9.5 6-3.5 6-9.5 6-9.5-6-9.5-6Z"/><circle cx="12" cy="12" r="2.5"/></svg></button>',
            f'<a class="artifact-icon-button" id="summary-download" href="{summary_url}" download="summary.md" title="{download_label}" aria-label="{download_label}"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12m-5-5 5 5 5-5M5 20h14"/></svg></a>',
            f'<button type="button" class="artifact-icon-button" id="summary-copy" title="{copy_label}" aria-label="{copy_label}"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="8" y="8" width="11" height="12" rx="2"/><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h2"/></svg></button>',
            '</div></div>',
        ])
    parts.append(
        f'<div class="meta">{t("uuid_label")}: <code>{html.escape(str(conversation.get("uuid") or ""))}</code><br>{t("created_label")}: {html.escape(str(conversation.get("created_at") or ""))}</div>'
    )
    parts.append(
        f'<div class="chips"><a class="chip" href="README.md">{t("link_readme")}</a><a class="chip" href="chat.md">{t("link_markdown")}</a><a class="chip" href="conversation.json">{t("link_json")}</a><a class="chip" href="recovered_files.json">{t("link_files_json")}</a><a class="chip" href="presented_files.json">{t("link_artifacts_json")}</a><a class="chip" href="../../raw/artifacts/">{t("link_all_artifacts")}</a></div>'
    )
    parts.append("</section>")
    command_index = 0

    def presented_paths(block: dict[str, Any]) -> list[str]:
        kind = block.get("type")
        if kind == "tool_use" and str(block.get("name") or "") in PRESENT_TOOLS:
            inp = block.get("input") or {}
            if not isinstance(inp, dict):
                return []
            paths = [str(path) for path in (inp.get("filepaths") or inp.get("files") or []) if path]
            if inp.get("file_path"):
                paths.append(str(inp["file_path"]))
            return paths
        if kind == "tool_result":
            return [
                str(item.get("file_path") or "")
                for item in block.get("content") or []
                if isinstance(item, dict) and item.get("type") == "local_resource" and item.get("file_path")
            ]
        return []

    def matching_artifact(raw_path: str, message_uuid: str) -> dict[str, Any]:
        exact = [item for item in artifact_rows if item.get("raw_path") == raw_path and item.get("message_uuid") == message_uuid]
        if not exact:
            exact = [item for item in artifact_rows if item.get("raw_path") == raw_path]
        if exact:
            return max(exact, key=lambda item: int(item["version"]))
        name = Path(raw_path).name or raw_path
        local = resource_links.get(raw_path, "")
        suffix = Path(local or name).suffix.lower()
        return {
            "name": name,
            "date": message_times.get(message_uuid, ""),
            "version": 1,
            "local": local,
            "raw_path": raw_path,
            "message_uuid": message_uuid,
            "previewable": suffix in {".pdf", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".html", ".htm", ".txt", ".md", ".json", ".csv"},
        }

    def copy_chat_markdown() -> str:
        out = [f"# {title}"]
        summary_text = str(conversation.get("summary") or "").strip()
        if summary_text:
            out.extend(["", "## Summary", "", summary_text])
        for message in conversation.get("chat_messages") or []:
            message_parts: list[str] = []
            inline_seen: set[str] = set()
            for attachment in attachments_by_message.get(str(message.get("uuid")), []):
                attachment_name = str(attachment.get("name") or t("pasted_content"))
                message_parts.append(
                    f'📎 **{t("attachment_label")}:** '
                    f'[{attachment_name}]({href(str(attachment["local_path"]))})'
                )
            blocks = message.get("content") or []
            if not blocks and message.get("text"):
                message_parts.append(str(message["text"]).strip())
            for block in blocks:
                if not isinstance(block, dict):
                    continue
                if block.get("type") == "text":
                    value = str(block.get("text") or "").strip()
                    if value:
                        message_parts.append(value)
                    continue
                for raw_path in presented_paths(block):
                    identity = raw_path.casefold()
                    if identity in inline_seen:
                        continue
                    inline_seen.add(identity)
                    item = matching_artifact(raw_path, str(message.get("uuid") or ""))
                    name = str(item["name"]).replace("[", "\\[").replace("]", "\\]")
                    artifact_type = Path(str(item["name"])).suffix.lstrip(".").upper() or "FILE"
                    details = f'{artifact_type} · {t("version_label")} {item["version"]}'
                    if item.get("date"):
                        details += f' · {str(item["date"]).replace("T", " ")[:16]}'
                    if item.get("local"):
                        message_parts.append(f'📎 **{t("copy_artifact_label")}:** [{name}]({href(str(item["local"]))}) · {details}')
                    else:
                        message_parts.append(f'📎 **{t("copy_artifact_label")}:** {name} · {details} · {t("metadata_only")}')
            message_parts = [part for part in message_parts if part]
            if not message_parts:
                continue
            role = t("role_user") if message.get("sender") == "human" else t("role_assistant")
            out.extend(["", f"## {role}", "", "\n\n".join(message_parts)])
        return "\n".join(out).rstrip() + "\n"

    copy_payload = (
        json.dumps(copy_chat_markdown(), ensure_ascii=False)
        .replace("&", "\\u0026")
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("\u2028", "\\u2028")
        .replace("\u2029", "\\u2029")
    )
    parts.append(f'<script type="application/json" id="copy-chat-data">{copy_payload}</script>')

    def tool_step(block: dict[str, Any]) -> str:
        nonlocal command_index
        kind = block.get("type")
        if kind == "thinking":
            value = str(block.get("thinking") or "").strip()
            if not value:
                return ""
            return f'<details class="tool-step"><summary>{t("thinking")}</summary><div class="tool-result-body markdown">{render_markdown(value, resource_links, lang)}</div></details>'
        if kind == "tool_use":
            name = str(block.get("name") or "tool")
            inp = block.get("input") or {}
            if not isinstance(inp, dict):
                inp = {}
            chips: list[str] = []
            if name in COMMAND_TOOLS:
                command_index += 1
                chips.append(f'<a class="chip" href="commands/{command_index:04d}_{safe_component(name, 30)}.sh">{t("command_label")} #{command_index}</a>')
            candidate_paths = [inp.get("file_path"), inp.get("path"), *(inp.get("filepaths") or [])]
            linked_names: list[str] = []
            for raw_path in candidate_paths:
                if not raw_path:
                    continue
                linked_names.append(Path(str(raw_path)).name)
                if str(raw_path) in resource_links:
                    chips.append(f'<a class="chip" href="{href(resource_links[str(raw_path)])}">📄 {html.escape(Path(str(raw_path)).name)}</a>')
            lowered = name.casefold()
            if name in COMMAND_TOOLS:
                label = t("tool_command")
            elif "read" in lowered:
                label = t("tool_read")
            elif any(token in lowered for token in ("write", "edit", "create", "replace")):
                label = t("tool_change")
            else:
                label = t("tool_action")
            suffix = f' · {html.escape(linked_names[0])}' if linked_names else f' · {html.escape(name)}'
            payload = html.escape(json.dumps(inp, ensure_ascii=False, indent=2))
            return f'<details class="tool-step"><summary>{html.escape(label)}{suffix}</summary><div class="chips">{"".join(chips)}</div><pre>{payload}</pre></details>'
        if kind == "tool_result":
            texts = [
                render_markdown(item.get("text"), resource_links, lang)
                for item in block.get("content") or []
                if isinstance(item, dict) and item.get("type") == "text"
            ]
            if not texts:
                return ""
            name = str(block.get("name") or "tool")
            return f'<details class="tool-step"><summary>{t("result_prefix")} · {html.escape(name)}</summary><div class="tool-result-body markdown">{"".join(texts)}</div></details>'
        return ""

    messages = conversation.get("chat_messages") or []
    virtual_messages = []
    for msg_index, message in enumerate(messages, 1):
        message_start = len(parts)
        sender = message.get("sender")
        role = t("role_user") if sender == "human" else t("role_assistant")
        css_role = "user" if sender == "human" else "assistant"
        message_attachments = attachments_by_message.get(str(message.get("uuid")), [])
        attachment_class = " has-attachments" if sender == "human" and message_attachments else ""
        has_user_text = bool(str(message.get("text") or "").strip()) or any(
            isinstance(block, dict) and block.get("type") == "text" and str(block.get("text") or "").strip()
            for block in message.get("content") or []
        )
        wrap_user_bubble = bool(attachment_class and has_user_text)
        parts.append(f'<article class="message {css_role}{attachment_class}" id="message-{msg_index}"><div class="head"><span class="role">{role}</span><span class="time">{html.escape(str(message.get("created_at") or ""))}</span></div>')
        if wrap_user_bubble:
            parts.append('<div class="user-bubble">')
        blocks = message.get("content") or []
        # Claude exports the same user upload twice: once in top-level
        # `attachments` (where extracted text is available) and once as a
        # UUID-only `document`/`image` content block.  A recovered attachment
        # therefore satisfies one such placeholder in the same message.
        covered_binary_placeholders = len(message_attachments)
        if not blocks and message.get("text"):
            parts.append(f'<div class="markdown">{render_markdown(message["text"], resource_links, lang)}</div>')
        block_index = 0
        inline_seen: set[str] = set()
        while block_index < len(blocks):
            block = blocks[block_index]
            block_index += 1
            if not isinstance(block, dict):
                continue
            kind = block.get("type")
            if kind == "text":
                parts.append(f'<div class="markdown">{render_markdown(block.get("text"), resource_links, lang)}</div>')
                continue
            paths = presented_paths(block)
            if paths:
                for raw_path in paths:
                    identity = raw_path.casefold()
                    if identity in inline_seen:
                        continue
                    inline_seen.add(identity)
                    parts.append(artifact_card(matching_artifact(raw_path, str(message.get("uuid") or "")), inline=True))
                continue
            if kind in {"thinking", "tool_use", "tool_result"}:
                group = [block]
                while block_index < len(blocks):
                    candidate = blocks[block_index]
                    if not isinstance(candidate, dict) or candidate.get("type") not in {"thinking", "tool_use", "tool_result"} or presented_paths(candidate):
                        break
                    group.append(candidate)
                    block_index += 1
                steps = [value for item in group if (value := tool_step(item))]
                if steps:
                    summary = html.escape(t("tool_log_summary").format(count=len(steps)))
                    parts.append(f'<details class="tool-log"><summary>{summary}</summary><div class="tool-log-body">{"".join(steps)}</div></details>')
            elif kind in {"image", "document"}:
                if covered_binary_placeholders:
                    covered_binary_placeholders -= 1
                else:
                    parts.append(f'<div class="resource missing">{html.escape(kind.title())}: UUID <code>{html.escape(str(block.get("file_uuid") or ""))}</code> — {t("binary_missing")}</div>')
        if wrap_user_bubble:
            parts.append('</div>')
        for attachment in message_attachments:
            parts.append(attachment_card(attachment, str(message.get("created_at") or "")))
        parts.append("</article>")
        if len(messages) > 40:
            markup = "".join(parts[message_start:])
            del parts[message_start:]
            text_size = sum(len(str(block.get("text") or "")) for block in blocks
                            if isinstance(block, dict) and block.get("type") == "text")
            if not blocks:
                text_size = len(str(message.get("text") or ""))
            virtual_messages.append([markup, min(2000, 90 + (text_size // 75) * 23)])
    if virtual_messages:
        # Escaping '<' keeps even literal script end tags inside the JSON.
        payload = json.dumps(virtual_messages, ensure_ascii=False).replace("<", "\\u003c")
        parts.append(f'<div id="message-window"></div><script type="application/json" id="message-data">{payload}</script>')
    parts.append(f'</main><div class="right-resizer" id="right-resizer" role="separator" aria-orientation="vertical" aria-label="{html.escape(t("artifacts_title"), quote=True)}"></div><aside><section class="panel">')
    parts.append(f'<h2>{t("artifacts_title")} <span class="count">{len(artifact_rows)}</span></h2>')
    parts.append(f'<div class="artifact-toolbar"><label for="artifact-sort" class="count">{t("sort_label")}</label><select id="artifact-sort"><option value="newest">{t("sort_newest")}</option><option value="oldest">{t("sort_oldest")}</option><option value="name">{t("sort_name")}</option><option value="version">{t("sort_version")}</option></select></div>')
    parts.append('<div class="artifact-list" id="artifact-list">')
    for item in artifact_rows:
        parts.append(artifact_card(item))
    parts.append(f'</div><div class="preview-box" id="preview-box"><div class="preview-head"><strong class="preview-title" id="preview-title">{t("preview_title")}</strong><button class="preview-close" id="preview-close" title="{t("preview_close")}">×</button></div><div class="preview-meta" id="preview-meta"></div><div class="preview-resizer" id="preview-resizer"><iframe id="artifact-preview" title="{t("preview_iframe_title")}" style="display:none"></iframe><div class="md-preview markdown" id="md-preview" style="display:none"></div><div class="resize-handle" id="resize-handle"></div></div></div>')
    parts.append(f'<details><summary>{t("recovered_files_title")} <span class="count">{len(recovered)}</span></summary><div class="artifact-toolbar"><label for="file-sort" class="count">{t("sort_label")}</label><select id="file-sort"><option value="newest">{t("sort_newest")}</option><option value="oldest">{t("sort_oldest")}</option><option value="name">{t("sort_name")}</option><option value="revisions">{t("sort_revisions")}</option></select></div><div class="file-list" id="file-list">')
    for item in recovered:
        updated = str(item.get("updated_at") or item.get("created_at") or "")
        created = str(item.get("created_at") or updated)
        revisions = int(item.get("revisions") or len(item.get("origins") or []) or 1)
        visible_updated = updated.replace("T", " ")[:16] if updated else t("date_unknown")
        created_note = f' · {t("file_created")} {created.replace("T", " ")[:16]}' if created and created != updated else ""
        filename = Path(str(item["local_path"])).name
        parts.append(f'<article class="file-card" data-date="{html.escape(updated, quote=True)}" data-name="{html.escape(filename.casefold(), quote=True)}" data-revisions="{revisions}"><div class="file-name">{html.escape(filename)}</div><div class="file-meta">{t("file_modified")} {html.escape(visible_updated)}{html.escape(created_note)} · {t("file_revisions")} {revisions}</div><a href="{href(str(item["local_path"]))}" target="_blank">{t("button_open")}</a></article>')
    parts.append("</div></details></section></aside></div>")
    sort_locale = L("sort_locale", lang)
    js_block = """<script>
function applyChatTheme(theme){
  if(theme!=='light'&&theme!=='dark')return;
  document.documentElement.dataset.theme=theme;
}
window.addEventListener('message',event=>{
  if(event.data&&event.data.type==='restore-theme')applyChatTheme(event.data.theme);
});
window.addEventListener('storage',event=>{if(event.key==='restore_theme')applyChatTheme(event.newValue);});
const artifactList=document.getElementById('artifact-list');
const artifactSort=document.getElementById('artifact-sort');
const copyChatButton=document.getElementById('copy-chat');
const copyChatLabels={
  ready:__COPY_READY__,
  copied:__COPY_COPIED__,
  failed:__COPY_FAILED__
};
function absoluteMarkdownLinks(markdown){
  return markdown.replace(/\]\(([^)]+)\)/g,(match,target)=>{
    if(/^[a-z][a-z0-9+.-]*:/i.test(target)||target.startsWith('#'))return match;
    try{return ']('+new URL(target,document.baseURI).href+')';}catch(error){return match;}
  });
}
async function copyText(text){
  if(navigator.clipboard&&window.isSecureContext){await navigator.clipboard.writeText(text);return;}
  const textarea=document.createElement('textarea');textarea.value=text;textarea.setAttribute('readonly','');textarea.style.cssText='position:fixed;opacity:0;pointer-events:none';document.body.appendChild(textarea);textarea.select();
  const copied=document.execCommand('copy');textarea.remove();if(!copied)throw new Error('copy command failed');
}
copyChatButton.addEventListener('click',async()=>{
  try{
    const markdown=JSON.parse(document.getElementById('copy-chat-data').textContent);
    await copyText(absoluteMarkdownLinks(markdown));copyChatButton.classList.add('copied');copyChatButton.title=copyChatLabels.copied;copyChatButton.setAttribute('aria-label',copyChatLabels.copied);
    setTimeout(()=>{copyChatButton.classList.remove('copied');copyChatButton.title=copyChatLabels.ready;copyChatButton.setAttribute('aria-label',copyChatLabels.ready);},1600);
  }catch(error){console.error('Copy chat:',error);copyChatButton.title=copyChatLabels.failed;copyChatButton.setAttribute('aria-label',copyChatLabels.failed);}
});
function sortArtifacts(){
  const mode=artifactSort.value;
  [...artifactList.children].sort((a,b)=>{
    if(mode==='name') return a.dataset.name.localeCompare(b.dataset.name,'__SORT_LOCALE__');
    if(mode==='version') return Number(b.dataset.version)-Number(a.dataset.version) || b.dataset.date.localeCompare(a.dataset.date);
    return mode==='oldest' ? a.dataset.date.localeCompare(b.dataset.date) : b.dataset.date.localeCompare(a.dataset.date);
  }).forEach(node=>artifactList.appendChild(node));
}
artifactSort.addEventListener('change',sortArtifacts); sortArtifacts();
const fileList=document.getElementById('file-list'), fileSort=document.getElementById('file-sort');
function sortFiles(){
  const mode=fileSort.value;
  [...fileList.children].sort((a,b)=>{
    if(mode==='name') return a.dataset.name.localeCompare(b.dataset.name,'__SORT_LOCALE__');
    if(mode==='revisions') return Number(b.dataset.revisions)-Number(a.dataset.revisions) || b.dataset.date.localeCompare(b.dataset.date);
    return mode==='oldest' ? a.dataset.date.localeCompare(b.dataset.date) : b.dataset.date.localeCompare(a.dataset.date);
  }).forEach(node=>fileList.appendChild(node));
}
fileSort.addEventListener('change',sortFiles); sortFiles();
// Mouse-resizable artifacts pane. Width is remembered for the next chat.
const chatLayout=document.querySelector('.layout'), rightResizer=document.getElementById('right-resizer');
function setAsideWidth(width){
  const maximum=Math.min(700,chatLayout.getBoundingClientRect().width*.65);
  const value=Math.max(240,Math.min(maximum,width));
  chatLayout.style.setProperty('--aside-w',value+'px');
  return value;
}
try{const saved=Number(localStorage.getItem('restore_aside_width'));if(saved)setAsideWidth(saved);}catch(error){}
rightResizer.addEventListener('pointerdown',event=>{
  if(matchMedia('(max-width:900px)').matches)return;
  event.preventDefault();rightResizer.setPointerCapture(event.pointerId);rightResizer.classList.add('dragging');document.body.classList.add('resizing');
});
rightResizer.addEventListener('pointermove',event=>{
  if(!rightResizer.hasPointerCapture(event.pointerId))return;
  const bounds=chatLayout.getBoundingClientRect();setAsideWidth(bounds.right-event.clientX-8);
});
function finishAsideResize(event){
  if(!rightResizer.hasPointerCapture(event.pointerId))return;
  rightResizer.releasePointerCapture(event.pointerId);rightResizer.classList.remove('dragging');document.body.classList.remove('resizing');
  try{localStorage.setItem('restore_aside_width',parseFloat(getComputedStyle(chatLayout).getPropertyValue('--aside-w'))||340);}catch(error){}
}
rightResizer.addEventListener('pointerup',finishAsideResize);rightResizer.addEventListener('pointercancel',finishAsideResize);
// Inline artifacts and the right-hand list are two views of the same item.
function onButton(selector, action){
  document.addEventListener('click',event=>{
    const button=event.target instanceof Element && event.target.closest(selector);
    if(button)action(button);
  });
}
onButton('.locate-artifact',button=>{
  const target=artifactList.querySelector('.artifact-card[data-artifact-key="'+button.dataset.artifactKey+'"]');
  if(!target)return;
  target.scrollIntoView({behavior:'smooth',block:'center'});
  target.classList.remove('located');void target.offsetWidth;target.classList.add('located');
});
// Preview is moved into the selected artifact card and grows only downward.
const previewBox=document.getElementById('preview-box'), previewIframe=document.getElementById('artifact-preview'), previewTitle=document.getElementById('preview-title'), previewMeta=document.getElementById('preview-meta'), mdPreview=document.getElementById('md-preview');
const previewHome=previewBox.parentElement;
function closePreview(){
  previewIframe.removeAttribute('src');previewIframe.style.display='none';
  mdPreview.innerHTML='';mdPreview.style.display='none';mdPreview.classList.remove('plain');
  previewMeta.textContent='';previewBox.classList.remove('open','standalone');
  previewHome.classList.remove('attachment-preview-open');
  previewHome.appendChild(previewBox);
}
onButton('.preview-button',button=>{
  var title=button.dataset.title, isMd=button.dataset.md==='1', isText=button.dataset.text==='1';
  var card=button.closest('.artifact-card');
  var attachment=button.closest('.input-attachment-wrap');
  previewHome.classList.toggle('attachment-preview-open',Boolean(attachment));
  previewBox.classList.toggle('standalone',Boolean(attachment));
  if(attachment)previewHome.prepend(previewBox);else if(card)card.appendChild(previewBox);
  previewTitle.textContent=title;
  previewMeta.textContent=button.dataset.meta||'';
  if(isText){
    previewIframe.style.display='none';previewIframe.removeAttribute('src');
    mdPreview.style.display='block';mdPreview.classList.add('plain');mdPreview.innerHTML='';
    var textEl=attachment?attachment.querySelector('.attachment-data'):null;
    var plainText='';
    if(textEl){try{plainText=JSON.parse(textEl.textContent);}catch(error){console.error('Attachment preview JSON:',error);}}
    var pre=document.createElement('pre');pre.className='plain-text-preview';pre.textContent=plainText;mdPreview.appendChild(pre);
  }else if(isMd){
    previewIframe.style.display='none';previewIframe.removeAttribute('src');
    mdPreview.style.display='block';mdPreview.classList.remove('plain');
    var mdEl=card?card.querySelector('.md-data'):null;
    var mdText='';
    if(mdEl){try{mdText=JSON.parse(mdEl.textContent);}catch(error){console.error('Markdown preview JSON:',error);}}
    mdPreview.innerHTML=mdText?renderMarkdownSimple(mdText):'<p style="color:#999">Empty.</p>';
  }else{
    var src=button.dataset.src;
    mdPreview.style.display='none';mdPreview.innerHTML='';mdPreview.classList.remove('plain');
    previewIframe.style.display='block';
    previewIframe.src=src;
  }
  previewBox.classList.add('open');
});
onButton('.attachment-copy',async button=>{
  var data=button.closest('.input-attachment-wrap').querySelector('.attachment-data'), value='';
  if(data){try{value=JSON.parse(data.textContent);}catch(error){console.error('Attachment copy JSON:',error);}}
  try{await copyText(value);button.classList.add('copied');setTimeout(()=>button.classList.remove('copied'),1200);}catch(error){console.error('Attachment copy:',error);}
});
const summaryDetails=document.getElementById('conv-summary');
if(summaryDetails){
  const previewButton=document.getElementById('summary-preview');
  const copyButton=document.getElementById('summary-copy');
  previewButton.addEventListener('click',()=>{summaryDetails.open=!summaryDetails.open;});
  summaryDetails.addEventListener('toggle',()=>previewButton.setAttribute('aria-expanded',String(summaryDetails.open)));
  copyButton.addEventListener('click',async()=>{
    const url=document.getElementById('summary-download').getAttribute('href');
    const label=copyButton.getAttribute('aria-label');
    try{
      await copyText(decodeURIComponent(url.slice(url.indexOf(',')+1)).trimEnd());
      copyButton.classList.add('copied');
      copyButton.title=copyChatLabels.copied;
    }catch(error){
      copyButton.title=copyChatLabels.failed;
    }
    setTimeout(()=>{copyButton.title=label;copyButton.classList.remove('copied');},1600);
  });
}
document.getElementById('preview-close').addEventListener('click',closePreview);
document.addEventListener('keydown',event=>{{if(event.key==='Escape'&&previewBox.classList.contains('open'))closePreview();}});
// Inline previews grow to the right; the sidebar is anchored on the right
// and grows to the left.
const resizeHandle=document.getElementById('resize-handle');
const previewResizer=document.getElementById('preview-resizer');
const asideEl=document.querySelector('aside');
resizeHandle.style.touchAction='none';
let previewDrag=null;
resizeHandle.addEventListener('pointerdown',event=>{{
  if(event.button!==0||previewDrag)return;
  event.preventDefault();event.stopPropagation();
  const inAside=asideEl.contains(previewBox);
  const card=previewBox.closest('.inline-artifact');
  previewDrag={
    x:event.clientX,
    y:event.clientY,
    height:previewResizer.getBoundingClientRect().height,
    width:(inAside?asideEl:card)?.getBoundingClientRect().width,
    card,
    inAside
  };
  resizeHandle.setPointerCapture(event.pointerId);
  document.body.classList.add('resizing-preview');
}});
resizeHandle.addEventListener('pointermove',event=>{{
  if(!previewDrag||!resizeHandle.hasPointerCapture(event.pointerId))return;
  event.preventDefault();
  const rect=previewResizer.getBoundingClientRect();
  // Height: distance from top of resizer to pointer Y
  let newH=previewDrag.height+event.clientY-previewDrag.y;
  newH=Math.max(120,Math.min(newH,window.innerHeight-rect.top-10));
  previewResizer.style.height=newH+'px';
  const dx=event.clientX-previewDrag.x;
  if(dx&&previewDrag.card){
    const message=previewDrag.card.closest('.message');
    const main=message.closest('main');
    const style=getComputedStyle(main);
    const limit=main.clientWidth-parseFloat(style.paddingLeft)-parseFloat(style.paddingRight);
    const width=Math.min(limit,Math.max(240,previewDrag.width+dx));
    message.style.maxWidth=Math.max(860,width)+'px';
    previewDrag.card.style.maxWidth=width+'px';
  }else if(dx&&previewDrag.inAside&&rightResizer.offsetWidth){{
    const maxW=Math.min(900,window.innerWidth*0.85);
    const width=Math.max(280,Math.min(previewDrag.width-dx,maxW));
    chatLayout.style.setProperty('--aside-w',width+'px');
  }}
}});
function finishResize(event){{
  if(!previewDrag)return;
  previewDrag=null;
  if(resizeHandle.hasPointerCapture(event.pointerId))resizeHandle.releasePointerCapture(event.pointerId);
  document.body.classList.remove('resizing-preview');
}}
resizeHandle.addEventListener('pointerup',finishResize);
resizeHandle.addEventListener('pointercancel',finishResize);
resizeHandle.addEventListener('lostpointercapture',finishResize);
// Small dependency-free renderer for local previews. The Python generator
// already guarantees that mdText is a string; all source HTML is escaped here.
function renderMarkdownSimple(md){
  const escapeHtml=value=>value.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const codeBlocks=[];
  md=md.replace(/```([^\\n]*)\\n([\\s\\S]*?)```/g,(match,lang,code)=>{
    const index=codeBlocks.length;
    codeBlocks.push('<pre><code>'+escapeHtml(code)+'</code></pre>');
    return '@@CODEBLOCK'+index+'@@';
  });
  md=escapeHtml(md);
  md=md.replace(/^###\\s+(.+)$/gm,'<h3>$1</h3>').replace(/^##\\s+(.+)$/gm,'<h2>$1</h2>').replace(/^#\\s+(.+)$/gm,'<h1>$1</h1>');
  md=md.replace(/^&gt;\\s?(.*)$/gm,'<blockquote>$1</blockquote>');
  md=md.replace(/(?:^(?:\\s*[-*+]\\s+.+)\\n?)+/gm,block=>'<ul>'+block.trim().split('\\n').map(line=>'<li>'+line.replace(/^\\s*[-*+]\\s+/,'')+'</li>').join('')+'</ul>');
  md=md.replace(/(?:^(?:\\s*\\d+[.]\\s+.+)\\n?)+/gm,block=>'<ol>'+block.trim().split('\\n').map(line=>'<li>'+line.replace(/^\\s*\\d+[.]\\s+/, '')+'</li>').join('')+'</ol>');
  md=md.replace(/\\*\\*(.+?)\\*\\*/g,'<strong>$1</strong>').replace(/~~(.+?)~~/g,'<del>$1</del>').replace(/\\*([^*\\n]+)\\*/g,'<em>$1</em>');
  md=md.replace(/`([^`\\n]+)`/g,'<code>$1</code>');
  md=md.replace(/\\[([^\\]]+)\\]\\(([^)\\s]+)\\)/g,'<a href="$2" target="_blank" rel="noopener">$1</a>');
  // Tables: header row | separator | data rows
  md=md.replace(/(?:^\\|.+)\\n(?:\\|[\\s:-]+\\|[^\\n]*)\\n((?:\\|.+)\\n?)+/gm,table=>{
    const rows=table.trim().split('\\n').filter(r=>r.trim());
    if(rows.length<2)return table;
    const parseRow=row=>row.replace(/^\\|/,'').replace(/\\|$/,'').split('|').map(c=>c.trim());
    const header=parseRow(rows[0]);
    const body=rows.slice(2).map(parseRow);
    let out='<table><thead><tr>'+header.map(c=>'<th>'+c+'</th>').join('')+'</tr></thead><tbody>';
    for(const row of body)out+='<tr>'+row.map(c=>'<td>'+c+'</td>').join('')+'</tr>';
    return out+'</tbody></table>';
  });
  md=md.split(/\\n\\s*\\n/).map(block=>{
    const value=block.trim();
    if(!value)return '';
    if(/^<(?:h[1-6]|ul|ol|pre|blockquote)/.test(value)||/^@@CODEBLOCK\\d+@@$/.test(value))return value;
    return '<p>'+value.replace(/\\n/g,'<br>')+'</p>';
  }).join('\\n');
  return md.replace(/@@CODEBLOCK(\\d+)@@/g,(match,index)=>codeBlocks[Number(index)]||'');
}
// Parse message markup only as it approaches the chat viewport.
function mountMessageWindow(){
  const container=document.getElementById('message-window');
  if(!container)return;
  const messages=JSON.parse(document.getElementById('message-data').textContent);
  const mounted=new Set();
  const openDetails=new Map();
  const slots=[];
  function mount(index){
    if(mounted.has(index))return;
    const slot=slots[index];
    slot.innerHTML=messages[index][0];
    slot.style.height='';
    const open=openDetails.get(index);
    if(open)slot.querySelectorAll('details').forEach((detail,i)=>{detail.open=open.has(i);});
    mounted.add(index);
  }
  function unmount(index){
    if(!mounted.has(index))return;
    const slot=slots[index];
    if(slot.contains(document.activeElement)||slot.contains(previewBox))return;
    const open=new Set();
    slot.querySelectorAll('details').forEach((detail,i)=>{if(detail.open)open.add(i);});
    if(open.size)openDetails.set(index,open);else openDetails.delete(index);
    slot.style.height=slot.getBoundingClientRect().height+'px';
    slot.replaceChildren();
    mounted.delete(index);
  }
  for(const [index,message] of messages.entries()){
    const slot=document.createElement('div');
    slot.className='message-slot';
    slot.dataset.index=index;
    slot.style.height=message[1]+'px';
    slots.push(slot);
  }
  container.replaceChildren(...slots);
  if('IntersectionObserver' in window){
    const main=container.closest('main');
    let root;
    let observer;
    function observeSlots(){
      const nextRoot=getComputedStyle(main).overflowY==='visible'?null:main;
      if(observer&&nextRoot===root)return;
      observer?.disconnect();
      root=nextRoot;
      observer=new IntersectionObserver(entries=>{
        for(const entry of entries){
          const index=Number(entry.target.dataset.index);
          if(entry.isIntersecting)mount(index);else unmount(index);
        }
      },{root,rootMargin:'800px 0px'});
      slots.forEach(slot=>observer.observe(slot));
    }
    observeSlots();
    // Revisit a pinned message after it loses focus or its preview closes.
    function releaseOffscreen(){
      const viewport=root?root.getBoundingClientRect():{top:0,bottom:window.innerHeight};
      for(const index of mounted){
        const rect=slots[index].getBoundingClientRect();
        if(rect.bottom<viewport.top-800||rect.top>viewport.bottom+800)unmount(index);
      }
    }
    document.addEventListener('focusout',()=>requestAnimationFrame(releaseOffscreen));
    document.addEventListener('click',()=>requestAnimationFrame(releaseOffscreen));
    window.addEventListener('resize',observeSlots);
  }else{
    slots.forEach((slot,index)=>mount(index));
  }
  function jumpToMessage(){
    const match=/^#message-(\\d+)$/.exec(location.hash);
    if(!match)return;
    const index=Number(match[1])-1;
    if(!slots[index])return;
    mount(index);
    slots[index].scrollIntoView();
  }
  window.addEventListener('hashchange',jumpToMessage);
  jumpToMessage();
  return mount;
}
const mountMessage=mountMessageWindow();
const chatMain=document.querySelector('main');
const scrollRoot=()=>getComputedStyle(chatMain).overflowY==='visible'?document.scrollingElement:chatMain;
const messageBlocks=()=>chatMain.querySelectorAll(mountMessage?'.message-slot':'.message');
const readingLine=()=>chatMain.querySelector('.hero').getBoundingClientRect().bottom;
window.addEventListener('message',event=>{
  if(event.source!==window.parent||event.data?.type!=='restore-get-position')return;
  const blocks=[...messageBlocks()];
  const top=readingLine();
  const index=blocks.findIndex(block=>block.getBoundingClientRect().bottom>top);
  const rect=blocks[index]?.getBoundingClientRect();
  const position=rect?{index,fraction:Math.max(0,(top-rect.top)/Math.max(1,rect.height)),
    offset:Math.max(0,rect.top-top)}:null;
  event.source.postMessage({type:'restore-chat-position',position},'*');
});
const savedPosition=new URLSearchParams(location.search).get('position');
if(savedPosition){
  try{
    const {index,fraction,offset}=JSON.parse(savedPosition);
    const block=messageBlocks()[index];
    if(block&&Number.isFinite(fraction)&&Number.isFinite(offset)){
      mountMessage?.(index);
      requestAnimationFrame(()=>{
        const rect=block.getBoundingClientRect();
        scrollRoot().scrollTop+=rect.top-readingLine()+rect.height*fraction-offset;
      });
    }
  }catch(error){console.warn('Could not restore chat position',error);}
}
</script></body></html>""".replace("__SORT_LOCALE__", sort_locale).replace("__COPY_READY__", json.dumps(t("button_copy_chat"), ensure_ascii=False)).replace("__COPY_COPIED__", json.dumps(t("button_copy_chat_copied"), ensure_ascii=False)).replace("__COPY_FAILED__", json.dumps(t("button_copy_chat_failed"), ensure_ascii=False))
    parts.append(js_block)
    return "".join(parts)


def render_index(rows: list[dict[str, Any]]) -> str:
    lines = ["# Restored Claude conversations", "", "The export contains metadata, chats, extracted attachment text, and generator sources. Presented binary files are usually metadata-only in the export; see each conversation's `presented_files.json`.", "", "| Created | Conversation | Recovered files | Presented files |", "|---|---|---:|---:|"]
    for row in rows:
        title = str(row["title"]).replace("|", "\\|")
        folder = str(row["folder"]).replace(" ", "%20")
        lines.append(f"| {row['created'][:10]} | [{title}]({folder}/chat.md) | {row['recovered_files']} | {row['presented_files']} |")
    return "\n".join(lines) + "\n"


def chat_view_url(folder: str, lang: str) -> str:
    return f"viewer{lang_suffix(lang)}.html?chat={quote(folder, safe='')}"


def split_chat_page(page: str) -> tuple[str, str, str]:
    """Separate the shared document shell and runtime from the conversation markup."""
    head, body = page.split("</head><body>", 1)
    markup, runtime = body.rsplit("<script>", 1)
    return head, markup, runtime.removesuffix("</script></body></html>")


def render_chat_viewer(lang: str) -> str:
    head, _, runtime = split_chat_page(render_chat_html({}, [], [], [], Path("."), lang))
    return head + f"""</head><body><script>
const folder=new URLSearchParams(location.search).get('chat');
if(folder && folder!=='.' && folder!=='..' && !/[\\\\/]/.test(folder)){{
  const url=new URL('conversations/'+encodeURIComponent(folder)+'/',location.href);
  const data=document.createElement('script');
  data.src=new URL('chat-data.js',url).href;
  data.onerror=()=>{{document.body.textContent='Unable to load chat-data.js';}};
  data.onload=()=>{{
    const chat=window.restoredChatData;
    delete window.restoredChatData;
    const base=document.createElement('base');
    base.href=url.href;
    document.head.appendChild(base);
    document.title=chat.title;
    document.body.innerHTML=chat.pages[{json.dumps(lang)}];
    document.addEventListener('click',event=>{{
      const link=event.target instanceof Element&&event.target.closest('a[href^="#"]');
      if(link){{event.preventDefault();location.hash=link.getAttribute('href');}}
    }});
    startChat();
  }};
  document.head.appendChild(data);
}}
function startChat(){{
{runtime}
}}
</script></body></html>"""


def render_index_html(rows: list[dict[str, Any]], lang: str = DEFAULT_LANG) -> str:
    ti = lambda key: L(f"index.{key}", lang)
    sort_locale = L("sort_locale", lang)
    cards = []
    newest_rows = sorted(rows, key=lambda item: (str(item.get("updated") or item["created"]), str(item["title"])), reverse=True)
    for row in newest_rows:
        activity_date = str(row.get("updated") or row["created"])
        folder = str(row["folder"])
        badge = ""
        if row.get("has_summary"):
            label = html.escape(L("chat.summary_label", lang))
            badge = (
                f'<svg class="summary-marker" viewBox="0 0 16 16" role="img" aria-label="{label}">'
                f'<title>{label}</title><rect x="3" y="2" width="10" height="12" rx="2"/>'
                '<path d="M5.5 6h5m-5 3h5m-5 3H8"/></svg>'
            )
        cards.append(
            f'<a class="conversation" href="{chat_view_url(folder, lang)}" target="chat-frame" data-folder="{href(folder)}" data-date="{html.escape(activity_date, quote=True)}" data-title="{html.escape(str(row["title"]).casefold(), quote=True)}"><strong>{html.escape(str(row["title"]))}</strong>'
            f'{badge}<span>{html.escape(activity_date.replace("T", " ")[:16])} · {row["presented_files"]} {ti("card_artifacts")} · {row["recovered_files"]} {ti("card_files")}</span></a>'
        )
    first_url = chat_view_url(str(newest_rows[0]["folder"]), lang) if newest_rows else ""

    # Language switcher options
    lang_options = []
    for code in AVAILABLE_LANGS:
        name = LOCALES[code].get("lang_name", code)
        sel = " selected" if code == lang else ""
        lang_options.append(f'<option value="{code}"{sel}>{html.escape(name)}</option>')

    return f"""<!doctype html><html lang="{L("html_lang", lang)}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(ti("page_title"))}</title><script>try{{document.documentElement.dataset.theme=localStorage.getItem('restore_theme')||'dark'}}catch(error){{document.documentElement.dataset.theme='dark'}}</script><style>
:root{{color-scheme:light;--sidebar-w:306px;--bg:#f7f6f2;--sidebar:#f0eee8;--surface:#fff;--surface-2:#e8e5de;--ink:#292724;--muted:#77736c;--line:#dedbd2;--line-strong:#c8c3b8;--accent:#d97757;--icon:#625f59}}
:root[data-theme="dark"]{{color-scheme:dark;--bg:#181817;--sidebar:#1f1f1e;--surface:#292927;--surface-2:#343330;--ink:#eeeDE8;--muted:#a6a39c;--line:#343330;--line-strong:#4a4843;--accent:#d97757;--icon:#c4c1ba}}
*{{box-sizing:border-box}}html,body{{height:100%;margin:0;font:13px/1.42 ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}}body{{display:grid;grid-template-columns:var(--sidebar-w) 8px minmax(0,1fr);overflow:hidden;background:var(--bg);color:var(--ink);transition:background .18s ease,color .18s ease}}body.resizing{{cursor:col-resize;user-select:none}}button,input,select{{font:inherit}}
.sidebar{{height:100vh;background:var(--sidebar);color:var(--ink);display:flex;flex-direction:column;border-right:1px solid var(--line);overflow-x:hidden;min-width:0;transition:background .18s ease}}.sidebar-head{{height:48px;padding:0 10px 0 14px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--line)}}.brand{{font-family:Georgia,"Times New Roman",serif;font-size:23px;font-weight:700;letter-spacing:-.4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}.brand small{{font:500 10px/1 ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:var(--muted);letter-spacing:.02em}}
.theme-toggle{{width:32px;height:32px;flex:0 0 32px;display:grid;place-items:center;border:1px solid transparent;border-radius:8px;background:transparent;color:var(--icon);cursor:pointer}}.theme-toggle:hover{{border-color:var(--line);background:var(--surface)}}.theme-toggle svg{{width:17px;height:17px;fill:none;stroke:currentColor;stroke-width:1.7;stroke-linecap:round;stroke-linejoin:round}}.theme-toggle .sun{{display:none}}:root[data-theme="dark"] .theme-toggle .moon{{display:none}}:root[data-theme="dark"] .theme-toggle .sun{{display:block}}
.tools{{padding:11px 10px 8px;display:grid;gap:7px}}.tools input,.tools select{{width:100%;border:1px solid var(--line);border-radius:8px;background:var(--surface);color:var(--ink);padding:8px 9px;outline:0}}.tools input:focus,.tools select:focus{{border-color:var(--line-strong);box-shadow:0 0 0 2px color-mix(in srgb,var(--accent) 18%,transparent)}}.tools input::placeholder{{color:var(--muted)}}
.links{{display:flex;gap:5px 11px;padding:2px 14px 10px;flex-wrap:wrap;align-items:center;border-bottom:1px solid var(--line)}}.links a{{color:var(--muted);font-size:11px;text-decoration:none}}.links a:hover{{color:var(--ink)}}.lang-switch{{margin-left:auto;border:1px solid var(--line);border-radius:7px;background:var(--surface);color:var(--ink);padding:3px 5px;font-size:11px}}
.conversation-list{{overflow:auto;padding:7px 7px 18px;scrollbar-width:thin;scrollbar-color:var(--line-strong) transparent}}.conversation{{position:relative;display:block;color:var(--ink);text-decoration:none;padding:8px 10px 8px 25px;border-radius:8px;margin:1px 0}}.conversation[hidden]{{display:none!important}}.conversation::before{{content:"";position:absolute;left:10px;top:14px;width:5px;height:5px;border:1px solid var(--muted);border-radius:50%;opacity:.55}}.conversation:hover{{background:color-mix(in srgb,var(--surface) 65%,transparent)}}.conversation.active{{background:var(--surface-2)}}.conversation.active::before{{background:var(--accent);border-color:var(--accent);opacity:1}}.conversation strong{{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-weight:520}}.conversation span{{display:block;color:var(--muted);font-size:10px;margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}.empty{{color:var(--muted);padding:15px;display:none}}
.left-resizer{{position:relative;height:100vh;cursor:col-resize;touch-action:none;background:var(--bg)}}.left-resizer::after{{content:"";position:absolute;left:3px;top:10px;bottom:10px;width:1px;background:var(--line)}}.left-resizer:hover::after,.left-resizer.dragging::after{{width:2px;background:var(--accent)}}.viewer{{min-width:0;background:var(--bg)}}iframe{{display:block;width:100%;height:100vh;border:0;background:var(--bg)}}
.conversation:has(.summary-marker) strong{{padding-right:19px}}
.summary-marker{{position:absolute;right:9px;top:9px;width:15px;height:15px;fill:none;stroke:var(--accent);stroke-width:1.35;stroke-linecap:round}}
@media(max-width:720px){{body{{grid-template-columns:1fr;overflow:auto}}.sidebar{{height:46vh}}.left-resizer{{display:none}}.viewer{{height:54vh}}.viewer iframe{{height:54vh}}}}
</style></head><body><nav class="sidebar"><div class="sidebar-head"><div class="brand">Claude <small>RESTORED</small></div><button type="button" class="theme-toggle" id="theme-toggle" aria-label="{html.escape(ti("theme_light"), quote=True)}" title="{html.escape(ti("theme_light"), quote=True)}"><svg class="moon" viewBox="0 0 24 24" aria-hidden="true"><path d="M20.5 14.4A8.5 8.5 0 0 1 9.6 3.5 8.5 8.5 0 1 0 20.5 14.4Z"/></svg><svg class="sun" viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="3.5"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg></button></div><div class="tools"><input id="search" type="search" placeholder="{html.escape(ti("search_placeholder"))}"><select id="sort"><option value="newest">{html.escape(ti("sort_newest"))}</option><option value="oldest">{html.escape(ti("sort_oldest"))}</option><option value="title">{html.escape(ti("sort_title"))}</option></select></div><div class="links"><a href="RESTORATION_NOTES.md" target="_blank">{html.escape(ti("link_description"))}</a><a href="raw/artifacts/INDEX.html" target="_blank">{html.escape(ti("link_artifacts"))}</a><a href="project_documents/" target="_blank">{html.escape(ti("link_projects"))}</a><a href="SUMMARY.json" target="_blank">{html.escape(ti("link_summary"))}</a><select class="lang-switch" id="lang-switch">{"".join(lang_options)}</select></div><div class="conversation-list" id="conversation-list">""" + "".join(cards) + f"""<div class="empty" id="empty">{html.escape(ti("empty"))}</div></div></nav><div class="left-resizer" id="left-resizer" role="separator" aria-orientation="vertical"></div><main class="viewer"><iframe name="chat-frame" id="chat-frame" src="__FIRST_CHAT__" title="{html.escape(ti("iframe_title"))}"></iframe></main><script>
const list=document.getElementById('conversation-list'), search=document.getElementById('search'), sort=document.getElementById('sort');
const conversations=[...list.querySelectorAll('.conversation')];
const leftResizer=document.getElementById('left-resizer');
const chatFrame=document.getElementById('chat-frame'), themeToggle=document.getElementById('theme-toggle');
const themeLabels={{light:{json.dumps(ti("theme_light"), ensure_ascii=False)},dark:{json.dumps(ti("theme_dark"), ensure_ascii=False)}}};
function applyTheme(theme,persist=true){{
  if(theme!=='light'&&theme!=='dark')theme='dark';
  document.documentElement.dataset.theme=theme;
  const target=theme==='dark'?'light':'dark';
  themeToggle.title=themeLabels[target];themeToggle.setAttribute('aria-label',themeLabels[target]);
  if(persist){{try{{localStorage.setItem('restore_theme',theme);}}catch(error){{}}}}
  try{{chatFrame.contentWindow.postMessage({{type:'restore-theme',theme}},'*');}}catch(error){{}}
}}
applyTheme(document.documentElement.dataset.theme,false);
themeToggle.addEventListener('click',()=>applyTheme(document.documentElement.dataset.theme==='dark'?'light':'dark'));
chatFrame.addEventListener('load',()=>applyTheme(document.documentElement.dataset.theme,false));
function setSidebarWidth(width){{
  const maximum=Math.min(650,window.innerWidth*.55);
  const value=Math.max(200,Math.min(maximum,width));
  document.body.style.setProperty('--sidebar-w',value+'px');
  return value;
}}
try{{const saved=Number(localStorage.getItem('restore_sidebar_width'));if(saved)setSidebarWidth(saved);}}catch(error){{}}
leftResizer.addEventListener('pointerdown',event=>{{
  if(matchMedia('(max-width:720px)').matches)return;
  event.preventDefault();leftResizer.setPointerCapture(event.pointerId);leftResizer.classList.add('dragging');document.body.classList.add('resizing');
}});
leftResizer.addEventListener('pointermove',event=>{{if(leftResizer.hasPointerCapture(event.pointerId))setSidebarWidth(event.clientX-4);}});
function finishSidebarResize(event){{
  if(!leftResizer.hasPointerCapture(event.pointerId))return;
  leftResizer.releasePointerCapture(event.pointerId);leftResizer.classList.remove('dragging');document.body.classList.remove('resizing');
  try{{localStorage.setItem('restore_sidebar_width',parseFloat(getComputedStyle(document.body).getPropertyValue('--sidebar-w'))||290);}}catch(error){{}}
}}
leftResizer.addEventListener('pointerup',finishSidebarResize);leftResizer.addEventListener('pointercancel',finishSidebarResize);
function update(){{const q=search.value.trim().toLocaleLowerCase('{sort_locale}'); conversations.forEach(x=>x.hidden=q&&!x.dataset.title.includes(q)); const visible=conversations.filter(x=>!x.hidden); visible.sort((a,b)=>sort.value==='title'?a.dataset.title.localeCompare(b.dataset.title,'{sort_locale}'):(sort.value==='oldest'?a.dataset.date.localeCompare(b.dataset.date):b.dataset.date.localeCompare(a.dataset.date))).forEach(x=>list.appendChild(x)); document.getElementById('empty').style.display=visible.length?'none':'block';}}
search.addEventListener('input',update);sort.addEventListener('change',update);update();
conversations.forEach(link=>link.addEventListener('click',()=>{{conversations.forEach(x=>x.classList.remove('active'));link.classList.add('active');
  sessionStorage.setItem('restore_active_chat',link.dataset.folder);
}}));
if(conversations.length) conversations[0].classList.add('active');
// Language switcher: save state, set cookie, redirect to localized index
document.getElementById('lang-switch').addEventListener('change',function(){{
  var lang=this.value;
  document.cookie='restore_lang='+lang+';path=/;max-age=31536000';
  // Save selected conversation folder and scroll position in URL params
  var active=document.querySelector('.conversation.active');
  var chatParam='';
  if(active){{
    chatParam='?chat='+encodeURIComponent(active.dataset.folder);
  }}
  sessionStorage.setItem('restore_sidebar_scroll',document.querySelector('.conversation-list').scrollTop);
  var suffix=lang==='{DEFAULT_LANG}'?'':'_'+lang;
  var base=document.location.pathname.replace(/(_(en|zh|ru))?\.html$/,'');
  const next=new URL(base+suffix+'.html'+chatParam,location.href);
  let timer;
  function navigate(position){{
    clearTimeout(timer);
    window.removeEventListener('message',receivePosition);
    if(position&&chatParam)next.searchParams.set('position',JSON.stringify(position));
    location.href=next.href;
  }}
  function receivePosition(event){{
    if(event.source===chatFrame.contentWindow&&event.data?.type==='restore-chat-position'){{
      navigate(event.data.position);
    }}
  }}
  window.addEventListener('message',receivePosition);
  timer=setTimeout(()=>navigate(null),300);
  chatFrame.contentWindow.postMessage({{type:'restore-get-position'}},'*');
}});
// Restore selected conversation from ?chat= param or sessionStorage
(function(){{
  var params=new URLSearchParams(document.location.search);
  var chatFolder=params.get('chat');
  if(!chatFolder){{
    // Also check hash as fallback
    chatFolder=decodeURIComponent(document.location.hash.slice(1))||null;
  }}
  // Try sessionStorage as fallback
  if(!chatFolder) chatFolder=sessionStorage.getItem('restore_active_chat');
  var savedScroll=sessionStorage.getItem('restore_sidebar_scroll');
  if(chatFolder){{
    // Find the conversation link by folder name
    var link=null;
    conversations.forEach(function(c){{
      if(c.dataset.folder===chatFolder || decodeURIComponent(c.dataset.folder)===chatFolder) link=c;
    }});
    if(link){{
      conversations.forEach(x=>x.classList.remove('active'));
      link.classList.add('active');
      const viewerUrl=new URL(link.href);
      if(params.has('position'))viewerUrl.searchParams.set('position',params.get('position'));
      chatFrame.src=viewerUrl.href;
      // Scroll the link into view
      link.scrollIntoView({{block:'nearest'}});
      sessionStorage.setItem('restore_active_chat',chatFolder);
    }}
  }}
  if(savedScroll){{document.querySelector('.conversation-list').scrollTop=Number(savedScroll);sessionStorage.removeItem('restore_sidebar_scroll');}}
}})();
</script></body></html>""".replace("__FIRST_CHAT__", html.escape(first_url, quote=True))


def restore_conversation(conversation: dict[str, Any], root: Path) -> dict[str, Any]:
    uuid = str(conversation.get("uuid") or "unknown")
    created = str(conversation.get("created_at") or "unknown-date")
    dirname = f"{created[:10]}_{safe_component(conversation.get('name'), 72)}_{uuid[:8]}"
    conv_dir = root / dirname
    files_dir = conv_dir / "files"
    commands_dir = conv_dir / "commands"
    inline_dir = conv_dir / "inline_scripts"
    attachments_dir = conv_dir / "attachments"
    conv_dir.mkdir(parents=True, exist_ok=True)
    json_dump(conv_dir / "conversation.json", conversation)
    (conv_dir / "chat.md").write_text(render_chat_markdown(conversation), encoding="utf-8")

    virtual_files: dict[str, str] = {}
    path_origins: dict[str, list[str]] = {}
    path_dates: dict[str, list[str]] = {}
    edit_failures: list[dict[str, Any]] = []
    presented: list[dict[str, Any]] = []
    presented_by_event: dict[tuple[str, str], dict[str, Any]] = {}
    command_count = 0
    inline_count = 0
    attachment_count = 0
    attachment_records: list[dict[str, Any]] = []

    def record_presented(item: dict[str, Any], tool_use_id: str = "") -> None:
        """Record one presentation event, merging its tool call/result pair only."""
        raw_path = str(item.get("file_path") or "")
        if not raw_path:
            return
        message_uuid = str(item.get("message_uuid") or "")
        event_id = tool_use_id or f"message:{message_uuid}:{item.get('tool') or ''}"
        event_key = (event_id, raw_path)
        existing = presented_by_event.get(event_key)
        if existing is None:
            item["tool_use_id"] = tool_use_id or None
            presented.append(item)
            presented_by_event[event_key] = item
            return
        # A tool_result carries richer resource metadata than its tool_use.
        # Merge it into the same event without collapsing later presentations
        # of the same path, which have a different tool-use ID/message UUID.
        for key, value in item.items():
            if value is not None and value != "":
                existing[key] = value

    for msg_index, message in enumerate(conversation.get("chat_messages") or [], 1):
        message_date = str(message.get("created_at") or "")
        for file_index, attachment in enumerate(message.get("attachments") or [], 1):
            if not isinstance(attachment, dict) or attachment.get("extracted_content") is None:
                continue
            attachment_count += 1
            original_name = str(attachment.get("file_name") or "")
            name = original_name or f"message_{msg_index:04d}_attachment_{file_index}"
            # `extracted_content` is always text, even when the original upload
            # was DOCX/PDF.  Keep that distinction explicit so the recovered
            # text is not mistaken for the unavailable source binary.
            target = attachments_dir / f"message_{msg_index:04d}_{file_index:02d}_{safe_component(Path(str(name)).stem)}.extracted.txt"
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(str(attachment["extracted_content"]), encoding="utf-8")
            attachment_records.append({
                "message_uuid": message.get("uuid"),
                "name": original_name,
                "file_type": attachment.get("file_type"),
                "local_path": target.relative_to(conv_dir).as_posix(),
                "bytes": int(attachment.get("file_size") or len(str(attachment["extracted_content"]).encode("utf-8"))),
                "lines": len(str(attachment["extracted_content"]).splitlines()),
                "content": str(attachment["extracted_content"]),
            })

        for block in message.get("content") or []:
            if not isinstance(block, dict):
                continue
            if block.get("type") == "tool_result":
                for item in block.get("content") or []:
                    if isinstance(item, dict) and item.get("type") == "local_resource":
                        record_presented({
                            "message_uuid": message.get("uuid"),
                            "tool": block.get("name"),
                            "uuid": item.get("uuid"),
                            "file_path": item.get("file_path"),
                            "name": item.get("name"),
                            "mime_type": item.get("mime_type"),
                            "bytes_in_export": False,
                        }, str(block.get("tool_use_id") or ""))
                continue
            if block.get("type") != "tool_use":
                continue
            name = str(block.get("name") or "")
            inp = block.get("input") or {}
            if not isinstance(inp, dict):
                continue
            if name in WRITE_TOOLS:
                raw_path = inp.get("file_path") or inp.get("path")
                content = inp.get("content") if "content" in inp else inp.get("file_text")
                if raw_path and content is not None:
                    key = path_key(str(raw_path))
                    virtual_files[key] = str(content)
                    path_origins.setdefault(key, []).append(f"{name} in message {message.get('uuid')}")
                    if message_date:
                        path_dates.setdefault(key, []).append(message_date)
            elif name in EDIT_TOOLS:
                raw_path = inp.get("file_path") or inp.get("path")
                old = inp.get("old_string") if "old_string" in inp else inp.get("old_str")
                new = inp.get("new_string") if "new_string" in inp else inp.get("new_str")
                if raw_path and old is not None and new is not None:
                    key = path_key(str(raw_path))
                    current = virtual_files.get(key)
                    if current is None or str(old) not in current:
                        edit_failures.append({"path": str(raw_path), "tool": name, "message_uuid": message.get("uuid"), "reason": "base file missing" if current is None else "old text not found", "replacement": str(new)})
                    else:
                        virtual_files[key] = current.replace(str(old), str(new), -1 if inp.get("replace_all") else 1)
                        path_origins.setdefault(key, []).append(f"{name} in message {message.get('uuid')}")
                        if message_date:
                            path_dates.setdefault(key, []).append(message_date)
            elif name in COMMAND_TOOLS:
                command = str(inp.get("command") or "")
                if not command:
                    continue
                command_count += 1
                commands_dir.mkdir(parents=True, exist_ok=True)
                command_path = commands_dir / f"{command_count:04d}_{safe_component(name, 30)}.sh"
                command_path.write_text(command.rstrip() + "\n", encoding="utf-8")
                cwd = initial_cwd(command)
                for heredoc in heredocs(command):
                    inline_count += 1
                    interpreter = heredoc["interpreter"]
                    suffix = {"python": ".py", "python3": ".py", "node": ".js", "bash": ".sh", "sh": ".sh"}.get(interpreter, ".txt")
                    inline_dir.mkdir(parents=True, exist_ok=True)
                    (inline_dir / f"{command_count:04d}_{inline_count:04d}_{safe_component(interpreter, 20)}{suffix}").write_text(heredoc["body"], encoding="utf-8")
                    if heredoc["target"]:
                        key = path_key(heredoc["target"], cwd)
                        if heredoc["append"] and key in virtual_files:
                            virtual_files[key] += heredoc["body"]
                        else:
                            virtual_files[key] = heredoc["body"]
                        path_origins.setdefault(key, []).append(f"heredoc in command {command_count}")
                        if message_date:
                            path_dates.setdefault(key, []).append(message_date)
            elif name in PRESENT_TOOLS:
                presented_paths = inp.get("filepaths") or inp.get("files") or []
                if inp.get("file_path"):
                    presented_paths = [*presented_paths, inp["file_path"]]
                for raw_path in presented_paths:
                    record_presented(
                        {"message_uuid": message.get("uuid"), "tool": name, "uuid": None, "file_path": raw_path, "name": Path(str(raw_path)).stem, "mime_type": None, "bytes_in_export": False},
                        str(block.get("id") or ""),
                    )

    recovered_manifest: list[dict[str, Any]] = []
    for key, content in sorted(virtual_files.items()):
        local = write_vfs(files_dir, key, content)
        dates = sorted(path_dates.get(key, []))
        recovered_manifest.append({
            "original_path": "/" + key,
            "local_path": local.relative_to(conv_dir).as_posix(),
            "bytes": len(content.encode("utf-8")),
            "sha256": hashlib.sha256(content.encode("utf-8")).hexdigest(),
            "origins": path_origins.get(key, []),
            "created_at": dates[0] if dates else created,
            "updated_at": dates[-1] if dates else created,
            "revisions": len(path_origins.get(key, [])),
        })

    json_dump(conv_dir / "recovered_files.json", recovered_manifest)
    json_dump(conv_dir / "presented_files.json", presented)
    pages = {}
    for lang in AVAILABLE_LANGS:
        chat_html = render_chat_html(conversation, recovered_manifest, presented, attachment_records, conv_dir, lang)
        _, pages[lang], _ = split_chat_page(chat_html)
    chat_data = {"title": conversation.get("name") or "Untitled", "pages": pages}
    (conv_dir / "chat-data.js").write_text(
        "window.restoredChatData = " + json.dumps(chat_data, ensure_ascii=True) + ";\n",
        encoding="utf-8",
    )
    if edit_failures:
        json_dump(conv_dir / "unapplied_edits.json", edit_failures)
    office_presented = sum(Path(str(item.get("file_path") or "")).suffix.lower() in OFFICE_EXTENSIONS for item in presented)
    readme = [
        f"# {conversation.get('name') or 'Untitled'}",
        "",
        f"- Conversation UUID: `{uuid}`",
        f"- Recovered text/source files: {len(recovered_manifest)}",
        f"- Attachments with extracted text: {attachment_count}",
        f"- Shell commands preserved: {command_count}",
        f"- Heredoc bodies preserved: {inline_count}",
        f"- Presented file records: {len(presented)} ({office_presented} office/PDF)",
        f"- Unapplied edits: {len(edit_failures)}",
        "",
        "`presented_files.json` records output files known to Claude. `bytes_in_export: false` means the export only retained metadata; use the recovered generators and command history to reproduce the file.",
    ]
    (conv_dir / "README.md").write_text("\n".join(readme) + "\n", encoding="utf-8")
    return {
        "created": created,
        "updated": str(conversation.get("updated_at") or created),
        "has_summary": bool(str(conversation.get("summary") or "").strip()),
        "title": conversation.get("name") or "Untitled",
        "uuid": uuid,
        "folder": dirname,
        "recovered_files": len(recovered_manifest),
        "presented_files": len(presented),
        "office_presented": office_presented,
        "commands": command_count,
        "inline_scripts": inline_count,
        "attachments": attachment_count,
        "unapplied_edits": len(edit_failures),
    }


def copy_export_sections(raw_dir: Path, output: Path) -> None:
    # The complete original sections already live under output/raw.  Do not
    # duplicate the large artifact tree: on drvfs, copytree's directory
    # metadata update can also fail even though file writes are allowed.
    projects_dir = raw_dir / "projects"
    if projects_dir.exists():
        docs_root = output / "project_documents"
        for project_file in sorted(projects_dir.glob("*.json")):
            project = json.loads(project_file.read_text(encoding="utf-8"))
            project_dir = docs_root / f"{safe_component(project.get('name'), 80)}_{str(project.get('uuid', ''))[:8]}"
            json_dump(project_dir / "project.json", project)
            for index, doc in enumerate(project.get("docs") or [], 1):
                if not isinstance(doc, dict) or doc.get("content") is None:
                    continue
                filename = doc.get("filename") or f"document_{index}.txt"
                target = project_dir / "docs" / Path(*[safe_component(part, 120) for part in PurePosixPath(str(filename)).parts if part not in {"/", "..", "."}])
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_text(str(doc["content"]), encoding="utf-8")

    artifacts_dir = raw_dir / "artifacts"
    if artifacts_dir.exists():
        cards: list[str] = []
        for artifact_dir in sorted(path for path in artifacts_dir.iterdir() if path.is_dir()):
            metadata_file = artifact_dir / "artifact.json"
            title = artifact_dir.name
            kind = "artifact"
            active = ""
            if metadata_file.exists():
                metadata = json.loads(metadata_file.read_text(encoding="utf-8"))
                kind = str(metadata.get("kind") or kind)
                active = str(metadata.get("active_version") or "")
                versions = metadata.get("versions") or []
                if versions:
                    title = str(versions[-1].get("title") or title)
            candidates = [artifact_dir / "page.md"]
            if active:
                candidates.extend(sorted((artifact_dir / "versions").glob(f"{active}*")) if (artifact_dir / "versions").exists() else [])
            target = next((path for path in candidates if path.exists()), metadata_file)
            target_rel = target.relative_to(artifacts_dir).as_posix() if target else ""
            cards.append(f'<article><h2><a href="{href(target_rel)}">{html.escape(title)}</a></h2><div>{html.escape(kind)} · <code>{html.escape(artifact_dir.name)}</code></div></article>')
        artifact_html = f"""<!doctype html><html lang="{L("html_lang", DEFAULT_LANG)}"><head><meta charset="utf-8"><title>{html.escape(L("artifacts_index.page_title", DEFAULT_LANG))}</title><style>body{{max-width:1000px;margin:35px auto;padding:0 20px;background:#f5f3ee;font:15px/1.5 system-ui;color:#24231f}}article{{background:white;border:1px solid #ddd8ce;border-radius:12px;padding:14px 18px;margin:10px 0}}h2{{font-size:18px;margin:0 0 4px}}a{{color:#0758a8}}code,article div{{color:#706d65}}</style></head><body><h1>{html.escape(L("artifacts_index.heading", DEFAULT_LANG))}</h1>""" + "".join(cards) + "</body></html>"
        (artifacts_dir / "INDEX.html").write_text(artifact_html, encoding="utf-8")


def restore_export(source: Path, output: Path) -> dict[str, Any]:
    """Restore one export and return its machine-readable summary."""
    source = source.resolve()
    output = output.resolve()
    raw_dir = output / "raw"
    conversations_dir = output / "conversations"
    output.mkdir(parents=True, exist_ok=True)

    archive_report = unpack_archives(source, raw_dir)
    json_dump(output / "archive_report.json", archive_report)
    errors = [entry for entry in archive_report if entry["error"]]
    conversations_file = raw_dir / "conversations.json"
    if errors:
        print(f"warning: {len(errors)} archive(s) failed validation", file=sys.stderr)
    if not conversations_file.exists():
        raise FileNotFoundError(conversations_file)

    conversations = json.loads(conversations_file.read_text(encoding="utf-8"))
    rows = [restore_conversation(conv, conversations_dir) for conv in conversations]
    rows.sort(key=lambda row: (row["created"], row["title"]))
    json_dump(output / "restoration_report.json", rows)
    (output / "INDEX.md").write_text(render_index(rows), encoding="utf-8")
    for lang in AVAILABLE_LANGS:
        (output / f"viewer{lang_suffix(lang)}.html").write_text(render_chat_viewer(lang), encoding="utf-8")
        (output / f"INDEX{lang_suffix(lang)}.html").write_text(render_index_html(rows, lang), encoding="utf-8")
    copy_export_sections(raw_dir, output)

    totals = Counter()
    for row in rows:
        for key in ("recovered_files", "presented_files", "office_presented", "commands", "inline_scripts", "attachments", "unapplied_edits"):
            totals[key] += row[key]
    summary = {
        "conversations": len(rows),
        **totals,
        "archives": len(archive_report),
        "archive_errors": len(errors),
    }
    json_dump(output / "SUMMARY.json", summary)
    return summary


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path.cwd(), help="directory containing export ZIP files")
    parser.add_argument("--output", type=Path, default=Path("restored_export"), help="restoration destination")
    args = parser.parse_args()
    try:
        summary = restore_export(args.source, args.output)
    except FileNotFoundError as exc:
        print(f"error: {exc} not found", file=sys.stderr)
        return 2
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
