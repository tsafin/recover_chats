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
            "previewable": path.suffix.lower() in {".pdf", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".html", ".htm", ".txt", ".md", ".json", ".csv"},
        })

    parts = [f"""<!doctype html>
<html lang="{L("html_lang", lang)}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
:root{{--bg:#f5f3ee;--paper:#fff;--ink:#24231f;--muted:#706d65;--line:#ddd8ce;--user:#e8f1ff;--ai:#fff8e8;--accent:#b44b2a;--code:#202124}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--ink);font:15px/1.55 system-ui,-apple-system,"Segoe UI",sans-serif}}
a{{color:#0758a8}}.layout{{display:grid;grid-template-columns:minmax(0,1fr) 340px;gap:20px;max-width:1500px;margin:auto;padding:20px}}
main{{min-width:0}}.hero,.message,.panel{{background:var(--paper);border:1px solid var(--line);border-radius:14px;box-shadow:0 2px 8px #0000000a}}
.hero{{padding:22px;margin-bottom:18px}}.hero h1{{margin:0 0 8px;font-size:26px}}.meta{{color:var(--muted)}}
.message{{padding:18px;margin:14px 0}}.message.user{{border-left:5px solid #3478c8;background:linear-gradient(90deg,var(--user),#fff 20%)}}
.message.assistant{{border-left:5px solid #d18a24;background:linear-gradient(90deg,var(--ai),#fff 20%)}}
.head{{display:flex;justify-content:space-between;gap:12px;margin-bottom:12px}}.role{{font-weight:750}}.time{{color:var(--muted);font-size:12px}}.text{{white-space:pre-wrap;overflow-wrap:anywhere}}.markdown{{overflow-wrap:anywhere}}.markdown>:first-child{{margin-top:0}}.markdown>:last-child{{margin-bottom:0}}.markdown h1,.markdown h2,.markdown h3{{line-height:1.25;margin:1.1em 0 .45em}}.markdown h1{{font-size:1.55em}}.markdown h2{{font-size:1.3em}}.markdown h3{{font-size:1.12em}}.markdown blockquote{{margin:10px 0;padding:3px 14px;border-left:4px solid #b9b4aa;color:#555;background:#f7f6f3}}.markdown table{{border-collapse:collapse;display:block;max-width:100%;overflow:auto}}.markdown th,.markdown td{{border:1px solid var(--line);padding:6px 10px}}.markdown th{{background:#f1efe9}}.markdown img{{max-width:100%}}.markdown ul,.markdown ol{{padding-left:25px}}
details{{border:1px solid var(--line);border-radius:9px;margin:10px 0;background:#fafafa}}summary{{cursor:pointer;padding:9px 12px;font-weight:650}}details>div,details>pre{{margin:0;padding:12px;border-top:1px solid var(--line)}}
pre{{white-space:pre-wrap;overflow-wrap:anywhere;background:var(--code);color:#e8eaed;font:12px/1.45 ui-monospace,SFMono-Regular,Consolas,monospace}}
.chips{{display:flex;flex-wrap:wrap;gap:7px;margin:8px 0}}.chip,.file-link{{display:inline-block;border:1px solid #b7cce4;border-radius:999px;padding:3px 9px;background:#f3f8ff;text-decoration:none;font-size:12px}}
.resource{{border-left:3px solid #6f9d68;padding:8px 10px;margin:7px 0;background:#f4faf2}}.missing{{border-color:#c8a34c;background:#fff9e9}}.source-reference{{color:#6b6255;text-decoration-style:dotted}}
aside{{position:sticky;top:12px;height:calc(100vh - 24px);overflow:auto}}.panel{{padding:15px}}.panel h2{{font-size:15px;margin:15px 0 7px}}.panel h2:first-child{{margin-top:0}}.panel ul{{padding-left:20px;margin:5px 0}}.panel li{{margin:5px 0;overflow-wrap:anywhere}}.count{{color:var(--muted);font-size:12px}}
.artifact-toolbar{{display:flex;gap:7px;align-items:center;margin:8px 0 11px}}.artifact-toolbar select{{min-width:0;flex:1;border:1px solid var(--line);border-radius:8px;background:white;padding:6px}}.artifact-list,.file-list{{display:grid;gap:8px}}.artifact-card,.file-card{{border:1px solid var(--line);border-radius:10px;padding:10px;background:#fbfaf7}}.artifact-name,.file-name{{font-weight:650;overflow-wrap:anywhere}}.artifact-meta,.file-meta{{color:var(--muted);font-size:12px;margin:3px 0 7px}}.artifact-actions{{display:flex;gap:7px;flex-wrap:wrap}}.artifact-actions a,.artifact-actions button,.file-card a{{border:1px solid #b7cce4;border-radius:7px;padding:3px 7px;background:#f3f8ff;color:#0758a8;font:inherit;font-size:12px;text-decoration:none;cursor:pointer}}.artifact-missing{{color:#936b16;font-size:12px}}.preview-box{{display:none;margin-top:12px;border-top:1px solid var(--line);padding-top:12px}}.preview-box.open{{display:block}}.preview-head{{display:flex;justify-content:space-between;gap:8px;align-items:center;margin-bottom:7px}}.preview-box iframe{{width:100%;height:48vh;border:1px solid var(--line);border-radius:8px;background:white}}.preview-close{{border:0;background:transparent;font-size:18px;cursor:pointer}}
.lang-switch{{position:absolute;top:18px;right:30px;border:1px solid var(--line);border-radius:8px;background:white;padding:4px 6px;font:inherit;font-size:12px}}
@media(max-width:900px){{.layout{{grid-template-columns:1fr;padding:12px}}aside{{position:static;height:auto;order:-1}}.lang-switch{{right:14px}}}}
</style><title>""", html.escape(title), "</title></head><body><div class=\"layout\"><main>"]
    parts.extend([
        '<section class="hero">', f"<h1>{html.escape(title)}</h1>",
        f'<div class="meta">{t("uuid_label")}: <code>{html.escape(str(conversation.get("uuid") or ""))}</code><br>{t("created_label")}: {html.escape(str(conversation.get("created_at") or ""))}</div>',
        f'<div class="chips"><a class="chip" href="README.md">{t("link_readme")}</a><a class="chip" href="chat.md">{t("link_markdown")}</a><a class="chip" href="conversation.json">{t("link_json")}</a><a class="chip" href="recovered_files.json">{t("link_files_json")}</a><a class="chip" href="presented_files.json">{t("link_artifacts_json")}</a><a class="chip" href="../../raw/artifacts/">{t("link_all_artifacts")}</a></div>',
        "</section>",
    ])
    command_index = 0
    for msg_index, message in enumerate(conversation.get("chat_messages") or [], 1):
        sender = message.get("sender")
        role = t("role_user") if sender == "human" else t("role_assistant")
        css_role = "user" if sender == "human" else "assistant"
        parts.append(f'<article class="message {css_role}" id="message-{msg_index}"><div class="head"><span class="role">{role}</span><span class="time">{html.escape(str(message.get("created_at") or ""))}</span></div>')
        for attachment in attachments_by_message.get(str(message.get("uuid")), []):
            parts.append(f'<div class="resource">📎 <a href="{href(str(attachment["local_path"]))}">{html.escape(str(attachment["name"]))}</a> <span class="count">{attachment.get("bytes", 0)} bytes</span></div>')
        blocks = message.get("content") or []
        if not blocks and message.get("text"):
            parts.append(f'<div class="markdown">{render_markdown(message["text"], resource_links, lang)}</div>')
        for block in blocks:
            if not isinstance(block, dict):
                continue
            kind = block.get("type")
            if kind == "text":
                parts.append(f'<div class="markdown">{render_markdown(block.get("text"), resource_links, lang)}</div>')
            elif kind == "thinking":
                parts.append(f'<details><summary>{t("thinking")}</summary><div class="markdown">{render_markdown(block.get("thinking"), resource_links, lang)}</div></details>')
            elif kind == "tool_use":
                name = str(block.get("name") or "tool")
                inp = block.get("input") or {}
                chips: list[str] = []
                if name in COMMAND_TOOLS:
                    command_index += 1
                    chips.append(f'<a class="chip" href="commands/{command_index:04d}_{safe_component(name, 30)}.sh">{t("command_label")} #{command_index}</a>')
                if isinstance(inp, dict):
                    candidate_paths = [inp.get("file_path"), inp.get("path")]
                    candidate_paths.extend(inp.get("filepaths") or [])
                    for raw_path in candidate_paths:
                        if raw_path and str(raw_path) in resource_links:
                            chips.append(f'<a class="chip" href="{href(resource_links[str(raw_path)])}">📄 {html.escape(Path(str(raw_path)).name)}</a>')
                payload = html.escape(json.dumps(inp, ensure_ascii=False, indent=2))
                parts.append(f'<details><summary>🛠 {html.escape(name)}</summary><div class="chips">{"".join(chips)}</div><pre>{payload}</pre></details>')
            elif kind == "tool_result":
                resources = []
                texts = []
                for item in block.get("content") or []:
                    if not isinstance(item, dict):
                        continue
                    if item.get("type") == "local_resource":
                        raw_path = str(item.get("file_path") or "")
                        local = resource_links.get(raw_path)
                        if local:
                            resources.append(f'<div class="resource">✅ <a href="{href(local)}">{html.escape(Path(raw_path).name)}</a> · {html.escape(str(item.get("mime_type") or ""))}</div>')
                        else:
                            resources.append(f'<div class="resource missing">{t("metadata_only")}: {html.escape(raw_path)} · UUID <code>{html.escape(str(item.get("uuid") or ""))}</code></div>')
                    elif item.get("type") == "text":
                        texts.append(render_markdown(item.get("text"), resource_links, lang))
                if resources or texts:
                    parts.append(f'<details><summary>{t("result_prefix")}: {html.escape(str(block.get("name") or "tool"))}</summary>{"".join(resources)}')
                    for text_item in texts:
                        parts.append(f'<div class="markdown">{text_item}</div>')
                    parts.append("</details>")
            elif kind in {"image", "document"}:
                parts.append(f'<div class="resource missing">{html.escape(kind.title())}: UUID <code>{html.escape(str(block.get("file_uuid") or ""))}</code> — {t("binary_missing")}</div>')
        parts.append("</article>")
    parts.append("</main><aside><section class=\"panel\">")
    parts.append(f'<h2>{t("artifacts_title")} <span class="count">{len(artifact_rows)}</span></h2>')
    parts.append(f'<div class="artifact-toolbar"><label for="artifact-sort" class="count">{t("sort_label")}</label><select id="artifact-sort"><option value="newest">{t("sort_newest")}</option><option value="oldest">{t("sort_oldest")}</option><option value="name">{t("sort_name")}</option><option value="version">{t("sort_version")}</option></select></div><div class="artifact-list" id="artifact-list">')
    for item in artifact_rows:
        local = str(item["local"])
        date = str(item["date"])
        visible_date = date.replace("T", " ")[:16] if date else t("date_unknown")
        parts.append(f'<article class="artifact-card" data-date="{html.escape(date, quote=True)}" data-name="{html.escape(str(item["name"]).casefold(), quote=True)}" data-version="{item["version"]}"><div class="artifact-name">{html.escape(str(item["name"]))}</div><div class="artifact-meta">{html.escape(visible_date)} · {t("version_label")} {item["version"]}</div>')
        if local:
            parts.append(f'<div class="artifact-actions"><a href="{href(local)}" target="_blank">{t("button_open")}</a>')
            if item["previewable"]:
                parts.append(f'<button type="button" class="preview-button" data-src="{href(local)}" data-title="{html.escape(str(item["name"]), quote=True)}">{t("button_preview")}</button>')
            parts.append('</div>')
        else:
            parts.append(f'<div class="artifact-missing" title="">⚠ {t("metadata_missing_warning")}</div>')
        parts.append('</article>')
    parts.append(f'</div><div class="preview-box" id="preview-box"><div class="preview-head"><strong id="preview-title">{t("preview_title")}</strong><button class="preview-close" id="preview-close" title="{t("preview_close")}">×</button></div><iframe id="artifact-preview" title="{t("preview_iframe_title")}"></iframe></div>')
    parts.append(f'<details><summary>{t("regenerated_title")} <span class="count">{len(regenerated)}</span></summary><ul>')
    for path in regenerated:
        if path.is_file():
            parts.append(f'<li><a href="{href(path.relative_to(conv_dir).as_posix())}" target="_blank">{html.escape(path.name)}</a></li>')
    parts.append("</ul></details>")
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
const artifactList=document.getElementById('artifact-list');
const artifactSort=document.getElementById('artifact-sort');
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
const previewBox=document.getElementById('preview-box'), preview=document.getElementById('artifact-preview'), previewTitle=document.getElementById('preview-title');
document.querySelectorAll('.preview-button').forEach(button=>button.addEventListener('click',()=>{preview.src=button.dataset.src;previewTitle.textContent=button.dataset.title;previewBox.classList.add('open');previewBox.scrollIntoView({behavior:'smooth',block:'start'});}));
document.getElementById('preview-close').addEventListener('click',()=>{preview.removeAttribute('src');previewBox.classList.remove('open');});
</script></body></html>""".replace("__SORT_LOCALE__", sort_locale)
    parts.append(js_block)
    return "".join(parts)


def render_index(rows: list[dict[str, Any]]) -> str:
    lines = ["# Restored Claude conversations", "", "The export contains metadata, chats, extracted attachment text, and generator sources. Presented binary files are usually metadata-only in the export; see each conversation's `presented_files.json`.", "", "| Created | Conversation | Recovered files | Presented files |", "|---|---|---:|---:|"]
    for row in rows:
        title = str(row["title"]).replace("|", "\\|")
        folder = str(row["folder"]).replace(" ", "%20")
        lines.append(f"| {row['created'][:10]} | [{title}]({folder}/chat.md) | {row['recovered_files']} | {row['presented_files']} |")
    return "\n".join(lines) + "\n"


def render_index_html(rows: list[dict[str, Any]], lang: str = DEFAULT_LANG) -> str:
    ti = lambda key: L(f"index.{key}", lang)
    sort_locale = L("sort_locale", lang)
    cards = []
    newest_rows = sorted(rows, key=lambda item: (str(item.get("updated") or item["created"]), str(item["title"])), reverse=True)
    for row in newest_rows:
        activity_date = str(row.get("updated") or row["created"])
        suffix = lang_suffix(lang)
        cards.append(
            f'<a class="conversation" href="conversations/{href(str(row["folder"]))}/chat{suffix}.html" target="chat-frame" data-date="{html.escape(activity_date, quote=True)}" data-title="{html.escape(str(row["title"]).casefold(), quote=True)}"><strong>{html.escape(str(row["title"]))}</strong>'
            f'<span>{html.escape(activity_date.replace("T", " ")[:16])} · {row["presented_files"]} {ti("card_artifacts")} · {row["recovered_files"]} {ti("card_files")}</span></a>'
        )
    first_url = f'conversations/{href(str(newest_rows[0]["folder"]))}/chat{lang_suffix(lang)}.html' if newest_rows else ""

    # Language switcher options
    lang_options = []
    for code in AVAILABLE_LANGS:
        name = LOCALES[code].get("lang_name", code)
        sel = " selected" if code == lang else ""
        lang_options.append(f'<option value="{code}"{sel}>{html.escape(name)}</option>')

    return f"""<!doctype html><html lang="{L("html_lang", lang)}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(ti("page_title"))}</title><style>
:root{{--sidebar:#181817;--sidebar2:#242422;--ink:#eee;--muted:#a9a7a1;--accent:#d97757}}*{{box-sizing:border-box}}html,body{{height:100%;margin:0;font:14px/1.4 system-ui,-apple-system,"Segoe UI",sans-serif}}body{{display:grid;grid-template-columns:290px minmax(0,1fr);overflow:hidden;background:#f5f3ee}}.sidebar{{height:100vh;background:var(--sidebar);color:var(--ink);display:flex;flex-direction:column;border-right:1px solid #343431}}.brand{{padding:15px 16px 10px;font-size:20px;font-weight:700}}.tools{{padding:0 10px 10px;display:grid;gap:7px}}.tools input,.tools select{{width:100%;border:1px solid #454541;border-radius:8px;background:var(--sidebar2);color:var(--ink);padding:8px 9px}}.links{{display:flex;gap:9px;padding:0 14px 10px;flex-wrap:wrap;align-items:center}}.links a{{color:#c9c6bf;font-size:12px}}.lang-switch{{border:1px solid #454541;border-radius:8px;background:var(--sidebar2);color:var(--ink);padding:4px 6px;font:inherit;font-size:12px}}.conversation-list{{overflow:auto;padding:4px 7px 16px}}.conversation{{display:block;color:var(--ink);text-decoration:none;padding:9px 10px;border-radius:8px;margin:2px 0}}.conversation:hover,.conversation.active{{background:#3a3a37}}.conversation.active{{box-shadow:inset 3px 0 var(--accent)}}.conversation strong{{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-weight:560}}.conversation span{{display:block;color:var(--muted);font-size:11px;margin-top:3px}}.empty{{color:var(--muted);padding:15px;display:none}}iframe{{width:100%;height:100vh;border:0;background:#f5f3ee}}@media(max-width:720px){{body{{grid-template-columns:1fr;overflow:auto}}.sidebar{{height:46vh}}.viewer{{height:54vh}}.viewer iframe{{height:54vh}}}}
</style></head><body><nav class="sidebar"><div class="brand">{html.escape(ti("brand"))}</div><div class="tools"><input id="search" type="search" placeholder="{html.escape(ti("search_placeholder"))}"><select id="sort"><option value="newest">{html.escape(ti("sort_newest"))}</option><option value="oldest">{html.escape(ti("sort_oldest"))}</option><option value="title">{html.escape(ti("sort_title"))}</option></select></div><div class="links"><a href="RESTORATION_NOTES.md" target="_blank">{html.escape(ti("link_description"))}</a><a href="raw/artifacts/INDEX.html" target="_blank">{html.escape(ti("link_artifacts"))}</a><a href="project_documents/" target="_blank">{html.escape(ti("link_projects"))}</a><a href="SUMMARY.json" target="_blank">{html.escape(ti("link_summary"))}</a><select class="lang-switch" id="lang-switch">{"".join(lang_options)}</select></div><div class="conversation-list" id="conversation-list">""" + "".join(cards) + f"""<div class="empty" id="empty">{html.escape(ti("empty"))}</div></div></nav><main class="viewer"><iframe name="chat-frame" id="chat-frame" src="__FIRST_CHAT__" title="{html.escape(ti("iframe_title"))}"></iframe></main><script>
const list=document.getElementById('conversation-list'), search=document.getElementById('search'), sort=document.getElementById('sort');
const conversations=[...list.querySelectorAll('.conversation')];
function update(){{const q=search.value.trim().toLocaleLowerCase('{sort_locale}'); conversations.forEach(x=>x.hidden=q&&!x.dataset.title.includes(q)); const visible=conversations.filter(x=>!x.hidden); visible.sort((a,b)=>sort.value==='title'?a.dataset.title.localeCompare(b.dataset.title,'{sort_locale}'):(sort.value==='oldest'?a.dataset.date.localeCompare(b.dataset.date):b.dataset.date.localeCompare(a.dataset.date))).forEach(x=>list.appendChild(x)); document.getElementById('empty').style.display=visible.length?'none':'block';}}
search.addEventListener('input',update);sort.addEventListener('change',update);update();
conversations.forEach(link=>link.addEventListener('click',()=>{{conversations.forEach(x=>x.classList.remove('active'));link.classList.add('active');}}));
if(conversations.length) conversations[0].classList.add('active');
// Language switcher: set cookie and redirect to the localized index
document.getElementById('lang-switch').addEventListener('change',function(){{
  var lang=this.value;
  document.cookie='restore_lang='+lang+';path=/;max-age=31536000';
  var suffix=lang==='{DEFAULT_LANG}'?'':'_'+lang;
  var base=document.location.pathname.replace(/(_(en|zh|ru))?\.html$/,'');
  document.location.href=base+suffix+'.html';
}});
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
    command_count = 0
    inline_count = 0
    attachment_count = 0
    attachment_records: list[dict[str, Any]] = []

    for msg_index, message in enumerate(conversation.get("chat_messages") or [], 1):
        message_date = str(message.get("created_at") or "")
        for file_index, attachment in enumerate(message.get("attachments") or [], 1):
            if not isinstance(attachment, dict) or attachment.get("extracted_content") is None:
                continue
            attachment_count += 1
            name = attachment.get("file_name") or f"message_{msg_index:04d}_attachment_{file_index}"
            # `extracted_content` is always text, even when the original upload
            # was DOCX/PDF.  Keep that distinction explicit so the recovered
            # text is not mistaken for the unavailable source binary.
            target = attachments_dir / f"message_{msg_index:04d}_{file_index:02d}_{safe_component(Path(str(name)).stem)}.extracted.txt"
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(str(attachment["extracted_content"]), encoding="utf-8")
            attachment_records.append({
                "message_uuid": message.get("uuid"),
                "name": name,
                "local_path": target.relative_to(conv_dir).as_posix(),
                "bytes": len(str(attachment["extracted_content"]).encode("utf-8")),
            })

        for block in message.get("content") or []:
            if not isinstance(block, dict):
                continue
            if block.get("type") == "tool_result":
                for item in block.get("content") or []:
                    if isinstance(item, dict) and item.get("type") == "local_resource":
                        presented.append({
                            "message_uuid": message.get("uuid"),
                            "tool": block.get("name"),
                            "uuid": item.get("uuid"),
                            "file_path": item.get("file_path"),
                            "name": item.get("name"),
                            "mime_type": item.get("mime_type"),
                            "bytes_in_export": False,
                        })
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
                    if not any(item.get("file_path") == raw_path for item in presented):
                        presented.append({"message_uuid": message.get("uuid"), "tool": name, "uuid": None, "file_path": raw_path, "name": Path(str(raw_path)).stem, "mime_type": None, "bytes_in_export": False})

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
    for lang in AVAILABLE_LANGS:
        chat_html = render_chat_html(conversation, recovered_manifest, presented, attachment_records, conv_dir, lang)
        (conv_dir / f"chat{lang_suffix(lang)}.html").write_text(chat_html, encoding="utf-8")
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


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path.cwd(), help="directory containing export ZIP files")
    parser.add_argument("--output", type=Path, default=Path("restored_export"), help="restoration destination")
    args = parser.parse_args()
    source = args.source.resolve()
    output = args.output.resolve()
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
        print(f"error: {conversations_file} not found", file=sys.stderr)
        return 2

    conversations = json.loads(conversations_file.read_text(encoding="utf-8"))
    rows = [restore_conversation(conv, conversations_dir) for conv in conversations]
    rows.sort(key=lambda row: (row["created"], row["title"]))
    json_dump(output / "restoration_report.json", rows)
    (output / "INDEX.md").write_text(render_index(rows), encoding="utf-8")
    for lang in AVAILABLE_LANGS:
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
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
