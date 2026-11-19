# AGENTS.md

> **Read [`RESTORE_NEW_CLAUDE_EXPORT.md`](RESTORE_NEW_CLAUDE_EXPORT.md) before doing any work.**
> Chinese version: [`RESTORE_NEW_CLAUDE_EXPORT_zh.md`](RESTORE_NEW_CLAUDE_EXPORT_zh.md)

## What this repo does

Restores Claude export archives (ZIP + manifest) into browsable
per-conversation HTML pages with regenerated PPTX/DOCX/XLSX files,
media assets (PNG/GIF/SVG/HTML), and a multi-language UI.

## Scripts

| Script | Purpose |
|---|---|
| `restore_pipeline.py` | Preferred entry point: restore → dependency-aware regeneration → timestamps → HTML → validation |
| `restore_export.py` | Main restorer: unpacks ZIP, builds per-conversation HTML (all languages), INDEX |
| `regenerate_python_office.py` | Runs recovered Python generators: Office (openpyxl, python-docx, python-pptx) + media (matplotlib, PIL, cv2, HTML builders) |
| `regenerate_node_office.py` | Runs recovered Node.js generators (pptxgenjs, docx) |
| `regenerate_recipe_docx.py` | Replays multi-step DOCX recipes (TeX → Pandoc → python-docx) |
| `regenerate_known_office.py` | Replays audited DOCX/XLSX recipes and the CPU frequency HTML/PDF recipe (Pandoc, wkhtmltopdf, Poppler) |
| `artifact_timestamps.py` | Restores regenerated artifact mtimes from presentation/tool-call logs |
| `check_restoration.py` | Validates local links and OOXML ZIP integrity |
| `winenv.py` | Cross-platform helpers: winget tool discovery, `to_posix_path()`, `kill_orphaned_tex_processes()` |
| `install_deps_windows.ps1` | Silent-installs all dependencies via winget + pip |
| `restore_all.bat` | Windows wrapper around `restore_pipeline.py` |

Shared importable pipeline code lives in `pipeline/`. Standalone scripts remain
supported CLI wrappers around the same stage functions.

Pipeline state is an atomic execution record, not a content-addressed cache.
Use `--from-stage` for an explicit restart. Do not infer freshness by comparing
artifact mtimes: regenerated artifacts intentionally receive historical mtimes
from the logs. Hashing is optional future work, not required for correctness.

## i18n

- Locale files: `locales/*.json` (ru, zh, en)
- Default language: `ru` → `viewer.html`, `INDEX.html`
- Other languages: `viewer_zh.html`, `INDEX_en.html`, etc.
- Each conversation provides `chat-data.js` for the shared viewer, including under `file://`.
- Language switcher: sets `restore_lang` cookie, preserves selected chat via `?chat=<folder>` URL param
- Chat scroll position preserved across language switches
- To add a language: drop `locales/<code>.json`, re-run `restore_export.py`

## UI features

- Dark/light theme with localStorage persistence
- Resizable sidebar and artifact panels
- Inline `.md` preview with table/code/heading support
- Tool-call blocks collapsed by default, compact styling
- Empty Thinking blocks hidden
- Copy chat as Markdown button

## Critical rules

1. **Never reuse `--output` for different exports.** Each export gets a fresh directory.
2. **Never substitute backslash paths into recovered source code.** Use `winenv.to_posix_path()`.
3. **Never use `python3` in subprocess calls.** Use `sys.executable`.
4. **Regeneration runs untrusted recovered code.** Read section 4 of the instructions.
5. Prefer `restore_pipeline.py`; it automatically re-renders HTML after regeneration.
6. The pipeline must finish with `broken_links` and `corrupt_office_files` equal to `0`.
