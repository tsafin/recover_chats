# AGENTS.md

> **Read [`RESTORE_NEW_CLAUDE_EXPORT.md`](RESTORE_NEW_CLAUDE_EXPORT.md) before doing any work.**
> Chinese version: [`RESTORE_NEW_CLAUDE_EXPORT_zh.md`](RESTORE_NEW_CLAUDE_EXPORT_zh.md)

## What this repo does

Restores Claude export archives (ZIP + manifest) into browsable
per-conversation HTML pages with regenerated PPTX/DOCX/XLSX files.

## Scripts

| Script | Purpose |
|---|---|
| `restore_export.py` | Main restorer: unpacks ZIP, builds per-conversation HTML, INDEX |
| `regenerate_python_office.py` | Runs recovered Python generators (openpyxl, python-docx, python-pptx) |
| `regenerate_node_office.py` | Runs recovered Node.js generators (pptxgenjs, docx) |
| `regenerate_recipe_docx.py` | Replays multi-step DOCX recipes (TeX → Pandoc → python-docx) |
| `regenerate_known_office.py` | Replays explicitly audited DOCX/XLSX recipes for this export |
| `check_restoration.py` | Validates local links and OOXML ZIP integrity |
| `winenv.py` | Cross-platform helpers: winget tool discovery, `to_posix_path()`, `kill_orphaned_tex_processes()` |
| `install_deps_windows.ps1` | Silent-installs all dependencies via winget + pip |
| `restore_all.bat` | One-command Windows pipeline |

## i18n

- Locale files: `locales/*.json` (ru, zh, en)
- Default language: `ru` → `chat.html`, `INDEX.html`
- Other languages: `chat_zh.html`, `INDEX_en.html`, etc.
- Language switcher sets `restore_lang` cookie (1-year expiry)
- To add a language: drop `locales/<code>.json`, re-run `restore_export.py`

## Critical rules

1. **Never reuse `--output` for different exports.** Each export gets a fresh directory.
2. **Never substitute backslash paths into recovered source code.** Use `winenv.to_posix_path()`.
3. **Never use `python3` in subprocess calls.** Use `sys.executable`.
4. **Regeneration runs untrusted recovered code.** Read section 4 of the instructions.
5. **After regeneration, re-run `restore_export.py`** to link files into HTML.
6. **Always run `check_restoration.py` at the end.** `broken_links` and `corrupt_office_files` must be `0`.
