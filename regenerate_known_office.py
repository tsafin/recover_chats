"""Run explicitly configured document recipes from local conversation folders."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from artifact_timestamps import restore_artifact_mtime
from winenv import ensure_tools_on_path


def run(command: list[str], cwd: Path) -> None:
    result = subprocess.run(command, cwd=cwd, capture_output=True, text=True, timeout=180)
    if result.returncode:
        raise RuntimeError(result.stderr[-2000:] or result.stdout[-2000:])


def execute_recipe(conversation: Path, recipe: dict, output_dir: Path) -> Path:
    kind = recipe["kind"]
    source = (conversation / recipe["source"]).resolve(strict=True)
    name = recipe["output"]
    if not source.is_relative_to(conversation) or not source.is_file():
        raise ValueError("source must be a file inside the conversation")
    if (not isinstance(name, str) or Path(name).name != name
            or any(char in name for char in "\\/:") or name in {"", ".", ".."}):
        raise ValueError("output must be a plain filename")
    suffix = Path(name).suffix.lower()
    allowed = {"copy": {".docx", ".xlsx", ".pptx", ".html", ".pdf"},
               "pandoc-docx": {".docx"}, "pandoc-html": {".html"},
               "latex-pdf": {".pdf"}, "html-pdf": {".pdf"}}
    if suffix not in allowed.get(kind, set()):
        raise ValueError("unsupported recipe or output extension")
    with tempfile.TemporaryDirectory(prefix="restore-document-") as temporary:
        work = Path(temporary)
        output = work / name
        if kind == "copy":
            if source.suffix.lower() != suffix:
                raise ValueError("copy requires matching extensions")
            shutil.copyfile(source, output)
        elif kind.startswith("pandoc-"):
            if source.suffix.lower() not in {".md", ".markdown", ".tex"}:
                raise ValueError("Pandoc input must be Markdown or TeX")
            options = ["--standalone", "--self-contained"] if kind == "pandoc-html" else []
            run(["pandoc", str(source), *options, "-o", str(output)], source.parent)
        elif kind == "latex-pdf":
            if source.suffix.lower() != ".tex":
                raise ValueError("pdfLaTeX input must be TeX")
            command = ["pdflatex", "-no-shell-escape", "-halt-on-error", "-interaction=nonstopmode",
                       "-jobname=compiled", "-output-directory=" + str(work), str(source)]
            for _ in range(2):
                run(command, source.parent)
            shutil.copyfile(work / "compiled.pdf", output)
        elif kind == "html-pdf":
            if source.suffix.lower() not in {".html", ".htm"}:
                raise ValueError("wkhtmltopdf input must be HTML")
            run(["wkhtmltopdf", "--disable-javascript", "--disable-local-file-access",
                 "--encoding", "utf-8", str(source), str(output)], source.parent)
        if suffix in {".docx", ".xlsx", ".pptx"}:
            with zipfile.ZipFile(output) as archive:
                if "[Content_Types].xml" not in archive.namelist() or archive.testzip() is not None:
                    raise ValueError("Invalid Office container")
        elif suffix == ".pdf":
            run(["pdfinfo", str(output)], work)
        elif not output.stat().st_size:
            raise ValueError("Empty HTML output")
        target = output_dir / name
        if target.is_symlink():
            raise ValueError("output must not be a symbolic link")
        shutil.copyfile(output, target)
    return target


def regenerate_local(root: Path, *, docx: bool) -> tuple[list[dict], dict]:
    ensure_tools_on_path()
    root = root.resolve()
    report = []
    for manifest in sorted(root.glob("*/recipes.local.json")):
        conversation = manifest.parent.resolve()
        if conversation.parent != root or manifest.is_symlink():
            continue
        output_dir = conversation / "regenerated"
        try:
            recipes = json.loads(manifest.read_text(encoding="utf-8"))
            if not isinstance(recipes, list) or not all(isinstance(item, dict) for item in recipes):
                raise ValueError("Expected a list of recipe objects")
            if output_dir.is_symlink():
                raise ValueError("regenerated must not be a symbolic link")
        except (OSError, ValueError) as error:
            report.append({"conversation": conversation.name, "status": "failed", "error": str(error)})
            continue
        for recipe in recipes:
            if (recipe.get("kind") == "pandoc-docx") != docx:
                continue
            item = {"conversation": conversation.name, "recipe": recipe.get("kind"),
                    "output": recipe.get("output")}
            try:
                output_dir.mkdir(exist_ok=True)
                output = execute_recipe(conversation, recipe, output_dir)
                item.update(status="ok", restored_mtime=restore_artifact_mtime(conversation, output))
            except Exception as error:
                item.update(status="failed", error=str(error))
            report.append(item)
    name = "recipe_docx" if docx else "known_office"
    report_path = root.parent / f"{name}_regeneration_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    failed = sum(item["status"] == "failed" for item in report)
    return report, {"candidates": len(report), "ok": len(report)-failed, "failed": failed,
                    "report": str(report_path)}


def regenerate_known(root: Path, node_modules: Path | None = None) -> tuple[list[dict], dict]:
    return regenerate_local(root, docx=False)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("conversations", type=Path)
    parser.add_argument("node_modules", type=Path, nargs="?")
    args = parser.parse_args()
    _, summary = regenerate_known(args.conversations, args.node_modules)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return int(bool(summary["failed"]))


if __name__ == "__main__":
    raise SystemExit(main())
