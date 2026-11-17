"""Cross-platform helpers for tool discovery and Windows-specific workarounds.

On Windows, tools installed via winget are placed in per-user directories that
are not always on PATH in non-interactive shells.  This module searches known
winget locations and augments ``os.environ['PATH']`` so that downstream
``shutil.which()`` calls and subprocess invocations find pandoc, pdflatex,
pdftoppm, soffice, and similar tools.

Usage (at the top of any regeneration script)::

    from winenv import ensure_tools_on_path
    ensure_tools_on_path()
"""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
from pathlib import Path


# Known winget install locations (per-user, no admin required).
_WINGET_PACKAGES = Path.home() / "AppData" / "Local" / "Microsoft" / "WinGet" / "Packages"

# Static install locations for specific winget packages.
_TOOL_DIRS: list[Path] = [
    Path.home() / "AppData" / "Local" / "Pandoc",                       # JohnMacFarlane.Pandoc
    Path.home() / "AppData" / "Local" / "Programs" / "MiKTeX" / "miktex" / "bin" / "x64",  # MiKTeX.MiKTeX
    Path("C:/Program Files/LibreOffice/program"),                       # TheDocumentFoundation.LibreOffice
]

# Poppler's bin lives inside a versioned subdirectory under the winget package root.
_POPOPPER_PACKAGE_PATTERN = "oschwartz10612.Poppler_*"


def _find_poppler() -> Path | None:
    """Locate the Poppler ``bin`` directory inside the winget package root."""
    if not _WINGET_PACKAGES.is_dir():
        return None
    for pkg in _WINGET_PACKAGES.iterdir():
        if pkg.name.startswith("oschwartz10612.Poppler_"):
            # Structure: <pkg>/poppler-<version>/Library/bin
            for version_dir in pkg.iterdir():
                candidate = version_dir / "Library" / "bin"
                if (candidate / "pdftoppm.exe").exists():
                    return candidate
    return None


def ensure_tools_on_path() -> None:
    """Add known winget tool directories to ``PATH`` if they are not already present.

    On non-Windows platforms this is a no-op — Linux/macOS installations put
    tools on PATH through the system package manager.
    """
    if platform.system() != "Windows":
        return

    extra: list[str] = []
    for tool_dir in _TOOL_DIRS:
        if tool_dir.is_dir():
            p = str(tool_dir)
            if p not in os.environ.get("PATH", ""):
                extra.append(p)

    poppler = _find_poppler()
    if poppler:
        p = str(poppler)
        if p not in os.environ.get("PATH", ""):
            extra.append(p)

    if extra:
        os.environ["PATH"] = os.pathsep.join(extra) + os.pathsep + os.environ.get("PATH", "")


def kill_orphaned_tex_processes() -> None:
    """Kill orphaned pdflatex/MiKTeX processes on Windows.

    On Windows, ``pdflatex`` can leave ``.log`` files locked in temporary
    directories, preventing ``TemporaryDirectory`` cleanup.  Calling this in a
    ``finally`` block after LaTeX runs avoids ``PermissionError: [WinError 32]``.
    """
    if platform.system() != "Windows":
        return
    for proc_name in ("pdflatex.exe", "miktex-pdflatex.exe", "texify.exe"):
        try:
            subprocess.run(
                ["taskkill", "/F", "/IM", proc_name],
                capture_output=True,
                timeout=10,
                check=False,
            )
        except Exception:
            pass


def to_posix_path(path: Path) -> str:
    """Return a forward-slash string representation of *path*.

    When substituting paths into recovered Python or JavaScript source code,
    Windows backslash paths are dangerous: ``\\a``, ``\\f``, ``\\b``, ``\\n``,
    ``\\t``, ``\\v``, ``\\0``, and ``\\2`` (from ``\\2026``) are interpreted as
    escape sequences by the recovered generator's string literals, corrupting
    the path.  Always use ``to_posix_path()`` when injecting paths into source.
    """
    return path.as_posix()
