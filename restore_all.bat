@echo off
REM ============================================================
REM  One-click restore script for Claude export on Windows.
REM
REM  Usage:
REM    restore_all.bat <source_dir> <output_dir>
REM
REM  Example:
REM    recover_chats\restore_all.bat . restored_export_new
REM ============================================================

setlocal enabledelayedexpansion

set SOURCE=%~1
if "%SOURCE%"=="" set SOURCE=.
set OUTPUT=%~2
if "%OUTPUT%"=="" set OUTPUT=restored_export_new

set SCRIPT_DIR=%~dp0
if "%SCRIPT_DIR:~-1%"=="\" set SCRIPT_DIR=%SCRIPT_DIR:~0,-1%

echo ============================================================
echo  Claude Export Restore - Windows
echo  Source: %SOURCE%
echo  Output: %OUTPUT%
echo  Scripts: %SCRIPT_DIR%
echo ============================================================

REM ── Step 0: Check Python ────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.10+ and add to PATH.
    exit /b 1
)

REM ── Step 1: Install Python dependencies if needed ──────────
python -c "import markdown_it" >nul 2>&1
if errorlevel 1 (
    echo Installing Python dependencies...
    python -m pip install markdown-it-py openpyxl python-docx python-pptx matplotlib pandas numpy pillow
)

REM ── Step 2: Install npm dependencies ───────────────────────
where node >nul 2>&1
if not errorlevel 1 (
    echo.
    echo === Installing Node.js dependencies ===
    npm install --prefix "%OUTPUT%\runtime" pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp --no-audit --no-fund
) else (
    echo WARNING: Node.js not found. Skipping Node.js regeneration.
)

REM ── Step 3: Unified restoration pipeline ───────────────────
echo.
echo === Running unified restore/regenerate/render/validate pipeline ===
python "%SCRIPT_DIR%\restore_pipeline.py" --source "%SOURCE%" --output "%OUTPUT%"
if errorlevel 1 (
    echo ERROR: Restoration pipeline failed. See %OUTPUT%\pipeline_report.json.
    exit /b 1
)

echo.
echo ============================================================
echo  Done. Open %OUTPUT%\INDEX.html in a browser.
echo ============================================================
endlocal
