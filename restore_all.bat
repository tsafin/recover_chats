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

REM ── Step 2: First restore pass ─────────────────────────────
echo.
echo === Step 1: First restore pass ===
python "%SCRIPT_DIR%\restore_export.py" --source "%SOURCE%" --output "%OUTPUT%"
if errorlevel 1 (
    echo ERROR: First restore pass failed.
    exit /b 1
)

REM ── Step 3: Install npm dependencies ───────────────────────
where node >nul 2>&1
if not errorlevel 1 (
    echo.
    echo === Step 2: Installing Node.js dependencies ===
    npm install --prefix "%OUTPUT%\runtime" pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp --no-audit --no-fund
) else (
    echo WARNING: Node.js not found. Skipping Node.js regeneration.
)

REM ── Step 4: Regenerate Office files ────────────────────────
echo.
echo === Step 3: Python Office regeneration ===
python "%SCRIPT_DIR%\regenerate_python_office.py" "%OUTPUT%\conversations"

if not errorlevel 1 (
    echo.
    echo === Step 4: Node.js Office regeneration ===
    if exist "%OUTPUT%\runtime\node_modules" (
        python "%SCRIPT_DIR%\regenerate_node_office.py" "%OUTPUT%\conversations" "%OUTPUT%\runtime\node_modules"
    ) else (
        echo Skipping: no node_modules directory.
    )
)

echo.
echo === Step 5: DOCX recipe regeneration ===
python "%SCRIPT_DIR%\regenerate_recipe_docx.py" "%OUTPUT%\conversations"

echo.
echo === Step 6: Known Office recipe regeneration ===
if exist "%OUTPUT%\runtime\node_modules" (
    python "%SCRIPT_DIR%\regenerate_known_office.py" "%OUTPUT%\conversations" "%OUTPUT%\runtime\node_modules"
) else (
    echo Skipping: no node_modules directory.
)

REM ── Step 5: Second restore pass ────────────────────────────
echo.
echo === Step 7: Second restore pass (link regenerated files) ===
python "%SCRIPT_DIR%\restore_export.py" --source "%SOURCE%" --output "%OUTPUT%"

REM ── Step 6: Validation ─────────────────────────────────────
echo.
echo === Step 8: Validation ===
python "%SCRIPT_DIR%\check_restoration.py" "%OUTPUT%"

echo.
echo ============================================================
echo  Done. Open %OUTPUT%\INDEX.html in a browser.
echo ============================================================
endlocal
