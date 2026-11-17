<#
.SYNOPSIS
  Installs all dependencies needed to restore Claude export chats on Windows.
.DESCRIPTION
  Uses winget (silent mode) for system tools and pip/npm for language packages.
  Run this once before executing restore_export.py and the regeneration scripts.
.EXAMPLE
  .\install_deps_windows.ps1
#>

$ErrorActionPreference = "Stop"

function Install-WingetPackage {
    param([string]$Id)
    Write-Host "Installing $Id via winget..." -ForegroundColor Cyan
    winget install --id $Id --accept-package-agreements --accept-source-agreements --silent
    if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne -1978335189) {
        # -1978335189 = already installed
        Write-Warning "winget install $Id returned exit code $LASTEXITCODE"
    }
}

function Install-PipPackages {
    param([string[]]$Packages)
    Write-Host "Installing Python packages: $($Packages -join ', ')..." -ForegroundColor Cyan
    python -m pip install --upgrade pip
    python -m pip install $Packages
}

function Install-NpmPackages {
    param([string]$Prefix, [string[]]$Packages)
    Write-Host "Installing npm packages into $Prefix..." -ForegroundColor Cyan
    npm install --prefix $Prefix $Packages --no-audit --no-fund
}

# ── System tools via winget (silent, no interaction) ──────────────────
Write-Host "`n=== Installing system tools via winget ===" -ForegroundColor Yellow

Install-WingetPackage "JohnMacFarlane.Pandoc"
Install-WingetPackage "oschwartz10612.Poppler"
Install-WingetPackage "MiKTeX.MiKTeX"
Install-WingetPackage "TheDocumentFoundation.LibreOffice"

# ── Python packages ───────────────────────────────────────────────────
Write-Host "`n=== Installing Python packages ===" -ForegroundColor Yellow

Install-PipPackages @(
    "markdown-it-py",
    "openpyxl",
    "python-docx",
    "python-pptx",
    "matplotlib",
    "pandas",
    "numpy",
    "pillow",
    "cairosvg"
)

# Note: cairosvg requires the native Cairo library (libcairo-2.dll).
# On Windows it may not work out of the box. The regeneration scripts
# handle this gracefully — DOCX recipes with SVG fall back to
# alternative rendering paths or report the limitation in the report.

# ── Node.js packages ──────────────────────────────────────────────────
Write-Host "`n=== Installing Node.js packages ===" -ForegroundColor Yellow

# These will be installed into the output directory after the first
# restore_export.py pass. For now, just check node is available.
$nodeVersion = node --version 2>$null
if (-not $nodeVersion) {
    Write-Warning "Node.js not found. Install it from https://nodejs.org or: winget install OpenJS.NodeJS.LTS"
} else {
    Write-Host "Node.js $nodeVersion found." -ForegroundColor Green
    Write-Host @"
After the first restore_export.py pass, install npm packages with:

  npm install --prefix <output>/runtime pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp --no-audit --no-fund

The restore script auto-generates HTML in all languages found in recover_chats/locales/.
To add a new language, create a locales/<lang>.json file and re-run restore_export.py.

"@ -ForegroundColor Gray
}

# ── Verify ────────────────────────────────────────────────────────────
Write-Host "`n=== Verifying installations ===" -ForegroundColor Yellow

$tools = @{
    "pandoc"    = $null
    "pdflatex"  = $null
    "pdftoppm"  = $null
}

# Tools installed by winget may not be on PATH yet in this session.
# The Python scripts auto-discover them via winenv.py at runtime.
# Here we just verify the known locations exist.

$pandocPath = "$env:LOCALAPPDATA\Pandoc\pandoc.exe"
$miktexPath = "$env:LOCALAPPDATA\Programs\MiKTeX\miktex\bin\x64\pdflatex.exe"
$librePath = "C:\Program Files\LibreOffice\program\soffice.exe"

$popplerDir = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\oschwartz10612.Poppler_*" -ErrorAction SilentlyContinue
$popplerPath = if ($populerDir) {
    Get-ChildItem "$($populerDir.FullName)\poppler-*\Library\bin\pdftoppm.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
} else { $null }

foreach ($item in @{
    "Pandoc"      = $pandocPath
    "MiKTeX/pdflatex" = $miktexPath
    "LibreOffice" = $librePath
    "Poppler/pdftoppm" = $popplerPath
}.GetEnumerator()) {
    if ($item.Value -and (Test-Path $item.Value)) {
        Write-Host "  OK  $($item.Key): $($item.Value)" -ForegroundColor Green
    } else {
        Write-Host "  MISSING  $($item.Key)" -ForegroundColor Red
    }
}

# Python packages
foreach ($pkg in @("markdown_it", "openpyxl", "docx", "pptx", "matplotlib", "pandas", "numpy", "PIL")) {
    $result = python -c "import $pkg; print('ok')" 2>$null
    if ($result -eq "ok") {
        Write-Host "  OK  Python: $pkg" -ForegroundColor Green
    } else {
        Write-Host "  MISSING  Python: $pkg" -ForegroundColor Red
    }
}

Write-Host @"
`n=== Next steps ===
1. Run the first restore pass:
   python recover_chats\restore_export.py --source . --output .\restored_export_new

2. Install npm packages (if Node.js is available):
   npm install --prefix restored_export_new\runtime pptxgenjs@3.12.0 docx@9.5.1 react react-dom react-icons sharp --no-audit --no-fund

3. Run regeneration scripts:
   python recover_chats\regenerate_python_office.py restored_export_new\conversations
   python recover_chats\regenerate_node_office.py restored_export_new\conversations restored_export_new\runtime\node_modules
   python recover_chats\regenerate_recipe_docx.py restored_export_new\conversations
   python recover_chats\regenerate_known_office.py restored_export_new\conversations restored_export_new\runtime\node_modules

4. Run second restore pass to link regenerated files:
   python recover_chats\restore_export.py --source . --output .\restored_export_new

5. Validate:
   python recover_chats\check_restoration.py restored_export_new
"@ -ForegroundColor Cyan
