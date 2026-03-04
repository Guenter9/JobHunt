#Requires -Version 5.1
<#
.SYNOPSIS
    Imports all VBA source files from .\src\ into Outlook's VbaProject.OTM.

.DESCRIPTION
    Uses COM automation to open the Outlook VBA project and import
    .bas (standard modules) and .cls (class modules) files.
    Run this script once after cloning the repository, or whenever
    you want to refresh Outlook with the latest source code.

.NOTES
    Outlook must be running before executing this script.
    Macro security must allow programmatic access to the VBA project:
      Outlook > File > Options > Trust Center > Trust Center Settings >
      Macro Settings > "Trust access to the VBA project object model"
#>

# Note: StrictMode is intentionally NOT set here – COM late-binding objects
# do not expose a fixed property list, which causes false "property not found"
# errors under StrictMode -Version Latest.
$ErrorActionPreference = "Stop"

$srcRoot = Join-Path $PSScriptRoot "..\src"

# ---- Validate prerequisites --------------------------------------------------
# Note: Marshal.GetActiveObject was removed in .NET 5+; instead verify the
# process is running first, then use New-Object -ComObject (Outlook is a
# single-instance COM server and returns the existing instance when running).
if (-not (Get-Process -Name outlook -ErrorAction SilentlyContinue)) {
    Write-Error "Outlook is not running. Please start Outlook and try again."
    exit 1
}

$outlook = $null
try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    Write-Error "Could not connect to Outlook: $_"
    exit 1
}

$vbaProject = $null
$vbeAvailable = $false
try {
    $vbe = $outlook.VBE
    if ($null -ne $vbe) {
        $vbaProject = $vbe.ActiveVBProject
        if ($null -ne $vbaProject) { $vbeAvailable = $true }
    }
} catch { }

if (-not $vbeAvailable) {
    # ---- Fallback: guided manual import --------------------------------------
    # Note: .frm files require a companion .frx binary – importing them directly
    # causes error &H80004005.  frmJobApplication must be created manually.
    $importExtensions = @("*.bas", "*.cls")
    $files = Get-ChildItem -Path $srcRoot -Recurse -Include $importExtensions |
             Where-Object { $_.BaseName -notin @("ThisOutlookSession") } |
             Sort-Object FullName

    $frmCodeFile = Join-Path $srcRoot "classes\frmJobApplication.frm"
    $sessionFile = Join-Path $srcRoot "classes\ThisOutlookSession.cls"

    Write-Host ""
    Write-Host "Automatic import is not available." -ForegroundColor Yellow
    Write-Host "Application.VBE is blocked or 'Trust access to the VBA project" -ForegroundColor Yellow
    Write-Host "object model' is disabled." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Follow these steps in the VBA Editor (Alt+F11):" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  STEP 1 – Create the UserForm (do this first):" -ForegroundColor White
    Write-Host "     a) Einfuegen / Insert > UserForm" -ForegroundColor Green
    Write-Host "     b) Properties window (F4): set Name = frmJobApplication" -ForegroundColor Green
    Write-Host "     c) Double-click the form to open its code window" -ForegroundColor Green
    Write-Host "     d) Replace ALL existing code with the contents of:" -ForegroundColor Green
    Write-Host "        $frmCodeFile" -ForegroundColor Cyan
    Write-Host "        (copy everything from 'Option Explicit' onwards)" -ForegroundColor Green
    Write-Host ""
    Write-Host "  STEP 2 – Import remaining files (File > Import / Ctrl+M):" -ForegroundColor White
    Write-Host ""

    $i = 1
    foreach ($f in $files) {
        Write-Host ("     {0}. {1}" -f $i, $f.FullName) -ForegroundColor Green
        $i++
    }

    Write-Host ""
    Write-Host "  STEP 3 – ThisOutlookSession (paste, do NOT import):" -ForegroundColor White
    Write-Host "     Open 'ThisOutlookSession' in the VBA Project tree and" -ForegroundColor Green
    Write-Host "     PASTE the code from:" -ForegroundColor Green
    Write-Host "     $sessionFile" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Opening src folder in Explorer..." -ForegroundColor Cyan
    Start-Process explorer.exe $srcRoot
    Write-Host ""
    exit 0
}

# ---- Import files ------------------------------------------------------------
# Note: .frm files require a companion .frx binary and cannot be imported
# programmatically without it.  frmJobApplication is handled separately below.
$importExtensions = @("*.bas", "*.cls")
$files = Get-ChildItem -Path $srcRoot -Recurse -Include $importExtensions

foreach ($file in $files) {
    Write-Host "Importing: $($file.Name) ..." -NoNewline

    $componentName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

    # Skip ThisOutlookSession – built-in document class, not importable
    if ($componentName -eq "ThisOutlookSession") {
        Write-Host " SKIPPED (apply manually)"
        continue
    }

    try {
        $existing = $vbaProject.VBComponents($componentName)
        $vbaProject.VBComponents.Remove($existing)
    } catch { }  # Ignore if component doesn't exist

    $vbaProject.VBComponents.Import($file.FullName) | Out-Null
    Write-Host " OK"
}

Write-Host ""
Write-Host "Import complete." -ForegroundColor Green
Write-Host ""
Write-Host "MANUAL STEPS STILL REQUIRED:" -ForegroundColor Yellow
Write-Host "  1. frmJobApplication: in VBA Editor, Insert > UserForm," -ForegroundColor Yellow
Write-Host "     rename to 'frmJobApplication', paste code from:" -ForegroundColor Yellow
$frmCodeFile = Join-Path (Split-Path $srcRoot) "src\classes\frmJobApplication.frm"
Write-Host "     $frmCodeFile" -ForegroundColor Cyan
Write-Host "  2. ThisOutlookSession: paste code from src\classes\ThisOutlookSession.cls" -ForegroundColor Yellow
Write-Host "     into the existing 'ThisOutlookSession' module." -ForegroundColor Yellow
