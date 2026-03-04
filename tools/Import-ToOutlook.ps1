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
    $importExtensions = @("*.bas", "*.cls", "*.frm")
    $files = Get-ChildItem -Path $srcRoot -Recurse -Include $importExtensions |
             Where-Object { $_.BaseName -ne "ThisOutlookSession" } |
             Sort-Object FullName

    Write-Host ""
    Write-Host "Automatic import is not available." -ForegroundColor Yellow
    Write-Host "Application.VBE is blocked (method not supported) or" -ForegroundColor Yellow
    Write-Host "'Trust access to the VBA project object model' is disabled." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please import the files manually in the VBA Editor:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  1. Press Alt+F11 in Outlook to open the VBA Editor" -ForegroundColor White
    Write-Host "  2. In the menu: File -> Import File  (or press Ctrl+M)" -ForegroundColor White
    Write-Host "  3. Import these files IN ORDER:" -ForegroundColor White
    Write-Host ""

    $i = 1
    foreach ($f in $files) {
        Write-Host ("     {0}. {1}" -f $i, $f.FullName) -ForegroundColor Green
        $i++
    }

    Write-Host ""
    Write-Host "  4. For ThisOutlookSession.cls:" -ForegroundColor White
    $sessionFile = Join-Path $srcRoot "classes\ThisOutlookSession.cls"
    Write-Host "     Open the file below and PASTE its code into the" -ForegroundColor White
    Write-Host "     'ThisOutlookSession' module in the VBA Editor:" -ForegroundColor White
    Write-Host "     $sessionFile" -ForegroundColor Green
    Write-Host ""

    # Open Explorer to the src folder for convenience
    Write-Host "Opening src folder in Explorer..." -ForegroundColor Cyan
    Start-Process explorer.exe $srcRoot

    # Also open each importable file in the default editor for copy-paste convenience
    Write-Host ""
    Write-Host "Tip: You can drag .bas/.cls files from Explorer directly into the" -ForegroundColor DarkGray
    Write-Host "     VBA Editor's Project window to import them." -ForegroundColor DarkGray
    Write-Host ""
    exit 0
}

# ---- Import files ------------------------------------------------------------
$importExtensions = @("*.bas", "*.cls", "*.frm")
$files = Get-ChildItem -Path $srcRoot -Recurse -Include $importExtensions

foreach ($file in $files) {
    Write-Host "Importing: $($file.Name) ..." -NoNewline

    # Remove existing component with the same name to avoid duplicates
    $componentName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

    # Skip ThisOutlookSession – it is a built-in document class, not importable
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
Write-Host "NOTE: Apply src\classes\ThisOutlookSession.cls manually by copying its" -ForegroundColor Yellow
Write-Host "      code into the 'ThisOutlookSession' module in the Outlook VBA IDE." -ForegroundColor Yellow
