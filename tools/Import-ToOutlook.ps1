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

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$srcRoot = Join-Path $PSScriptRoot "..\src"

# ---- Validate prerequisites --------------------------------------------------
$outlook = $null
try {
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} catch {
    Write-Error "Outlook is not running. Please start Outlook and try again."
    exit 1
}

$vbaProject = $outlook.VBE.ActiveVBProject
if (-not $vbaProject) {
    Write-Error "Could not access the Outlook VBA project. " +
                "Enable 'Trust access to the VBA project object model' in Trust Center settings."
    exit 1
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
