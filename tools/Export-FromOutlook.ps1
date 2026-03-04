#Requires -Version 5.1
<#
.SYNOPSIS
    Exports all VBA components from Outlook's VbaProject.OTM to .\src\.

.DESCRIPTION
    Use this script to save your current Outlook VBA code back into the
    Git repository before committing.

.NOTES
    Outlook must be running before executing this script.
    Enable 'Trust access to the VBA project object model' in Trust Center.
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

# vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2, vbext_ct_MSForm = 3, vbext_ct_Document = 100
$typeMap = @{
    1   = "modules"
    2   = "classes"
    3   = "forms"
    100 = "classes"   # Document module (e.g. ThisOutlookSession)
}

$extMap = @{
    1   = ".bas"
    2   = ".cls"
    3   = ".frm"
    100 = ".cls"
}

foreach ($component in $vbaProject.VBComponents) {
    $type = $component.Type
    if (-not $typeMap.ContainsKey($type)) {
        Write-Host "Skipping unknown type ($type): $($component.Name)"
        continue
    }

    $subFolder = Join-Path $srcRoot $typeMap[$type]
    if (-not (Test-Path $subFolder)) {
        New-Item -ItemType Directory -Path $subFolder | Out-Null
    }

    $ext      = $extMap[$type]
    $destFile = Join-Path $subFolder "$($component.Name)$ext"

    Write-Host "Exporting: $($component.Name)$ext ..." -NoNewline
    $component.Export($destFile) | Out-Null
    Write-Host " OK"
}

Write-Host ""
Write-Host "Export complete. Files written to: $srcRoot" -ForegroundColor Green
