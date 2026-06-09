<#
.SYNOPSIS
    Remove the ExcelEditor COM add-in registration (per-user).
#>
$ErrorActionPreference = 'Stop'

$ProgId = 'ExcelEditor.AddIn'
$Clsid  = '{A1B2C3D4-E5F6-4789-ABCD-EF0123456789}'

# Remove CLSID
$clsidPath = "HKCU:\Software\Classes\CLSID\$Clsid"
if (Test-Path $clsidPath) {
    Remove-Item -Path $clsidPath -Recurse -Force
    Write-Host "Removed CLSID $Clsid"
}

# Remove ProgId
$progPath = "HKCU:\Software\Classes\$ProgId"
if (Test-Path $progPath) {
    Remove-Item -Path $progPath -Recurse -Force
    Write-Host "Removed ProgId $ProgId"
}

# Remove Excel AddIns entry
$addinPath = "HKCU:\Software\Microsoft\Office\Excel\AddIns\$ProgId"
if (Test-Path $addinPath) {
    Remove-Item -Path $addinPath -Recurse -Force
    Write-Host "Removed Excel add-in entry"
}

Write-Host "Unregistration complete." -ForegroundColor Green
