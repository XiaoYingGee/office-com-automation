<#
.SYNOPSIS
    Per-user (HKCU) registration for the ExcelEditor in-process COM add-in.
    No administrator rights required.

.PARAMETER DllPath
    Optional path to ExcelEditorAddin.dll. Defaults to the Release build output.
#>
[CmdletBinding()]
param(
    [string]$DllPath
)

$ErrorActionPreference = 'Stop'

$ProgId   = 'ExcelEditor.AddIn'
$Clsid    = '{A1B2C3D4-E5F6-4789-ABCD-EF0123456789}'
$ClassFqn = 'ExcelEditorAddin.Connect'

if (-not $DllPath) {
    $DllPath = Join-Path $PSScriptRoot 'ExcelEditorAddin\bin\Release\net48\ExcelEditorAddin.dll'
}
$DllPath = (Resolve-Path -LiteralPath $DllPath).Path
Write-Host "DLL: $DllPath"

$asmName     = [System.Reflection.AssemblyName]::GetAssemblyName($DllPath)
$asmFullName = $asmName.FullName
$asmVersion  = $asmName.Version.ToString()
$runtimeVer  = 'v4.0.30319'
$codeBase    = ([System.Uri]$DllPath).AbsoluteUri
Write-Host "Assembly: $asmFullName"

function Set-Key([string]$path) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
}

# 1) CLSID registration (per-user COM, mscoree shim)
$clsidRoot = "HKCU:\Software\Classes\CLSID\$Clsid"
Set-Key $clsidRoot
New-ItemProperty -Path $clsidRoot -Name '(default)' -Value $ClassFqn -PropertyType String -Force | Out-Null

$inproc = "$clsidRoot\InprocServer32"
Set-Key $inproc
New-ItemProperty -Path $inproc -Name '(default)'      -Value 'mscoree.dll'  -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'ThreadingModel' -Value 'Both'         -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'Class'          -Value $ClassFqn      -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'Assembly'       -Value $asmFullName   -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'RuntimeVersion' -Value $runtimeVer    -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'CodeBase'       -Value $codeBase      -PropertyType String -Force | Out-Null

$inprocVer = "$inproc\$asmVersion"
Set-Key $inprocVer
New-ItemProperty -Path $inprocVer -Name 'Class'          -Value $ClassFqn    -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inprocVer -Name 'Assembly'       -Value $asmFullName -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inprocVer -Name 'RuntimeVersion' -Value $runtimeVer  -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inprocVer -Name 'CodeBase'       -Value $codeBase    -PropertyType String -Force | Out-Null

$clsidProgId = "$clsidRoot\ProgId"
Set-Key $clsidProgId
New-ItemProperty -Path $clsidProgId -Name '(default)' -Value $ProgId -PropertyType String -Force | Out-Null

# 2) ProgId -> CLSID
$progRoot = "HKCU:\Software\Classes\$ProgId"
Set-Key $progRoot
New-ItemProperty -Path $progRoot -Name '(default)' -Value 'ExcelEditor In-Process Add-in' -PropertyType String -Force | Out-Null
$progClsid = "$progRoot\CLSID"
Set-Key $progClsid
New-ItemProperty -Path $progClsid -Name '(default)' -Value $Clsid -PropertyType String -Force | Out-Null

# 3) Excel per-user AddIns entry (LoadBehavior=3 => load at startup)
$addinKey = "HKCU:\Software\Microsoft\Office\Excel\AddIns\$ProgId"
Set-Key $addinKey
New-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value 'ExcelEditor In-Process Add-in' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name 'Description'  -Value 'In-process bridge for Excel automation' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null

Write-Host "Registered '$ProgId' (CLSID $Clsid) per-user. LoadBehavior=3." -ForegroundColor Green
