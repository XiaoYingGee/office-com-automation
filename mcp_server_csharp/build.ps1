<#
.SYNOPSIS
    Build ExcelMcp.exe using .NET Framework csc.exe (no dotnet SDK required).
    Downloads Newtonsoft.Json from NuGet if not present.
#>
$ErrorActionPreference = 'Stop'

$ScriptDir = $PSScriptRoot
$SrcDir = Join-Path $ScriptDir 'ExcelMcp'
$OutDir = Join-Path $SrcDir 'bin\Release\net48'
$Csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'

if (-not (Test-Path $Csc)) {
    $Csc = 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
Write-Host "csc: $Csc"

# --- Download Newtonsoft.Json if needed ---
$PkgDir = Join-Path $ScriptDir 'packages'
$JsonDll = Join-Path $PkgDir 'Newtonsoft.Json.dll'
if (-not (Test-Path $JsonDll)) {
    Write-Host "Downloading Newtonsoft.Json..."
    New-Item -ItemType Directory -Force -Path $PkgDir | Out-Null
    $nupkg = Join-Path $PkgDir 'newtonsoft.json.13.0.3.nupkg'
    Invoke-WebRequest -Uri 'https://www.nuget.org/api/v2/package/Newtonsoft.Json/13.0.3' -OutFile $nupkg
    # Extract dll from nupkg (it's a zip)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($nupkg)
    $entry = $zip.Entries | Where-Object { $_.FullName -like 'lib/net45/Newtonsoft.Json.dll' } | Select-Object -First 1
    if (-not $entry) { $entry = $zip.Entries | Where-Object { $_.FullName -like '*net4*/Newtonsoft.Json.dll' } | Select-Object -First 1 }
    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $JsonDll, $true)
    $zip.Dispose()
    Remove-Item $nupkg -Force
    Write-Host "  -> $JsonDll"
}

# --- Find Excel Interop assembly ---
$ExcelInterop = Get-ChildItem 'C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel' -Recurse -Filter '*.dll' | Select-Object -First 1
$OfficeCore = Get-ChildItem 'C:\Windows\assembly\GAC_MSIL\office' -Recurse -Filter '*.dll' | Select-Object -First 1

if (-not $ExcelInterop) {
    Write-Error "Microsoft.Office.Interop.Excel.dll not found in GAC. Is Office installed?"
    exit 1
}
Write-Host "Excel Interop: $($ExcelInterop.FullName)"
Write-Host "Office Core: $($OfficeCore.FullName)"

# --- Compile ---
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$Sources = Get-ChildItem $SrcDir -Filter '*.cs' | ForEach-Object { $_.FullName }
$References = @(
    "System.dll",
    "System.Core.dll",
    "Microsoft.CSharp.dll",
    $JsonDll,
    $ExcelInterop.FullName,
    $OfficeCore.FullName
)

$RefArgs = ($References | ForEach-Object { "/reference:`"$_`"" }) -join ' '
$SrcArgs = ($Sources | ForEach-Object { "`"$_`"" }) -join ' '
$OutExe = Join-Path $OutDir 'ExcelMcp.exe'

$cmd = "& `"$Csc`" /nologo /optimize /target:exe /out:`"$OutExe`" /platform:anycpu $RefArgs $SrcArgs"
Write-Host "`nCompiling..."
Write-Host $cmd
Invoke-Expression $cmd

if ($LASTEXITCODE -ne 0) {
    Write-Error "Compilation failed!"
    exit 1
}

# Copy Newtonsoft.Json.dll next to exe
Copy-Item $JsonDll $OutDir -Force

Write-Host "`nBuild successful: $OutExe" -ForegroundColor Green
