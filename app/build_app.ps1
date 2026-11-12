# Bien dich InAnG8.exe bang csc.exe cua .NET Framework 4 (co san trong Windows) - khong can cai SDK.
#   powershell -ExecutionPolicy Bypass -File build_app.ps1 [-Out <thu muc>]
param([string]$Out = (Join-Path $PSScriptRoot 'bin'))
$ErrorActionPreference = 'Stop'
$F = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$W = Join-Path $env:WINDIR 'System32\WinMetadata'
New-Item -ItemType Directory -Force $Out | Out-Null
$refs = @(
  "$F\System.dll", "$F\System.Core.dll", "$F\System.Drawing.dll", "$F\System.Windows.Forms.dll",
  "$F\System.Runtime.dll",
  "$W\Windows.Data.winmd", "$W\Windows.Storage.winmd", "$W\Windows.Foundation.winmd", "$W\Windows.UI.winmd"
) | ForEach-Object { "-r:`"$_`"" }
$args = @('-nologo', '-target:winexe', '-platform:anycpu', '-optimize+', '-codepage:65001',
          "-out:`"$(Join-Path $Out 'InAnG8.exe')`"") + $refs
if (Test-Path (Join-Path $PSScriptRoot 'app.ico')) { $args += "-win32icon:`"$(Join-Path $PSScriptRoot 'app.ico')`"" }
$args += (Get-ChildItem $PSScriptRoot -Filter *.cs | ForEach-Object { "`"$($_.FullName)`"" })
& "$F\csc.exe" @args
if ($LASTEXITCODE -ne 0) { throw "csc loi ($LASTEXITCODE)" }
Write-Host "OK -> $(Join-Path $Out 'InAnG8.exe')"
