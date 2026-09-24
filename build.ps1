# 用 Windows 內建的 C# 編譯器把 OrcadPartEditor.cs 編成單一 EXE（免安裝任何東西）
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$csc  = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) { $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe" }

$src = Join-Path $here "OrcadPartEditor.cs"
$out = Join-Path $here "OrcadPartEditor.exe"

$refs = @(
  "/r:System.dll",
  "/r:System.Core.dll",
  "/r:System.Data.dll",
  "/r:System.Xml.dll",
  "/r:System.Drawing.dll",
  "/r:System.Windows.Forms.dll",
  "/r:System.IO.Compression.dll",
  "/r:System.IO.Compression.FileSystem.dll"
)

& $csc /target:winexe /platform:anycpu /optimize+ /nologo "/out:$out" @refs $src
if ($LASTEXITCODE -eq 0) { Write-Host "OK -> $out" } else { Write-Host "Build failed ($LASTEXITCODE)" }
