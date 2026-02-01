param(
  [ValidateSet('Release','Debug')]
  [string]$Configuration = 'Release',

  [string]$Runtime = 'win-x64',

  [switch]$SelfContained,

  # Opcional: setear versión del release (ej: 1.2.3)
  [string]$Version
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

# 1) Publish
& (Join-Path $PSScriptRoot 'publish.ps1') -Configuration $Configuration -Runtime $Runtime -SelfContained:$SelfContained -Version $Version

# 2) Compilar instalador (Inno Setup)
$innoCandidates = @(
  "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
  "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
) | Where-Object { Test-Path $_ }

if ($innoCandidates.Count -eq 0) {
  throw "No se encontró ISCC.exe (Inno Setup). Instalá Inno Setup 6 y reintentá."
}

$iscc = $innoCandidates[0]
$iss = Join-Path $repoRoot 'Installer\ConvertidorDeOrdenes.iss'

Write-Host "Compilando instalador con: $iscc" -ForegroundColor Cyan

$isccArgs = @()
if ($Version) {
  $isccArgs += "/DMyAppVersion=$Version"
}
$isccArgs += $iss

& $iscc @isccArgs

Write-Host "OK. El instalador queda en la carpeta Installer (OutputBaseFilename)." -ForegroundColor Green
