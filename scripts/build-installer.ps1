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
$programFilesX86 = [Environment]::GetFolderPath('ProgramFilesX86')
$programFiles    = [Environment]::GetFolderPath('ProgramFiles')

$possible = @(
  (Join-Path $programFilesX86 'Inno Setup 6\ISCC.exe'),
  (Join-Path $programFiles    'Inno Setup 6\ISCC.exe')
)

$iscc = $possible | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $iscc) {
  throw "No se encontró ISCC.exe (Inno Setup). Instalá Inno Setup 6 y reintentá."
}

Write-Host "OK. ISCC encontrado en: $iscc" -ForegroundColor Green

$iss = Join-Path $repoRoot 'Installer\ConvertidorDeOrdenes.iss'

$outputDir = Join-Path $repoRoot 'Installer\Output'
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$outputBaseFilename = 'ConvertidorDeOrdenes-Setup'

Write-Host "Compilando instalador con: $iscc" -ForegroundColor Cyan

$isccArgs = @()
if ($Version) {
  $isccArgs += "/DMyAppVersion=$Version"
}
$isccArgs += "/O$outputDir"
$isccArgs += "/F$outputBaseFilename"
$isccArgs += $iss

& $iscc @isccArgs

$expectedInstaller = Join-Path $outputDir ($outputBaseFilename + '.exe')
if (-not (Test-Path $expectedInstaller)) {
  Write-Host "No se encontró el instalador esperado en: $expectedInstaller" -ForegroundColor Red
  Write-Host "Contenido de ${outputDir}:" -ForegroundColor Yellow
  Get-ChildItem -Path $outputDir -Force | Select-Object FullName,Length | Out-String | Write-Host
  throw "Fallo al generar el instalador (archivo faltante): $expectedInstaller"
}

Write-Host "OK. El instalador queda en la carpeta Installer (OutputBaseFilename)." -ForegroundColor Green
