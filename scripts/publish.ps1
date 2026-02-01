param(
  [ValidateSet('Release','Debug')]
  [string]$Configuration = 'Release',

  # 'win-x64' recomendado para PCs actuales. Cambiá si necesitás win-x86.
  [string]$Runtime = 'win-x64',

  # Publicación portable (requiere .NET Desktop Runtime instalado)
  [switch]$SelfContained,

  # Opcional: setear versión del assembly (ej: 1.2.3)
  [string]$Version
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$desktopProj = Join-Path $repoRoot 'ConvertidorDeOrdenes.Desktop\ConvertidorDeOrdenes.Desktop.csproj'
$outDir = Join-Path $repoRoot 'artifacts\publish'

if (Test-Path $outDir) {
  Remove-Item -Recurse -Force $outDir
}
New-Item -ItemType Directory -Path $outDir | Out-Null

$sc = if ($SelfContained) { 'true' } else { 'false' }

$extraArgs = @()
if ($Version) {
  $extraArgs += "-p:Version=$Version"
}

Write-Host "Publishing to: $outDir" -ForegroundColor Cyan

dotnet publish $desktopProj `
  -c $Configuration `
  -r $Runtime `
  --self-contained $sc `
  -o $outDir `
  @extraArgs

Write-Host "OK. Publish generado en: $outDir" -ForegroundColor Green
Write-Host "Verificá que existan: DB\\Empresas.xlsx y PrestacionesMap.csv dentro del publish." -ForegroundColor Yellow
