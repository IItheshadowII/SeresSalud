param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [string]$Branch = "main"
)

# Script de automatización de releases locales.
# Hace, en orden:
# 1) Actualiza la versión en el proyecto Desktop
# 2) Construye el instalador llamando a build-installer.ps1
# 3) Crea commit "Release X.Y.Z" (si hay cambios)
# 4) Crea tag vX.Y.Z
# 5) git push de la rama y del tag (dispara workflow de GitHub Actions)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot   = Resolve-Path (Join-Path $scriptRoot "..")

function Update-DesktopVersion {
    param(
        [string]$Version
    )

    $csprojPath = Join-Path $repoRoot 'ConvertidorDeOrdenes.Desktop\ConvertidorDeOrdenes.Desktop.csproj'
    if (-not (Test-Path $csprojPath)) {
        throw "No se encontró el archivo de proyecto: $csprojPath"
    }

    Write-Host "Actualizando versión en $csprojPath a $Version" -ForegroundColor Cyan

    [xml]$xml = Get-Content $csprojPath

    # Buscar el PropertyGroup que contenga un elemento <Version>
    $pg = $null
    foreach ($group in $xml.Project.PropertyGroup) {
        if ($group.SelectSingleNode('Version')) {
            $pg = $group
            break
        }
    }

    if (-not $pg) {
        throw "No se encontró el nodo <Version> en el csproj"
    }

    $pg.Version         = $Version
    $pg.AssemblyVersion = "$Version.0"
    $pg.FileVersion     = "$Version.0"

    $xml.Save($csprojPath)
}

function Invoke-Git {
    param(
        [string]$Arguments
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'git'
    $psi.Arguments = $Arguments
    $psi.WorkingDirectory = $repoRoot.Path
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false

    $p = [System.Diagnostics.Process]::Start($psi)
    $p.WaitForExit()

    $out = $p.StandardOutput.ReadToEnd()
    $err = $p.StandardError.ReadToEnd()

    if ($out) { Write-Host $out.TrimEnd() }
    if ($p.ExitCode -ne 0) {
        throw "git $Arguments falló con código $($p.ExitCode): $err"
    }
}

Write-Host "=== Release $Version ===" -ForegroundColor Green

# 1) Actualizar versión
Update-DesktopVersion -Version $Version

# 2) Construir instalador
Write-Host "Ejecutando build-installer.ps1..." -ForegroundColor Green
& (Join-Path $scriptRoot 'build-installer.ps1') -Version $Version

# 3) Commit (si hay cambios)
Write-Host "Revisando cambios en git..." -ForegroundColor Green

$psiStatus = New-Object System.Diagnostics.ProcessStartInfo
$psiStatus.FileName = 'git'
$psiStatus.Arguments = 'status --porcelain'
$psiStatus.WorkingDirectory = $repoRoot.Path
$psiStatus.RedirectStandardOutput = $true
$psiStatus.UseShellExecute = $false

$pStatus = [System.Diagnostics.Process]::Start($psiStatus)
$pStatus.WaitForExit()
$changes = $pStatus.StandardOutput.ReadToEnd()

if (-not [string]::IsNullOrWhiteSpace($changes)) {
    Write-Host "Hay cambios, preparando commit..." -ForegroundColor Green

    # Incluir también archivos nuevos (ej: scripts/release.ps1)
    Invoke-Git "add -A"

    try {
        Invoke-Git "commit -m \"Release $Version\""
    } catch {
        # Si no hay nada que commitear (por ejemplo porque solo había outputs ignorados), continuar.
        Write-Warning "No se pudo crear commit (posible 'nothing to commit'): $_"
    }
} else {
    Write-Host "No hay cambios para commitear." -ForegroundColor Yellow
}

# 4) Tag
Write-Host "Creando tag v$Version..." -ForegroundColor Green
try {
    Invoke-Git "tag v$Version"
} catch {
    Write-Warning "No se pudo crear el tag (quizás ya existe): $_"
}

# 5) Push rama + tag
Write-Host "Haciendo push de $Branch y del tag v$Version..." -ForegroundColor Green
Invoke-Git "push origin $Branch"
Invoke-Git "push origin v$Version"

Write-Host "Release $Version completado. Verificá en GitHub que el workflow de release se ejecute correctamente." -ForegroundColor Green
