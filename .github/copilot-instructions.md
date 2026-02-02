# Copilot instructions — ConvertidorDeOrdenes (Seres Salud)

## Big picture (WinForms + Core)
- `ConvertidorDeOrdenes.Desktop` es la app WinForms. Entry point: `ConvertidorDeOrdenes.Desktop/Program.cs`.
- `ConvertidorDeOrdenes.Core` concentra parsing/normalización/validación/export.
- Flujo principal (UI): `ConvertidorDeOrdenes.Desktop/Forms/MainForm.cs` (botón **Analizar**).

## Data flow (lo que no conviene romper)
1) Parse (CSV o XLSX) → `ParseResult.Rows` (`OutputRow` es el modelo A–Y).
2) Auto-resolución de empresa con DB local (`Empresas.xlsx`).
3) Revisión/edición por el usuario (CompanyResolutionForm).
4) Normalizar + validar. Si hay errores, se setea `OutputRow.DescripcionError` y se bloquea export.
5) Export a XLS 97-2003.

## Conventions / reglas de negocio del dominio
- **Frecuencia** válida: `A`, `S`, `R` (ver `ConvertidorDeOrdenes.Core/Services/Validator.cs`).
- **Regla específica**: si `Frecuencia == "R"`, el `Riesgo` se reemplaza por el texto de `Prestacion` (ver `ConvertidorDeOrdenes.Core/Services/Normalizer.cs`).
- `Riesgo` y `DescripcionRiesgo` se truncan a 90 caracteres (mismo archivo).
- Los parsers intentan ser tolerantes con encabezados: normalizan (sin acentos, solo letras/números/espacios, lowercase) y aceptan aliases (ver `ConvertidorDeOrdenes.Core/Parsers/CsvOrderParser.cs` y `ConvertidorDeOrdenes.Core/Parsers/XlsxOrderParser.cs`).
- CSV: se lee como `iso-8859-1` y registra codepages (`Encoding.RegisterProvider(...)`). Mantener esto si se toca el parsing.

## Output (XLS) — orden fijo de columnas
- El XLS es **Excel 97-2003** vía NPOI HSSF: `ConvertidorDeOrdenes.Core/Services/XlsExporter.cs`.
- El orden A–Y está “cableado” (headers y celdas). No reordenar ni renombrar sin actualizar la UI (grid) y el exportador.
- Algunas columnas se fuerzan vacías/constantes en export hoy (ej: `E`, `O`, `W`).

## Empresas.xlsx (DB local) y rutas
- Persistencia por usuario en `%LOCALAPPDATA%\Seres Salud\ConvertidorDeOrdenes\...` (ver `ConvertidorDeOrdenes.Desktop/Services/AppPaths.cs`).
- Seed inicial: `DB/Empresas.xlsx` junto al exe → copiado a LocalAppData si falta (ver `ConvertidorDeOrdenes.Desktop/Services/AppInitializer.cs`).
- Lectura de DB y compatibilidad de formatos: `ConvertidorDeOrdenes.Core/Services/CompanyRepositoryExcel.cs` (soporta 2 esquemas y hace backup antes de borrar).

## PrestacionesMap
- El mapeo se carga desde la carpeta de instalación (mismo directorio del exe): `PrestacionesMap.csv` o `PrestacionesMap.xlsx` (ver `ConvertidorDeOrdenes.Core/Services/PrestacionMapper.cs`).

## Updates (GitHub Releases)
- Se consulta `releases/latest` y se espera un tag `vX.Y.Z` y un asset preferido `ConvertidorDeOrdenes-Setup.exe` (ver `ConvertidorDeOrdenes.Desktop/Services/Updates/UpdateService.cs`).
- Token opcional para repos privados: env `SERESSALUD_GITHUB_TOKEN` o `GITHUB_TOKEN`.
- Auto-check: máximo 1 vez cada 24h; estado en `update_state.json` dentro de LocalAppData.

## Developer workflows (PowerShell/.NET 8)
- SDK fijado por `global.json` (NET 8). Preferir `dotnet` en vez de “click-build” para reproducibilidad.
- Build:
  - `dotnet restore`
  - `dotnet build ConvertidorDeOrdenes.sln -c Release`
- Run Desktop:
  - `dotnet run --project ConvertidorDeOrdenes.Desktop -c Debug`
- Smoke test parser (rápido):
  - `dotnet run --project ParserTester` (espera `revaluaciones_30_06012026.csv` en la raíz del repo).
- Publish e instalador:
  - `pwsh -File .\scripts\publish.ps1 -Configuration Release -Runtime win-x64 -Version 1.2.3`
  - `pwsh -File .\scripts\build-installer.ps1 -Version 1.2.3` (requiere Inno Setup 6: `ISCC.exe`).
  - `pwsh -File .\scripts\release.ps1 -Version 1.2.3` (bump versión + installer + commit/tag/push).
