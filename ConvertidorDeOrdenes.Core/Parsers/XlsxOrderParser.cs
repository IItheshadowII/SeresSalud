using ConvertidorDeOrdenes.Core.Models;
using ClosedXML.Excel;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

namespace ConvertidorDeOrdenes.Core.Parsers;

/// <summary>
/// Parser para archivos XLSX de Anuales/Semestrales (múltiples solapas)
/// </summary>
public class XlsxOrderParser
{
    public ParseResult Parse(string filePath, string frecuencia, string referente)
    {
        var result = new ParseResult();

        try
        {
            using var workbook = new XLWorkbook(filePath);

            // Ruta de debug para CP
            var debugCpPath = Path.Combine(Path.GetDirectoryName(filePath) ?? "", "cp_debug.txt");
            try
            {
                File.WriteAllText(debugCpPath, $"Archivo: {filePath}{Environment.NewLine}");
            }
            catch { }

            // Buscar hoja "Resumen" o primera hoja
            IXLWorksheet? resumenSheet = null;
            
            foreach (var sheet in workbook.Worksheets)
            {
                if (sheet.Name.Contains("Resumen", StringComparison.OrdinalIgnoreCase) ||
                    sheet.Name.Contains("resumen", StringComparison.OrdinalIgnoreCase))
                {
                    resumenSheet = sheet;
                    break;
                }
            }

            // Si no hay hoja resumen, procesar todas las hojas excepto la primera
            var worksheetsToProcess = new List<IXLWorksheet>();
            
            if (resumenSheet != null)
            {
                // Extraer referencias a solapas desde el resumen (si existen)
                // Por ahora, procesar todas las demás hojas
                try
                {
                    var e3 = resumenSheet.Cell(3, 5).GetString();
                    File.AppendAllText(debugCpPath, $"Hoja Resumen encontrada. E3='{e3}'{Environment.NewLine}");
                }
                catch { }

                worksheetsToProcess = workbook.Worksheets
                    .Where(ws => !ws.Name.Equals(resumenSheet.Name, StringComparison.OrdinalIgnoreCase))
                    .Take(50) // Máximo 50 solapas
                    .ToList();
            }
            else
            {
                // Procesar todas las hojas
                worksheetsToProcess = workbook.Worksheets.Take(50).ToList();
            }

            foreach (var worksheet in worksheetsToProcess)
            {
                try
                {
                    ProcessWorksheet(worksheet, frecuencia, referente, result, debugCpPath);
                }
                catch (Exception ex)
                {
                    result.Warnings.Add($"Error procesando solapa '{worksheet.Name}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            result.Errors.Add($"Error leyendo archivo XLSX: {ex.Message}");
        }

        return result;
    }

    private void ProcessWorksheet(IXLWorksheet worksheet, string frecuencia, string referente, ParseResult result, string? debugCpPath)
    {
        // Buscar fila de encabezados
        var headerRow = FindHeaderRow(worksheet);
        if (headerRow == null)
        {
            result.Warnings.Add($"No se encontraron encabezados en solapa '{worksheet.Name}'");
            return;
        }

        var columnMap = MapColumnsNormalized(worksheet, headerRow.RowNumber());

        // Extraer datos de empresa de la solapa (generalmente están arriba)
        var empresaData = ExtractCompanyDataFromSheet(worksheet, headerRow.RowNumber(), debugCpPath);

        // Procesar filas de datos (después del encabezado)
        var dataStartRow = headerRow.RowNumber() + 1;
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? dataStartRow;

        for (int rowNum = dataStartRow; rowNum <= lastRow; rowNum++)
        {
            var dataRow = worksheet.Row(rowNum);
            
            // Verificar si la fila está vacía
            if (IsEmptyRow(dataRow))
                continue;

            try
            {
                var outputRow = new OutputRow
                {
                    Frecuencia = frecuencia,
                    Referente = referente,
                    
                    // Datos de empresa (desde extracto de la solapa)
                    Contrato = empresaData.Contrato,
                    CuitEmpleador = empresaData.CUIT,
                    Empleador = empresaData.RazonSocial,
                    NroEstablecimiento = empresaData.NroEstablecimiento,
                    CodPostal = empresaData.CodPostal,
                    Localidad = empresaData.Localidad,
                    Provincia = empresaData.Provincia,
                    Telefono = empresaData.Telefono
                };

                // Completar/overridar datos de empresa desde columnas si existen
                var contratoFila = GetCellValue(dataRow, columnMap,
                    "Contrato", "Nro Contrato", "Nro. Contrato");
                if (!string.IsNullOrWhiteSpace(contratoFila))
                    outputRow.Contrato = contratoFila;

                var empleadorFila = GetCellValue(dataRow, columnMap,
                    "Empresa", "Empleador", "Razon Social", "Razón Social");
                if (!string.IsNullOrWhiteSpace(empleadorFila))
                    outputRow.Empleador = empleadorFila;

                var localidadFila = GetCellValue(dataRow, columnMap, "Localidad");
                if (!string.IsNullOrWhiteSpace(localidadFila))
                    outputRow.Localidad = localidadFila;

                var provinciaFila = GetCellValue(dataRow, columnMap, "Provincia");
                if (!string.IsNullOrWhiteSpace(provinciaFila))
                    outputRow.Provincia = provinciaFila;

                // Datos de trabajador (desde columnas) con nombres de encabezado flexibles
                outputRow.Cuil = GetCellValue(dataRow, columnMap,
                    "CUIL", "Cuil", "CUIT", "Cuil beneficiario", "CUIL Beneficiario");

                var nombre1 = GetCellValue(dataRow, columnMap,
                    "Nombre Beneficiario", "Beneficiario", "Apellido y Nombre", "Nombre y Apellido");
                var nombre2 = GetCellValue(dataRow, columnMap,
                    "Apellido y Nombre", "Apellido y nombre", "NyA");
                outputRow.TrabajadorApellidoNombre = ($"{nombre1} {nombre2}").Trim();
                outputRow.TrabajadorApellidoNombre = outputRow.TrabajadorApellidoNombre.Trim();
                
                outputRow.Riesgo = GetCellValue(dataRow, columnMap,
                    "Riesgo", "Comentarios", "Comentario", "Descripcion Riesgo", "Descripción Riesgo");

                // Para formatos nuevos de anuales/semestrales la prestación viene en la
                // columna "Examen" (columna I). Agregamos "Examen" como alias y luego
                // limpiamos el código inicial ("C06 - ", "C13 - " , etc.).
                var prest1 = GetCellValue(dataRow, columnMap,
                    "Práctica solicitada", "Practica solicitada", "Prestación", "Prestacion", "Examen");
                var prest2 = GetCellValue(dataRow, columnMap,
                    "Práctica", "Practica");
                var prestRaw = ($"{prest1} {prest2}").Trim();
                outputRow.Prestacion = CleanPrestacion(prestRaw);

                result.Rows.Add(outputRow);
            }
            catch (Exception ex)
            {
                result.Warnings.Add($"Error en solapa '{worksheet.Name}' fila {rowNum}: {ex.Message}");
            }
        }
    }

    private IXLRow? FindHeaderRow(IXLWorksheet worksheet)
    {
        // Buscar en las primeras 20 filas
        for (int i = 1; i <= Math.Min(20, worksheet.LastRowUsed()?.RowNumber() ?? 20); i++)
        {
            var row = worksheet.Row(i);
            var cellValues = row.Cells().Select(c => c.GetString().ToUpper()).ToList();

            // Buscar indicadores típicos de encabezado
            if (cellValues.Any(v => v.Contains("CUIL") || v.Contains("BENEFICIARIO") || 
                                   v.Contains("PRESTAC") || v.Contains("RIESGO")))
            {
                return row;
            }
        }

        return null;
    }

    private Dictionary<string, int> MapColumnsNormalized(IXLWorksheet worksheet, int headerRowNum)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var headerRow = worksheet.Row(headerRowNum);

        foreach (var cell in headerRow.CellsUsed())
        {
            var headerText = cell.GetString().Trim();
            if (!string.IsNullOrWhiteSpace(headerText))
            {
                var norm = NormalizeHeader(headerText);
                if (!map.ContainsKey(norm))
                {
                    map[norm] = cell.Address.ColumnNumber;
                }
            }
        }

        return map;
    }

    private string GetCellValue(IXLRow row, Dictionary<string, int> columnMap, params string[] candidates)
    {
        var idx = FindColumnIndex(columnMap, candidates);
        if (idx > 0)
        {
            return row.Cell(idx).GetString().Trim();
        }
        return string.Empty;
    }

    private bool IsEmptyRow(IXLRow row)
    {
        return !row.CellsUsed().Any();
    }

    private string NormalizeHeader(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;

        var normalized = text.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder();
        foreach (var c in normalized)
        {
            var uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
            if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                sb.Append(c);
        }

        var cleaned = sb.ToString().Normalize(NormalizationForm.FormC);
        var only = new StringBuilder();
        foreach (var c in cleaned)
        {
            if (char.IsLetterOrDigit(c) || char.IsWhiteSpace(c))
                only.Append(c);
        }

        return only.ToString().Trim().ToLowerInvariant();
    }

    private int FindColumnIndex(Dictionary<string, int> normalizedMap, params string[] candidates)
    {
        foreach (var cand in candidates)
        {
            var n = NormalizeHeader(cand);
            if (normalizedMap.TryGetValue(n, out var idx))
                return idx;
        }

        // Fallback: búsqueda parcial
        foreach (var kv in normalizedMap)
        {
            foreach (var cand in candidates)
            {
                var n = NormalizeHeader(cand);
                if (kv.Key.Contains(n) || n.Contains(kv.Key))
                    return kv.Value;
            }
        }

        return -1;
    }

    private (string Contrato, string CUIT, string RazonSocial, string NroEstablecimiento,
             string Localidad, string Provincia, string Telefono, string CodPostal) ExtractCompanyDataFromSheet(IXLWorksheet worksheet, int headerRowNumber, string? debugCpPath)
    {
        var contrato = string.Empty;
        var cuit = string.Empty;
        var razonSocial = string.Empty;
        var nroEstablecimiento = string.Empty;
        var localidad = string.Empty;
        var provincia = string.Empty;
        var telefono = string.Empty;
        var codPostal = string.Empty;

        // Buscar en la zona superior, pero evitando la fila de encabezados (para no capturar
        // columnas como "Fecha Solicitud" cuando está al lado de "Nro Establecimiento").
        var last = worksheet.LastRowUsed()?.RowNumber() ?? 15;
        var maxRowToScan = Math.Min(15, last);
        if (headerRowNumber > 1)
            maxRowToScan = Math.Min(maxRowToScan, headerRowNumber - 1);

        for (int i = 1; i <= maxRowToScan; i++)
        {
            var row = worksheet.Row(i);
            
            foreach (var cell in row.CellsUsed())
            {
                var value = cell.GetString().ToUpper();

                if (value.Contains("CONTRATO") || value.Contains("NRO. CONTRATO"))
                {
                    // Buscar valor en la celda siguiente
                    var nextCell = cell.CellRight();
                    contrato = nextCell.GetString().Trim();
                }
                else if (value.Contains("CUIT"))
                {
                    var nextCell = cell.CellRight();
                    cuit = nextCell.GetString().Trim();
                }
                else if (value.Contains("RAZÓN SOCIAL") || value.Contains("RAZON SOCIAL") || value.Contains("EMPLEADOR"))
                {
                    var nextCell = cell.CellRight();
                    razonSocial = nextCell.GetString().Trim();
                }
                else if (value.Contains("ESTABLECIMIENTO"))
                {
                    var nextCell = cell.CellRight();
                    var candidate = nextCell.GetString().Trim();
                    // Evitar capturar encabezados/labels de otras columnas (p.ej. "Fecha Solicitud")
                    if (!candidate.Contains("FECHA", StringComparison.OrdinalIgnoreCase) &&
                        !candidate.Contains("SOLICITUD", StringComparison.OrdinalIgnoreCase))
                    {
                        nroEstablecimiento = candidate;
                    }
                }
                else if (value.Contains("LOCALIDAD"))
                {
                    var nextCell = cell.CellRight();
                    localidad = nextCell.GetString().Trim();
                }
                else if (value.Contains("PROVINCIA"))
                {
                    var nextCell = cell.CellRight();
                    provincia = nextCell.GetString().Trim();
                }
                else if (value.Contains("TELÉFONO") || value.Contains("TELEFONO"))
                {
                    var nextCell = cell.CellRight();
                    telefono = nextCell.GetString().Trim();
                }
            }
        }

        // Leer CodPostal desde la hoja "Resumen" usando el número de solapa
        try
        {
            var workbook = worksheet.Workbook;
            var resumenSheet = workbook.Worksheets
                .FirstOrDefault(ws => ws.Name.Contains("Resumen", StringComparison.OrdinalIgnoreCase));

            if (resumenSheet != null)
            {
                // Índice de esta hoja entre las hojas de datos (excluyendo la de resumen), 1-based
                var dataSheets = workbook.Worksheets
                    .Where(ws => !ws.Name.Contains("Resumen", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(ws => ws.Position)
                    .ToList();

                var dataIndex = dataSheets.FindIndex(ws => ws.Name.Equals(worksheet.Name, StringComparison.OrdinalIgnoreCase)) + 1;
                if (dataIndex <= 0)
                    dataIndex = worksheet.Position; // fallback

                // La fila de encabezados en la hoja Resumen no siempre es la 1;
                // en tu archivo suele ser la 2 (fila 1 tiene el título "Resumen de solicitudes").
                int headerRowResumen = 1;
                var lastRowResumen = resumenSheet.LastRowUsed()?.RowNumber() ?? 5;
                for (int r = 1; r <= Math.Min(lastRowResumen, 10); r++)
                {
                    var row = resumenSheet.Row(r);
                    var texts = row.CellsUsed().Select(c => c.GetString()).ToList();
                    if (texts.Any(t => t.Contains("Solapa", StringComparison.OrdinalIgnoreCase)))
                    {
                        headerRowResumen = r;
                        break;
                    }
                }

                var header = resumenSheet.Row(headerRowResumen);
                int colSolapa = 0;
                int colLocalidad = 0;

                foreach (var cell in header.CellsUsed())
                {
                    var text = cell.GetString();
                    if (colSolapa == 0 && text.Contains("Solapa", StringComparison.OrdinalIgnoreCase))
                        colSolapa = cell.Address.ColumnNumber;
                    else if (colLocalidad == 0 && text.Contains("Localidad", StringComparison.OrdinalIgnoreCase))
                        colLocalidad = cell.Address.ColumnNumber;
                }

                if (colSolapa > 0 && colLocalidad > 0)
                {
                    if (!string.IsNullOrEmpty(debugCpPath))
                    {
                        try
                        {
                            File.AppendAllText(debugCpPath,
                                $"Resumen headerRow={headerRowResumen} colSolapa={colSolapa} colLocalidad={colLocalidad} dataIndex={dataIndex}{Environment.NewLine}");
                        }
                        catch { }
                    }

                    var lastDataRowResumen = resumenSheet.LastRowUsed()?.RowNumber() ?? (headerRowResumen + 1);
                    for (int r = headerRowResumen + 1; r <= lastDataRowResumen; r++)
                    {
                        var solapaText = resumenSheet.Cell(r, colSolapa).GetString().Trim();
                        if (!int.TryParse(solapaText, out var solapaNum))
                            continue;

                        // La columna "Solapa" suele numerar solo las hojas de datos (1,2,3...),
                        // por eso la comparamos con el índice dentro de dataSheets.
                        if (solapaNum != dataIndex)
                            continue;

                        var locRaw = resumenSheet.Cell(r, colLocalidad).GetString().Trim();
                        var match = Regex.Match(locRaw, @"^\((\d{3,5})\)\s*(.+)$");
                        if (match.Success)
                        {
                            codPostal = match.Groups[1].Value.Trim();
                            var locText = match.Groups[2].Value.Trim();
                            if (string.IsNullOrWhiteSpace(localidad))
                                localidad = locText;

                            if (!string.IsNullOrEmpty(debugCpPath))
                            {
                                try
                                {
                                    File.AppendAllText(debugCpPath,
                                        $"Hoja='{worksheet.Name}' Solapa={solapaNum} dataIndex={dataIndex} filaResumen={r} LocalidadRaw='{locRaw}' CodPostal='{codPostal}' Localidad='{localidad}'{Environment.NewLine}");
                                }
                                catch { }
                            }
                        }

                        break;
                    }
                }
            }
        }
        catch
        {
            // Si falla algo al leer el resumen, continuar con los datos ya obtenidos
        }

        // Heurística adicional: en algunos formatos, en C2 viene "2- RAZON SOCIAL ...".
        // Tomar el número antes del guion como NroEstablecimiento.
        try
        {
            var c2 = worksheet.Cell(2, 3).GetString().Trim();
            if (TryParseNroEstablecimiento(c2, out var nroFromC2, out var razonFromC2))
            {
                nroEstablecimiento = nroFromC2;
                if (string.IsNullOrWhiteSpace(razonSocial) && !string.IsNullOrWhiteSpace(razonFromC2))
                    razonSocial = razonFromC2;
            }
        }
        catch
        {
            // Ignorar si la hoja no tiene C2 utilizable
        }

        return (contrato, cuit, razonSocial, nroEstablecimiento, localidad, provincia, telefono, codPostal);
    }

    private static bool TryParseNroEstablecimiento(string text, out string nro, out string razon)
    {
        nro = string.Empty;
        razon = string.Empty;

        if (string.IsNullOrWhiteSpace(text))
            return false;

        var match = Regex.Match(text, @"^\s*(\d+)\s*[-–]\s*(.+)$");
        if (!match.Success)
            return false;

        nro = match.Groups[1].Value.Trim();
        razon = match.Groups[2].Value.Trim();
        return !string.IsNullOrWhiteSpace(nro);
    }

    /// <summary>
    /// Limpia el texto de prestación quitando el código inicial tipo
    /// "C06 - EXAMEN CLINICO" o "80004: EXAMEN" y dejando solo el nombre.
    /// </summary>
    private static string CleanPrestacion(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var text = value.Trim();

        // Quitar sufijo "cod: XXX" si viniera de otros formatos
        var codMatch = Regex.Match(text, @"\s+cod:\s*[A-Z0-9]+\s*$", RegexOptions.IgnoreCase);
        if (codMatch.Success)
        {
            text = text.Substring(0, codMatch.Index).Trim();
        }

        // Quitar prefijo de código: letras/números + '-' o ':' y luego el nombre
        var prefixMatch = Regex.Match(text, @"^\s*[A-Z0-9]{2,10}\s*[-:]\s+(.+)$");
        if (prefixMatch.Success)
        {
            text = prefixMatch.Groups[1].Value.Trim();
        }

        return text;
    }
}
