using ConvertidorDeOrdenes.Core.Models;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.Text;

namespace ConvertidorDeOrdenes.Core.Parsers;

/// <summary>
/// Parser para archivos CSV de Reconfirmatorios/Reevaluaciones
/// </summary>
public class CsvOrderParser
{
    public ParseResult Parse(string filePath, string frecuencia, string referente)
    {
        var result = new ParseResult();

        try
        {
            // Registrar codepage para encoding latin-1
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Leer las primeras líneas para detectar delimitador y estructura
            var rawLines = File.ReadAllLines(filePath, Encoding.GetEncoding("iso-8859-1"));
            if (rawLines.Length == 0)
            {
                return result;
            }

            // Detectar delimitador más probable
            char detectedDelimiter = DetectDelimiter(rawLines.Take(10).ToArray());

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = detectedDelimiter.ToString(),
                HasHeaderRecord = true,
                TrimOptions = TrimOptions.Trim,
                BadDataFound = null,
                MissingFieldFound = null
            };

            using var reader = new StringReader(string.Join(Environment.NewLine, rawLines));
            using var csv = new CsvReader(reader, config);

            // Intentar lectura normal
            try
            {
                csv.Read();
                csv.ReadHeader();
            }
            catch
            {
                // Si falla la lectura del header, intentar parseo manual
            }

            var headers = csv.HeaderRecord ?? Array.Empty<string>();

            // Si el header se detectó como una única celda que contiene comas (todo en una columna),
            // reconstruir encabezados y filas usando splitter robusto
            if (headers.Length == 1 && rawLines.Length > 0 && rawLines[0].Contains(detectedDelimiter))
            {
                var headerLine = rawLines[0];
                var headerParts = SplitCsvLine(headerLine, detectedDelimiter).ToArray();
                var columnMap = MapColumnsNormalized(headerParts);

                // Procesar cada data line a partir de la segunda
                for (int i = 1; i < rawLines.Length; i++)
                {
                    var line = rawLines[i];
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var fields = SplitCsvLine(line, detectedDelimiter).ToArray();

                    try
                    {
                        var row = new OutputRow { Frecuencia = frecuencia, Referente = referente };

                        row.Contrato = GetField(fields, columnMap, "Contrato");
                        row.CuitEmpleador = GetField(fields, columnMap, "CUIT");
                        var nroEstablecimiento = GetField(fields, columnMap, "Nro. Establecimiento", "Nro Establecimiento", "Establecimiento");
                        var nombreEstablecimiento = GetField(fields, columnMap, "Nombre establecimiento", "Nombre Establecimiento", "Establecimiento");
                        ProcessEmpleador(nombreEstablecimiento, ref row);
                        if (string.IsNullOrWhiteSpace(row.NroEstablecimiento)) row.NroEstablecimiento = nroEstablecimiento;
                        row.Telefono = GetField(fields, columnMap, "Teléfono", "Telefono", "Teléfono / Celular", "Telefono / Celular");
                        row.Cuil = GetField(fields, columnMap, "CUIL");
                        row.TrabajadorApellidoNombre = GetField(fields, columnMap, "Nombre Beneficiario", "Beneficiario", "Apellido y Nombre", "Nombre y Apellido");
                        row.Prestacion = ExtractPrestacionName(GetField(fields, columnMap, "Práctica solicitada", "Practica solicitada", "Práctica", "Practica"));
                        row.Riesgo = GetField(fields, columnMap, "Comentarios", "Comentario", "Riesgo");
                        row.Localidad = GetField(fields, columnMap, "Localidad", "Localidad / Domicilio");
                        row.Provincia = GetField(fields, columnMap, "Provincia", "Provincia / Domicilio");
                        var emailBeneficiario = GetField(fields, columnMap, "Email Beneficiario", "Email Beneficiario", "Email");
                        var telCelular = GetField(fields, columnMap, "Teléfono / Celular", "Telefono / Celular", "Teléfono", "Telefono");
                        if (string.IsNullOrWhiteSpace(emailBeneficiario) || emailBeneficiario == "-")
                            row.Mail = GetField(fields, columnMap, "Email Agencia", "Email Agencia", "Email Agencia ");
                        else
                            row.Mail = emailBeneficiario;
                        if ((string.IsNullOrWhiteSpace(telCelular) || telCelular == "-" || telCelular.Trim() == "- -") && string.IsNullOrWhiteSpace(row.Telefono))
                            row.Telefono = GetField(fields, columnMap, "Teléfono Agencia", "Telefono Agencia");

                        result.Rows.Add(row);
                    }
                    catch (Exception ex)
                    {
                        result.Warnings.Add($"Error parseando fila (manual) {i + 1}: {ex.Message}");
                    }
                }

                return result;
            }

            // Mapeo de columnas por nombre (tolerante a variaciones)
            var columnMap2 = MapColumnsNormalized(headers);

            while (csv.Read())
            {
                try
                {
                    var row = new OutputRow
                    {
                        Frecuencia = frecuencia,
                        Referente = referente
                    };

                    // Extraer datos del CSV
                    row.Contrato = GetField(csv, columnMap2, "Contrato");
                    row.CuitEmpleador = GetField(csv, columnMap2, "CUIT");

                    var nroEstablecimiento = GetField(csv, columnMap2, "Nro. Establecimiento");
                    var nombreEstablecimiento = GetField(csv, columnMap2, "Nombre establecimiento");

                    // Procesar empleador (puede venir con formato "NRO - NOMBRE")
                    ProcessEmpleador(nombreEstablecimiento, ref row);

                    // Si no se extrajo número de establecimiento del nombre, usar columna dedicada
                    if (string.IsNullOrWhiteSpace(row.NroEstablecimiento))
                    {
                        row.NroEstablecimiento = nroEstablecimiento;
                    }

                    row.Telefono = GetField(csv, columnMap2, "Teléfono");
                    row.Cuil = GetField(csv, columnMap2, "CUIL");
                    row.TrabajadorApellidoNombre = GetField(csv, columnMap2, "Nombre Beneficiario");

                    var practicaSolicitada = GetField(csv, columnMap2, "Práctica solicitada");
                    row.Prestacion = ExtractPrestacionName(practicaSolicitada);

                    row.Riesgo = GetField(csv, columnMap2, "Comentarios");
                    row.Localidad = GetField(csv, columnMap2, "Localidad");
                    row.Provincia = GetField(csv, columnMap2, "Provincia");

                    var emailBeneficiario = GetField(csv, columnMap2, "Email Beneficiario");
                    var telCelular = GetField(csv, columnMap2, "Teléfono / Celular");

                    // Usar email/teléfono de agencia si no hay del beneficiario
                    if (string.IsNullOrWhiteSpace(emailBeneficiario) || emailBeneficiario == "-")
                    {
                        row.Mail = GetField(csv, columnMap2, "Email Agencia");
                    }
                    else
                    {
                        row.Mail = emailBeneficiario;
                    }

                    if (string.IsNullOrWhiteSpace(telCelular) || telCelular == "-" || telCelular.Trim() == "- -")
                    {
                        var telAgencia = GetField(csv, columnMap2, "Teléfono Agencia");
                        if (!string.IsNullOrWhiteSpace(row.Telefono))
                        {
                            // Ya tenemos teléfono
                        }
                        else
                        {
                            row.Telefono = telAgencia;
                        }
                    }

                    result.Rows.Add(row);
                }
                catch (Exception ex)
                {
                    result.Warnings.Add($"Error parseando fila: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            result.Errors.Add($"Error leyendo archivo CSV: {ex.Message}");
        }

        return result;
    }

    private Dictionary<string, int> MapColumns(string[] headers)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < headers.Length; i++)
        {
            var header = headers[i].Trim();
            map[header] = i;
        }

        return map;
    }

    private Dictionary<string,int> MapColumnsNormalized(string[] headers)
    {
        var map = new Dictionary<string,int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < headers.Length; i++)
        {
            var norm = NormalizeHeader(headers[i]);
            if (!map.ContainsKey(norm))
                map[norm] = i;
        }
        return map;
    }

    private string NormalizeHeader(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;
        // quitar acentos
        var normalized = text.Normalize(System.Text.NormalizationForm.FormD);
        var sb = new System.Text.StringBuilder();
        foreach (var c in normalized)
        {
            var uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
            if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                sb.Append(c);
        }

        var cleaned = sb.ToString().Normalize(System.Text.NormalizationForm.FormC);
        // mantener solo letras y dígitos y espacios
        var only = new System.Text.StringBuilder();
        foreach (var c in cleaned)
        {
            if (char.IsLetterOrDigit(c) || char.IsWhiteSpace(c)) only.Append(c);
        }

        return only.ToString().Trim().ToLowerInvariant();
    }

    private int FindColumnIndex(Dictionary<string,int> normalizedMap, params string[] candidates)
    {
        foreach (var cand in candidates)
        {
            var n = NormalizeHeader(cand);
            if (normalizedMap.TryGetValue(n, out var idx)) return idx;
        }

        // fallback: try partial contains
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

    private string GetField(string[] fields, Dictionary<string, int> normalizedMap, params string[] candidates)
    {
        var idx = FindColumnIndex(normalizedMap, candidates);
        if (idx >= 0 && idx < fields.Length)
            return fields[idx]?.Trim() ?? string.Empty;
        return string.Empty;
    }

    private static char DetectDelimiter(string[] sampleLines)
    {
        var delimiters = new[] { ',', ';', '\t', '|' };
        char best = ',';
        int bestCount = -1;

        foreach (var d in delimiters)
        {
            int count = 0;
            foreach (var line in sampleLines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                count += line.Count(c => c == d);
            }

            if (count > bestCount)
            {
                bestCount = count;
                best = d;
            }
        }

        return best;
    }

    private static IEnumerable<string> SplitCsvLine(string line, char delimiter)
    {
        if (line == null) yield break;

        var cells = new List<string>();
        var cur = new System.Text.StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            var c = line[i];
            if (c == '"')
            {
                // Double quote escape
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    cur.Append('"');
                    i++; // skip next
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (c == delimiter && !inQuotes)
            {
                cells.Add(cur.ToString());
                cur.Clear();
            }
            else
            {
                cur.Append(c);
            }
        }

        cells.Add(cur.ToString());

        foreach (var cell in cells)
            yield return cell.Trim().Trim('"');
    }

    private string GetField(CsvReader csv, Dictionary<string, int> normalizedMap, params string[] candidates)
    {
        var idx = FindColumnIndex(normalizedMap, candidates);
        if (idx >= 0)
        {
            try
            {
                return csv.GetField(idx)?.Trim() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
        return string.Empty;
    }

    private string ExtractPrestacionName(string practicaSolicitada)
    {
        if (string.IsNullOrWhiteSpace(practicaSolicitada))
            return string.Empty;

        // Remover "cod: XXX" al final
        var codIndex = practicaSolicitada.LastIndexOf(" cod:", StringComparison.OrdinalIgnoreCase);
        if (codIndex > 0)
        {
            return practicaSolicitada.Substring(0, codIndex).Trim();
        }

        return practicaSolicitada.Trim();
    }

    private void ProcessEmpleador(string empleador, ref OutputRow row)
    {
        if (string.IsNullOrWhiteSpace(empleador))
            return;

        // Patrón: "NRO - NOMBRE"
        var match = System.Text.RegularExpressions.Regex.Match(empleador, @"^(\d+)\s*-\s*(.+)$");
        if (match.Success)
        {
            row.NroEstablecimiento = match.Groups[1].Value;
            row.Empleador = match.Groups[2].Value.Trim();
        }
        else
        {
            row.Empleador = empleador.Trim();
        }
    }
}
