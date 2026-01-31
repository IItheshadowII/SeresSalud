using ConvertidorDeOrdenes.Core.Models;
using System.Text.RegularExpressions;

namespace ConvertidorDeOrdenes.Core.Services;

/// <summary>
/// Servicio para normalizar y limpiar datos
/// </summary>
public class Normalizer
{
    private readonly PrestacionMapper _prestacionMapper;

    public Normalizer(PrestacionMapper prestacionMapper)
    {
        _prestacionMapper = prestacionMapper;
    }

    /// <summary>
    /// Normaliza una fila completa de salida
    /// </summary>
    public void NormalizeRow(OutputRow row, List<string> warnings)
    {
        // Normalizar prestación
        row.Prestacion = NormalizePrestacion(row.Prestacion, warnings);

        // Regla específica: en frecuencia R, el riesgo debe ser el texto de la prestación
        if (string.Equals(row.Frecuencia?.Trim(), "R", StringComparison.OrdinalIgnoreCase))
        {
            row.Riesgo = row.Prestacion ?? string.Empty;
        }

        // Normalizar empleador y número de establecimiento
        NormalizeEmpleador(row);

        // Extraer código postal desde localidad si viene como "(6034) LOCALIDAD-B A"
        ExtractCodPostalFromLocalidad(row);

        // Normalizar localidad
        row.Localidad = NormalizeLocalidad(row.Localidad);

        // Normalizar provincia
        row.Provincia = NormalizeProvincia(row.Provincia);

        // Truncar riesgo / descripción de riesgo si exceden 90 caracteres
        if (!string.IsNullOrEmpty(row.Riesgo) && row.Riesgo.Length > 90)
        {
            warnings.Add($"Riesgo truncado de {row.Riesgo.Length} a 90 caracteres: {row.Riesgo.Substring(0, 30)}...");
            row.Riesgo = row.Riesgo.Substring(0, 90);
        }

        if (!string.IsNullOrEmpty(row.DescripcionRiesgo) && row.DescripcionRiesgo.Length > 90)
        {
            warnings.Add($"Descripción de riesgo truncada de {row.DescripcionRiesgo.Length} a 90 caracteres: {row.DescripcionRiesgo.Substring(0, 30)}...");
            row.DescripcionRiesgo = row.DescripcionRiesgo.Substring(0, 90);
        }
    }

    /// <summary>
    /// Normaliza prestación: quita "cod: XXX", elimina acentos, aplica mapeo
    /// </summary>
    public string NormalizePrestacion(string prestacion, List<string> warnings)
    {
        if (string.IsNullOrWhiteSpace(prestacion))
            return string.Empty;

        var normalized = prestacion.Trim();

        // Quitar "cod: XXX" al final (si viene de formatos viejos)
        var codMatch = Regex.Match(normalized, @"\s+cod:\s*[A-Z0-9]+\s*$", RegexOptions.IgnoreCase);
        if (codMatch.Success)
        {
            normalized = normalized.Substring(0, codMatch.Index).Trim();
        }

        // Quitar código al inicio tipo "C06 - EXAMEN CLINICO" o "80004: ALGO".
        // Dejamos solo el nombre de la prestación.
        var codePrefixMatch = Regex.Match(normalized, @"^\s*[A-Z0-9]{2,10}\s*[-:]\s+(.+)$");
        if (codePrefixMatch.Success)
        {
            normalized = codePrefixMatch.Groups[1].Value.Trim();
        }

        // Eliminar acentos
        normalized = RemoveAccents(normalized);

        // Aplicar mapeo si existe
        var mapped = _prestacionMapper.Map(normalized);
        if (mapped != normalized)
        {
            // Se aplicó mapeo
            normalized = mapped;
        }

        return normalized;
    }

    /// <summary>
    /// Si la localidad viene con formato "(6034) JUAN BAUTISTA ALBERDI-B A",
    /// extrae el código postal al campo CodPostal y deja solo la localidad.
    /// </summary>
    private void ExtractCodPostalFromLocalidad(OutputRow row)
    {
        if (!string.IsNullOrWhiteSpace(row.CodPostal))
            return;

        if (string.IsNullOrWhiteSpace(row.Localidad))
            return;

        var match = Regex.Match(row.Localidad, "^\\((\\d{3,5})\\)\\s*(.+)$");
        if (!match.Success)
            return;

        row.CodPostal = match.Groups[1].Value.Trim();
        row.Localidad = match.Groups[2].Value.Trim();
    }

    /// <summary>
    /// Normaliza empleador: separa "NRO - NOMBRE" si aplica
    /// </summary>
    private void NormalizeEmpleador(OutputRow row)
    {
        if (string.IsNullOrWhiteSpace(row.Empleador))
            return;

        // Patrón: "NRO - NOMBRE" (solo si empieza con dígitos)
        var match = Regex.Match(row.Empleador, @"^(\d+)\s*-\s*(.+)$");
        if (match.Success)
        {
            row.NroEstablecimiento = match.Groups[1].Value;
            row.Empleador = match.Groups[2].Value.Trim();
        }
    }

    /// <summary>
    /// Normaliza localidad: limpia CP y sufijos "BA", "B A", "BUENOS AIRES"
    /// </summary>
    public string NormalizeLocalidad(string localidad)
    {
        if (string.IsNullOrWhiteSpace(localidad))
            return string.Empty;

        var normalized = localidad.Trim();

        // Quitar CP entre paréntesis: "(1605) MUNRO-B A" -> "MUNRO-B A"
        normalized = Regex.Replace(normalized, @"^\(\d+\)\s*", "");

        // Quitar sufijos de provincia al final
        // Patrones: " BA", " B A", " BUENOS AIRES", "-BA", "-B A", etc.
        normalized = Regex.Replace(normalized, @"[-\s]+(B\s*A|BUENOS\s+AIRES)\s*$", "", RegexOptions.IgnoreCase);

        return normalized.Trim().ToUpper();
    }

    /// <summary>
    /// Normaliza provincia: convierte variantes a nombre completo
    /// </summary>
    public string NormalizeProvincia(string provincia)
    {
        if (string.IsNullOrWhiteSpace(provincia))
            return string.Empty;

        var normalized = provincia.Trim().ToUpper();

        // Diccionario de normalizaciones
        var mappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "BA", "BUENOS AIRES" },
            { "B A", "BUENOS AIRES" },
            { "BS AS", "BUENOS AIRES" },
            { "BS. AS.", "BUENOS AIRES" },
            { "BSAS", "BUENOS AIRES" },
            { "CF", "CAPITAL FEDERAL" },
            { "C.F.", "CAPITAL FEDERAL" },
            { "CABA", "CAPITAL FEDERAL" },
            { "CDAD. DE BS AS", "CAPITAL FEDERAL" },
            { "CIUDAD DE BUENOS AIRES", "CAPITAL FEDERAL" },
            { "CBA", "CORDOBA" },
            { "STA FE", "SANTA FE" },
            { "SF", "SANTA FE" },
            { "MZA", "MENDOZA" },
            { "TUC", "TUCUMAN" },
            { "SDE", "SANTIAGO DEL ESTERO" },
            { "STGO DEL ESTERO", "SANTIAGO DEL ESTERO" },
            { "SL", "SAN LUIS" },
            { "SJ", "SAN JUAN" }
        };

        if (mappings.TryGetValue(normalized, out var mapped))
        {
            return mapped;
        }

        return normalized;
    }

    /// <summary>
    /// Elimina acentos de una cadena
    /// </summary>
    private string RemoveAccents(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        var normalizedString = text.Normalize(System.Text.NormalizationForm.FormD);
        var stringBuilder = new System.Text.StringBuilder();

        foreach (var c in normalizedString)
        {
            var unicodeCategory = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
            if (unicodeCategory != System.Globalization.UnicodeCategory.NonSpacingMark)
            {
                stringBuilder.Append(c);
            }
        }

        return stringBuilder.ToString().Normalize(System.Text.NormalizationForm.FormC);
    }
}
