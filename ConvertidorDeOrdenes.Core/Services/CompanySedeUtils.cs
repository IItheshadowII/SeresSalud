using ConvertidorDeOrdenes.Core.Models;
using System.Text.RegularExpressions;

namespace ConvertidorDeOrdenes.Core.Services;

public static class CompanySedeUtils
{
    public static string ComputeSedeKey(CompanyRecord company)
        => ComputeSedeKey(company.Calle, company.Localidad, company.Provincia);

    public static string ComputeSedeKey(string? calle, string? localidad, string? provincia)
    {
        // La "sede" se define por domicilio + localidad + provincia.
        // Normalizamos para comparar de forma estable.
        var calleKey = NormalizeStreetPart(calle);
        var locKey = NormalizeLocalidadPart(localidad);
        var provKey = NormalizeProvinciaPart(provincia);

        return $"{calleKey}|{locKey}|{provKey}";
    }

    public static bool IsSameSede(CompanyRecord left, CompanyRecord right)
        => ComputeSedeKey(left).Equals(ComputeSedeKey(right), StringComparison.OrdinalIgnoreCase);

    public static bool HasComparableSedeData(CompanyRecord company)
        => HasComparableSedeData(company.Calle, company.Localidad, company.Provincia);

    public static bool HasComparableSedeData(string? calle, string? localidad, string? provincia)
    {
        return !string.IsNullOrWhiteSpace(NormalizeStreetPart(calle)) &&
               !string.IsNullOrWhiteSpace(NormalizeLocalidadPart(localidad)) &&
               !string.IsNullOrWhiteSpace(NormalizeProvinciaPart(provincia));
    }

    public static string NormalizeStreetPart(string? street)
        => NormalizeKeyPart(street);

    public static string NormalizeLocalidadPart(string? localidad)
    {
        if (string.IsNullOrWhiteSpace(localidad))
            return string.Empty;

        var normalized = localidad.Trim();
        normalized = Regex.Replace(normalized, @"^\(\d+\)\s*", string.Empty);
        normalized = Regex.Replace(normalized, @"[-\s]+(B\s*A|BUENOS\s+AIRES)\s*$", string.Empty, RegexOptions.IgnoreCase);
        return NormalizeKeyPart(normalized);
    }

    public static string NormalizeProvinciaPart(string? provincia)
    {
        var normalized = NormalizeKeyPart(provincia);

        return normalized switch
        {
            "BA" => "BUENOS AIRES",
            "B A" => "BUENOS AIRES",
            "BS AS" => "BUENOS AIRES",
            "BS AS." => "BUENOS AIRES",
            "BS. AS." => "BUENOS AIRES",
            "BSAS" => "BUENOS AIRES",
            _ => normalized
        };
    }

    private static string NormalizeKeyPart(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        text = text.Trim().ToUpperInvariant();

        // Colapsar espacios múltiples
        var collapsed = new System.Text.StringBuilder();
        bool lastWasSpace = false;
        foreach (var c in text)
        {
            var isSpace = char.IsWhiteSpace(c);
            if (isSpace)
            {
                if (!lastWasSpace)
                    collapsed.Append(' ');
                lastWasSpace = true;
            }
            else
            {
                collapsed.Append(c);
                lastWasSpace = false;
            }
        }

        return collapsed.ToString().Trim();
    }
}
