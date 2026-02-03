using ConvertidorDeOrdenes.Core.Models;

namespace ConvertidorDeOrdenes.Core.Services;

public static class CompanySedeUtils
{
    public static string ComputeSedeKey(CompanyRecord company)
        => ComputeSedeKey(company.Calle, company.Localidad, company.Provincia);

    public static string ComputeSedeKey(string? calle, string? localidad, string? provincia)
    {
        // La "sede" se define por domicilio + localidad + provincia.
        // Normalizamos para comparar de forma estable.
        var calleKey = NormalizeKeyPart(calle);
        var locKey = NormalizeKeyPart(localidad);
        var provKey = NormalizeKeyPart(provincia);

        return $"{calleKey}|{locKey}|{provKey}";
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
