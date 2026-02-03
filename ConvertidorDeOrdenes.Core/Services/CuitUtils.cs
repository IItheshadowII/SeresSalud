namespace ConvertidorDeOrdenes.Core.Services;

public static class CuitUtils
{
    public static string ExtractDigits(string? cuit)
    {
        if (string.IsNullOrWhiteSpace(cuit))
            return string.Empty;

        return new string(cuit.Where(char.IsDigit).ToArray());
    }

    public static bool IsValid11Digits(string? cuit)
    {
        var digits = ExtractDigits(cuit);
        return digits.Length == 11;
    }

    /// <summary>
    /// Formatea un CUIT/CUIL de 11 dígitos como xx-xxxxxxxx-x.
    /// Si no tiene exactamente 11 dígitos, devuelve el valor original trimmeado.
    /// </summary>
    public static string FormatOrKeep(string? cuit)
    {
        if (string.IsNullOrWhiteSpace(cuit))
            return string.Empty;

        var trimmed = cuit.Trim();
        var digits = ExtractDigits(trimmed);
        if (digits.Length != 11)
            return trimmed;

        return $"{digits[..2]}-{digits.Substring(2, 8)}-{digits[^1]}";
    }
}
