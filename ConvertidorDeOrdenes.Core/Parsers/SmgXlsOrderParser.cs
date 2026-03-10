using ConvertidorDeOrdenes.Core.Models;
using ExcelDataReader;
using System.Data;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ConvertidorDeOrdenes.Core.Parsers;

/// <summary>
/// Parser para archivos XLS (SMG)
/// </summary>
public class SmgXlsOrderParser
{
    public ParseResult Parse(string filePath, string frecuencia)
    {
        var result = new ParseResult();

        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = false
                }
            });

            if (dataSet.Tables.Count == 0)
            {
                result.Errors.Add("El archivo XLS no contiene hojas.");
                return result;
            }

            var table = dataSet.Tables[0];
            var headerRowIndex = FindHeaderRowIndex(table);
            if (headerRowIndex < 0)
            {
                result.Errors.Add("No se encontro la columna 'CUIL' en el archivo.");
                return result;
            }

            var empresaData = ExtractEmpresaData(table);
            var dataStartRow = headerRowIndex + 1;

            for (int rowIndex = dataStartRow; rowIndex < table.Rows.Count; rowIndex++)
            {
                var row = table.Rows[rowIndex];
                if (row == null || row.ItemArray.Length < 8)
                    continue;

                var apellido = GetCellString(row, 5);
                var nombre = GetCellString(row, 6);
                var nombreCompleto = string.Join(" ", new[] { apellido, nombre }.Where(x => !string.IsNullOrWhiteSpace(x)));

                var cuil = GetCellNumericString(row, 7);
                if (string.IsNullOrWhiteSpace(cuil) || cuil == "0")
                    continue;

                var prestacion = GetCellNumericString(row, 3);

                var outputRow = new OutputRow
                {
                    CuitEmpleador = empresaData.CuitEmpleador,
                    Empleador = empresaData.Empleador,
                    Calle = empresaData.Calle,
                    Localidad = empresaData.Localidad,
                    Provincia = empresaData.Provincia,
                    Telefono = empresaData.Telefono,
                    Contrato = empresaData.Contrato,
                    NroEstablecimiento = empresaData.NroEstablecimiento,
                    Frecuencia = frecuencia,
                    Cuil = cuil,
                    TrabajadorApellidoNombre = nombreCompleto,
                    Prestacion = prestacion,
                    Mail = empresaData.Mail,
                    Referente = empresaData.Referente
                };

                result.Rows.Add(outputRow);
            }
        }
        catch (Exception ex)
        {
            result.Errors.Add($"Error leyendo archivo XLS: {ex.Message}");
        }

        return result;
    }

    private static int FindHeaderRowIndex(DataTable table)
    {
        for (int i = 0; i < table.Rows.Count; i++)
        {
            var row = table.Rows[i];
            foreach (var cell in row.ItemArray)
            {
                var text = (cell?.ToString() ?? string.Empty).Trim();
                if (text.Equals("CUIL", StringComparison.OrdinalIgnoreCase))
                    return i;
            }
        }

        return -1;
    }

    private static (string CuitEmpleador, string Empleador, string Calle, string Localidad, string Provincia, string Telefono, string Contrato, string NroEstablecimiento, string Mail, string Referente) ExtractEmpresaData(DataTable table)
    {
        var cuitEmpleador = GetCellString(table, 7, 7);
        var empleador = GetCellString(table, 8, 7);
        var direccion = GetCellString(table, 10, 7);
        var infoContacto = GetCellString(table, 11, 7);
        var contrato = GetCellString(table, 7, 3);
        var nroEstablecimiento = GetCellString(table, 9, 7);

        var (mail, telefono, referente) = SplitContacto(infoContacto);
        var (calle, localidad, provincia) = SplitDireccion(direccion);

        localidad = NormalizeLocalidad(localidad);
        provincia = NormalizeProvincia(provincia);

        return (
            cuitEmpleador,
            empleador,
            calle,
            localidad,
            provincia,
            telefono,
            contrato,
            string.IsNullOrWhiteSpace(nroEstablecimiento) ? "0" : nroEstablecimiento,
            mail,
            referente
        );
    }

    private static (string Mail, string Telefono, string Referente) SplitContacto(string infoContacto)
    {
        if (string.IsNullOrWhiteSpace(infoContacto))
            return ("0", "0", "0");

        var partes = Regex.Split(infoContacto, "[;, ]+").Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
        var mails = new List<string>();
        var telefonos = new List<string>();
        var otros = new List<string>();

        foreach (var p in partes)
        {
            var item = p.Trim();
            if (item.Contains('@'))
            {
                mails.Add(item);
            }
            else if (Regex.IsMatch(item, "^[\\d+\\-\\(\\)\\s]+$"))
            {
                telefonos.Add(item);
            }
            else
            {
                otros.Add(item);
            }
        }

        var mail = mails.Count > 0 ? string.Join(",", mails) : "0";
        var telefono = telefonos.Count > 0 ? string.Join(",", telefonos) : "0";
        var referente = otros.Count > 0 ? string.Join(" ", otros) : "0";

        return (mail, telefono, referente);
    }

    private static (string Calle, string Localidad, string Provincia) SplitDireccion(string direccion)
    {
        if (string.IsNullOrWhiteSpace(direccion))
            return ("0", string.Empty, string.Empty);

        var partes = direccion.Split(',').Select(p => p.Trim()).Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
        var calle = partes.Count > 0 ? partes[0] : "0";
        var localidad = partes.Count > 1 ? partes[1] : string.Empty;
        var provincia = partes.Count > 2 ? partes[2] : string.Empty;

        return (calle, localidad, provincia);
    }

    private static string NormalizeLocalidad(string? localidad)
    {
        if (string.IsNullOrWhiteSpace(localidad))
            return string.Empty;

        var normalized = localidad.Trim();
        normalized = Regex.Replace(normalized, @"^\(\d+\)\s*", string.Empty);
        normalized = Regex.Replace(normalized, @"[-\s]+(B\s*A|BUENOS\s+AIRES)\s*$", string.Empty, RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized.ToUpperInvariant(), @"\s+", " ");
        return normalized.Trim();
    }

    private static string NormalizeProvincia(string? provincia)
    {
        if (string.IsNullOrWhiteSpace(provincia))
            return string.Empty;

        var normalized = Regex.Replace(provincia.Trim().ToUpperInvariant(), @"\s+", " ");

        return normalized switch
        {
            "BA" => "BUENOS AIRES",
            "B A" => "BUENOS AIRES",
            "BS AS" => "BUENOS AIRES",
            "BS. AS." => "BUENOS AIRES",
            "BSAS" => "BUENOS AIRES",
            _ => normalized
        };
    }

    private static string GetCellString(DataTable table, int rowIndex, int colIndex)
    {
        if (rowIndex < 0 || rowIndex >= table.Rows.Count)
            return string.Empty;

        return GetCellString(table.Rows[rowIndex], colIndex);
    }

    private static string GetCellString(DataRow row, int colIndex)
    {
        if (colIndex < 0 || colIndex >= row.ItemArray.Length)
            return string.Empty;

        var value = row[colIndex];
        if (value == null || value is DBNull)
            return string.Empty;

        return value.ToString()?.Trim() ?? string.Empty;
    }

    private static string GetCellNumericString(DataRow row, int colIndex)
    {
        if (colIndex < 0 || colIndex >= row.ItemArray.Length)
            return string.Empty;

        var value = row[colIndex];
        if (value == null || value is DBNull)
            return string.Empty;

        if (value is double d)
            return Convert.ToInt64(Math.Truncate(d)).ToString(CultureInfo.InvariantCulture);

        if (value is float f)
            return Convert.ToInt64(Math.Truncate(f)).ToString(CultureInfo.InvariantCulture);

        if (value is decimal dec)
            return Convert.ToInt64(Math.Truncate(dec)).ToString(CultureInfo.InvariantCulture);

        var text = value.ToString()?.Trim() ?? string.Empty;
        if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
        {
            return Convert.ToInt64(Math.Truncate(parsed)).ToString(CultureInfo.InvariantCulture);
        }

        return text;
    }
}
