using System.Globalization;
using System.Text;

namespace ConvertidorDeOrdenes.Core.Services;

/// <summary>
/// Mapeador de prestaciones desde archivo de conversión
/// </summary>
public class PrestacionMapper
{
    private readonly Dictionary<string, string> _mappings = new(StringComparer.OrdinalIgnoreCase);
    private readonly string _baseDirectory;

    public PrestacionMapper(string baseDirectory)
    {
        _baseDirectory = baseDirectory;
        LoadMappings();
    }

    /// <summary>
    /// Mapea una prestación origen a destino
    /// </summary>
    public string Map(string prestacionOrigen)
    {
        if (string.IsNullOrWhiteSpace(prestacionOrigen))
            return string.Empty;

        if (_mappings.TryGetValue(prestacionOrigen, out var mapped))
        {
            return mapped;
        }

        // No hay mapeo, devolver original
        return prestacionOrigen;
    }

    private void LoadMappings()
    {
        try
        {
            // Intentar cargar PrestacionesMap.csv
            var csvPath = Path.Combine(_baseDirectory, "PrestacionesMap.csv");
            if (File.Exists(csvPath))
            {
                LoadFromCsv(csvPath);
                return;
            }

            // Intentar cargar PrestacionesMap.xlsx
            var xlsxPath = Path.Combine(_baseDirectory, "PrestacionesMap.xlsx");
            if (File.Exists(xlsxPath))
            {
                LoadFromXlsx(xlsxPath);
                return;
            }

            // No hay archivo de mapeo
        }
        catch (Exception)
        {
            // Silenciar errores de carga, los mapeos quedan vacíos
        }
    }

    private void LoadFromCsv(string filePath)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        
        using var reader = new StreamReader(filePath, Encoding.GetEncoding("iso-8859-1"));
        
        // Saltar encabezado
        reader.ReadLine();

        while (!reader.EndOfStream)
        {
            var line = reader.ReadLine();
            if (string.IsNullOrWhiteSpace(line))
                continue;

            var parts = line.Split(',');
            if (parts.Length >= 2)
            {
                var origen = parts[0].Trim().Trim('"');
                var destino = parts[1].Trim().Trim('"');

                if (!string.IsNullOrWhiteSpace(origen) && !string.IsNullOrWhiteSpace(destino))
                {
                    _mappings[origen] = destino;
                }
            }
        }
    }

    private void LoadFromXlsx(string filePath)
    {
        try
        {
            using var workbook = new ClosedXML.Excel.XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.First();

            // Saltar encabezado (fila 1)
            var rows = worksheet.RowsUsed().Skip(1);

            foreach (var row in rows)
            {
                var origen = row.Cell(1).GetString().Trim();
                var destino = row.Cell(2).GetString().Trim();

                if (!string.IsNullOrWhiteSpace(origen) && !string.IsNullOrWhiteSpace(destino))
                {
                    _mappings[origen] = destino;
                }
            }
        }
        catch
        {
            // Error al leer XLSX, ignorar
        }
    }
}
