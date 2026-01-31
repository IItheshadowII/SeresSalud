using ConvertidorDeOrdenes.Core.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ConvertidorDeOrdenes.Core.Services;

/// <summary>
/// Exportador a formato XLS (Excel 97-2003) usando NPOI HSSF
/// </summary>
public class XlsExporter
{
    /// <summary>
    /// Exporta filas a archivo XLS con 3 hojas
    /// </summary>
    public void Export(List<OutputRow> rows, string outputPath)
    {
        // Crear workbook HSSF (BIFF8 - XLS)
        var workbook = new HSSFWorkbook();

        // Crear 3 hojas
        var sheet1 = workbook.CreateSheet("Hoja1");
        var sheet2 = workbook.CreateSheet("Hoja2");
        var sheet3 = workbook.CreateSheet("Hoja3");

        // Escribir datos en Hoja1
        WriteDataToSheet(sheet1, rows);

        // Guardar archivo
        using var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
        workbook.Write(fileStream);
    }

    private void WriteDataToSheet(ISheet sheet, List<OutputRow> rows)
    {
        // Crear fila de encabezados
        var headerRow = sheet.CreateRow(0);
        var headers = new[]
        {
            "CuitEmpleador",        // A
            "CIIU",                 // B
            "Empleador",            // C
            "Calle",                // D
            "CodPostal",            // E
            "Localidad",            // F
            "Provincia",            // G
            "ABMlocProv",           // H
            "Telefono",             // I
            "Fax",                  // J
            "Contrato",             // K
            "NroEstablecimiento",   // L
            "Frecuencia",           // M
            "Cuil",                 // N
            "NroDocumento",         // O
            "TrabajadorApellidoNombre", // P
            "Riesgo",               // Q
            "DescripcionRiesgo",    // R
            "ABMRiesgo",            // S
            "Prestacion",           // T
            "HistoriaClinica",      // U
            "Mail",                 // V
            "Referente",            // W
            "DescripcionError",     // X
            "Id"                    // Y
        };

        for (int i = 0; i < headers.Length; i++)
        {
            headerRow.CreateCell(i).SetCellValue(headers[i]);
        }

        // Escribir filas de datos
        int rowIndex = 1;
        foreach (var dataRow in rows)
        {
            var row = sheet.CreateRow(rowIndex++);

            row.CreateCell(0).SetCellValue(dataRow.CuitEmpleador);           // A
            row.CreateCell(1).SetCellValue(dataRow.CIIU);                    // B
            row.CreateCell(2).SetCellValue(dataRow.Empleador);               // C
            row.CreateCell(3).SetCellValue(dataRow.Calle);                   // D
            row.CreateCell(4).SetCellValue(string.Empty);                    // E (vacío)
            row.CreateCell(5).SetCellValue(dataRow.Localidad);               // F
            row.CreateCell(6).SetCellValue(dataRow.Provincia);               // G
            row.CreateCell(7).SetCellValue(dataRow.ABMlocProv);              // H
            row.CreateCell(8).SetCellValue(dataRow.Telefono);                // I
            row.CreateCell(9).SetCellValue(dataRow.Fax);                     // J
            row.CreateCell(10).SetCellValue(dataRow.Contrato);               // K
            row.CreateCell(11).SetCellValue(dataRow.NroEstablecimiento);     // L
            row.CreateCell(12).SetCellValue(dataRow.Frecuencia);             // M
            row.CreateCell(13).SetCellValue(dataRow.Cuil);                   // N
            row.CreateCell(14).SetCellValue(string.Empty);                   // O (forzado)
            row.CreateCell(15).SetCellValue(dataRow.TrabajadorApellidoNombre); // P
            row.CreateCell(16).SetCellValue(dataRow.Riesgo);                 // Q
            row.CreateCell(17).SetCellValue(dataRow.DescripcionRiesgo);      // R
            row.CreateCell(18).SetCellValue(dataRow.ABMRiesgo);              // S
            row.CreateCell(19).SetCellValue(dataRow.Prestacion);             // T
            row.CreateCell(20).SetCellValue(dataRow.HistoriaClinica);        // U
            row.CreateCell(21).SetCellValue(dataRow.Mail);                   // V
            row.CreateCell(22).SetCellValue(string.Empty);                   // W (forzado)
            row.CreateCell(23).SetCellValue(dataRow.DescripcionError);       // X
            row.CreateCell(24).SetCellValue(dataRow.Id);                     // Y
        }

        // Auto-ajustar columnas principales
        for (int i = 0; i < 25; i++)
        {
            sheet.AutoSizeColumn(i);
        }
    }
}
