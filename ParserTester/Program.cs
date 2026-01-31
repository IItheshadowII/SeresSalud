using ConvertidorDeOrdenes.Core.Parsers;
using System.Text;

Console.OutputEncoding = Encoding.UTF8;

var parser = new CsvOrderParser();
var file = Path.GetFullPath("..\\revaluaciones_30_06012026.csv");
if (!File.Exists(file))
{
	Console.WriteLine($"Archivo no encontrado: {file}");
	return;
}

var result = parser.Parse(file, "R", "TEST");
Console.WriteLine($"Total filas: {result.TotalRows}");
Console.WriteLine($"Warnings: {result.Warnings.Count}");
Console.WriteLine($"Errors: {result.Errors.Count}");

for (int i = 0; i < Math.Min(5, result.Rows.Count); i++)
{
	var r = result.Rows[i];
	Console.WriteLine($"Fila {i+1}: CUIT={r.CuitEmpleador} Empleador={r.Empleador} Prestacion={r.Prestacion} Localidad={r.Localidad} Provincia={r.Provincia}");
}
