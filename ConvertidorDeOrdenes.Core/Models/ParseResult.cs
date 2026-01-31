namespace ConvertidorDeOrdenes.Core.Models;

/// <summary>
/// Resultado del parseo de un archivo de entrada
/// </summary>
public class ParseResult
{
    public List<OutputRow> Rows { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
    public List<string> Errors { get; set; } = new();
    public int TotalRows => Rows.Count;
    public int UniqueCompanies => Rows.Select(r => r.CuitEmpleador).Distinct().Count();
    public bool HasErrors => Errors.Count > 0;
}
