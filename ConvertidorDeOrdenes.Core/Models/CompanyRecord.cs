namespace ConvertidorDeOrdenes.Core.Models;

/// <summary>
/// Representa un registro de empresa en la base Empresas.xlsx
/// </summary>
public class CompanyRecord
{
    public string CUIT { get; set; } = string.Empty;
    public string CIIU { get; set; } = string.Empty;
    public string Empleador { get; set; } = string.Empty;
    public string Calle { get; set; } = string.Empty;
    public string CodPostal { get; set; } = string.Empty;
    public string Localidad { get; set; } = string.Empty;
    public string Provincia { get; set; } = string.Empty;
    public string Telefono { get; set; } = string.Empty;
    public string Fax { get; set; } = string.Empty;
    public string Mail { get; set; } = string.Empty;

    /// <summary>
    /// Identificador temporal para uso interno (índice de fila)
    /// </summary>
    public int RowIndex { get; set; }
}
