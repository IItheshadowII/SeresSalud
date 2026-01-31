namespace ConvertidorDeOrdenes.Core.Models;

/// <summary>
/// Representa una fila de salida en el archivo XLS destino (columnas A-Y)
/// </summary>
public class OutputRow
{
    // A - OBLIGATORIO
    public string CuitEmpleador { get; set; } = string.Empty;

    // B - Opcional
    public string CIIU { get; set; } = string.Empty;

    // C - OBLIGATORIO
    public string Empleador { get; set; } = string.Empty;

    // D - Opcional
    public string Calle { get; set; } = string.Empty;

    // E - Opcional
    public string CodPostal { get; set; } = string.Empty;

    // F - OBLIGATORIO
    public string Localidad { get; set; } = string.Empty;

    // G - OBLIGATORIO
    public string Provincia { get; set; } = string.Empty;

    // H - Opcional
    public string ABMlocProv { get; set; } = string.Empty;

    // I - Opcional
    public string Telefono { get; set; } = string.Empty;

    // J - Opcional
    public string Fax { get; set; } = string.Empty;

    // K - Opcional
    public string Contrato { get; set; } = string.Empty;

    // L - Opcional
    public string NroEstablecimiento { get; set; } = string.Empty;

    // M - OBLIGATORIO (A/S/R)
    public string Frecuencia { get; set; } = string.Empty;

    // N - OBLIGATORIO trabajador
    public string Cuil { get; set; } = string.Empty;

    // O - Opcional
    public string NroDocumento { get; set; } = string.Empty;

    // P - OBLIGATORIO
    public string TrabajadorApellidoNombre { get; set; } = string.Empty;

    // Q - OBLIGATORIO (MAX 90 caracteres)
    public string Riesgo { get; set; } = string.Empty;

    // R - Opcional
    public string DescripcionRiesgo { get; set; } = string.Empty;

    // S - Opcional
    public string ABMRiesgo { get; set; } = string.Empty;

    // T - OBLIGATORIO
    public string Prestacion { get; set; } = string.Empty;

    // U - Opcional
    public string HistoriaClinica { get; set; } = string.Empty;

    // V - Opcional
    public string Mail { get; set; } = string.Empty;

    // W - OBLIGATORIO
    public string Referente { get; set; } = string.Empty;

    // X - Opcional (mensajes de validación internos)
    public string DescripcionError { get; set; } = string.Empty;

    // Y - Opcional
    public string Id { get; set; } = string.Empty;
}
