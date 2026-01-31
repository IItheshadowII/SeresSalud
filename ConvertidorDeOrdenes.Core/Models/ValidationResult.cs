namespace ConvertidorDeOrdenes.Core.Models;

/// <summary>
/// Resultado de la validación de filas
/// </summary>
public class ValidationResult
{
    public bool IsValid { get; set; }
    public List<string> Errors { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
}
