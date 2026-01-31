using ConvertidorDeOrdenes.Core.Models;

namespace ConvertidorDeOrdenes.Core.Services;

/// <summary>
/// Validador de filas de salida
/// </summary>
public class Validator
{
    /// <summary>
    /// Valida una fila de salida
    /// </summary>
    public ValidationResult Validate(OutputRow row)
    {
        var result = new ValidationResult { IsValid = true };

        // Validar campos OBLIGATORIOS
        if (string.IsNullOrWhiteSpace(row.CuitEmpleador))
        {
            result.Errors.Add("CUIT Empleador es obligatorio");
            result.IsValid = false;
        }

        if (string.IsNullOrWhiteSpace(row.Empleador))
        {
            result.Errors.Add("Empleador es obligatorio");
            result.IsValid = false;
        }

        if (string.IsNullOrWhiteSpace(row.Localidad))
        {
            result.Errors.Add("Localidad es obligatoria");
            result.IsValid = false;
        }

        if (string.IsNullOrWhiteSpace(row.Provincia))
        {
            result.Errors.Add("Provincia es obligatoria");
            result.IsValid = false;
        }

        if (string.IsNullOrWhiteSpace(row.Frecuencia))
        {
            result.Errors.Add("Frecuencia es obligatoria");
            result.IsValid = false;
        }
        else if (!new[] { "A", "S", "R" }.Contains(row.Frecuencia.ToUpper()))
        {
            result.Errors.Add($"Frecuencia inválida: '{row.Frecuencia}'. Debe ser A, S o R");
            result.IsValid = false;
        }

        if (string.IsNullOrWhiteSpace(row.Cuil))
        {
            result.Errors.Add("CUIL Trabajador es obligatorio");
            result.IsValid = false;
        }

        if (string.IsNullOrWhiteSpace(row.TrabajadorApellidoNombre))
        {
            result.Errors.Add("Apellido y Nombre del Trabajador es obligatorio");
            result.IsValid = false;
        }

        if (string.IsNullOrWhiteSpace(row.Riesgo))
        {
            result.Errors.Add("Riesgo es obligatorio");
            result.IsValid = false;
        }

        if (string.IsNullOrWhiteSpace(row.Prestacion))
        {
            result.Errors.Add("Prestación es obligatoria");
            result.IsValid = false;
        }

        // Validaciones de formato (warnings)
        if (!string.IsNullOrWhiteSpace(row.CuitEmpleador))
        {
            if (!IsValidCuitFormat(row.CuitEmpleador))
            {
                result.Warnings.Add($"Formato de CUIT posiblemente inválido: {row.CuitEmpleador}");
            }
        }

        if (!string.IsNullOrWhiteSpace(row.Cuil))
        {
            if (!IsValidCuitFormat(row.Cuil))
            {
                result.Warnings.Add($"Formato de CUIL posiblemente inválido: {row.Cuil}");
            }
        }

        return result;
    }

    /// <summary>
    /// Valida formato básico de CUIT/CUIL (XX-XXXXXXXX-X)
    /// </summary>
    private bool IsValidCuitFormat(string cuit)
    {
        if (string.IsNullOrWhiteSpace(cuit))
            return false;

        // Formato: XX-XXXXXXXX-X o solo dígitos (11 dígitos)
        var digitsOnly = new string(cuit.Where(char.IsDigit).ToArray());
        
        return digitsOnly.Length == 11;
    }
}
