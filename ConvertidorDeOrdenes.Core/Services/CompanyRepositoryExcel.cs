using ConvertidorDeOrdenes.Core.Models;
using ClosedXML.Excel;

namespace ConvertidorDeOrdenes.Core.Services;

/// <summary>
/// Repositorio de empresas basado en archivo Excel
/// </summary>
public class CompanyRepositoryExcel
{
    private readonly string _filePath;
    private List<CompanyRecord> _companies = new();

    public string FilePath => _filePath;
    public int CompaniesCount => _companies.Count;

    public CompanyRepositoryExcel(string baseDirectory)
    {
        _filePath = ResolveCompaniesFilePath(baseDirectory);
        LoadCompanies();
    }

    /// <summary>
    /// Busca empresas por CUIT
    /// </summary>
    public List<CompanyRecord> SearchByCuit(string cuit)
    {
        if (string.IsNullOrWhiteSpace(cuit))
            return new List<CompanyRecord>();

        var cuitClean = new string(cuit.Where(char.IsDigit).ToArray());

        return _companies.Where(c =>
        {
            var companyCuit = new string(c.CUIT.Where(char.IsDigit).ToArray());
            return companyCuit.Equals(cuitClean, StringComparison.OrdinalIgnoreCase);
        }).ToList();
    }

    /// <summary>
    /// Busca empresas por nombre (contiene)
    /// </summary>
    public List<CompanyRecord> SearchByName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return new List<CompanyRecord>();

        var searchNorm = NormalizeName(name);

        return _companies.Where(c =>
        {
            var companyNorm = NormalizeName(c.Empleador);
            return companyNorm.Contains(searchNorm) || searchNorm.Contains(companyNorm);
        }).ToList();
    }

    /// <summary>
    /// Busca empresas por CUIT o nombre
    /// </summary>
    public List<CompanyRecord> Search(string cuitOrName)
    {
        var byCuit = SearchByCuit(cuitOrName);
        if (byCuit.Any())
            return byCuit;

        return SearchByName(cuitOrName);
    }

    /// <summary>
    /// Agrega o actualiza una empresa
    /// </summary>
    public void SaveCompany(CompanyRecord company)
    {
        // Buscar si existe por CUIT exacto
        var existing = _companies.FirstOrDefault(c =>
        {
            var cuitClean = new string(c.CUIT.Where(char.IsDigit).ToArray());
            var companyCuit = new string(company.CUIT.Where(char.IsDigit).ToArray());
            return cuitClean.Equals(companyCuit, StringComparison.OrdinalIgnoreCase) &&
                   c.Localidad.Equals(company.Localidad, StringComparison.OrdinalIgnoreCase);
        });

        if (existing != null)
        {
            // Actualizar
            existing.CIIU = company.CIIU;
            existing.Empleador = company.Empleador;
            existing.Calle = company.Calle;
            existing.CodPostal = company.CodPostal;
            existing.Localidad = company.Localidad;
            existing.Provincia = company.Provincia;
            existing.Telefono = company.Telefono;
            existing.Fax = company.Fax;
            existing.Mail = company.Mail;
        }
        else
        {
            // Agregar nuevo
            company.RowIndex = _companies.Count > 0 ? _companies.Max(c => c.RowIndex) + 1 : 2;
            _companies.Add(company);
        }

        // Guardar a archivo
        SaveToFile();
    }

    /// <summary>
    /// Obtiene empresa por índice
    /// </summary>
    public CompanyRecord? GetByIndex(int index)
    {
        return _companies.FirstOrDefault(c => c.RowIndex == index);
    }

    /// <summary>
    /// Devuelve una copia de la lista completa de empresas.
    /// </summary>
    public List<CompanyRecord> GetAll()
    {
        return _companies
            .Select(c => new CompanyRecord
            {
                RowIndex = c.RowIndex,
                CUIT = c.CUIT,
                CIIU = c.CIIU,
                Empleador = c.Empleador,
                Calle = c.Calle,
                CodPostal = c.CodPostal,
                Localidad = c.Localidad,
                Provincia = c.Provincia,
                Telefono = c.Telefono,
                Fax = c.Fax,
                Mail = c.Mail
            })
            .ToList();
    }

    /// <summary>
    /// Elimina una empresa de la base y guarda el archivo, realizando un backup previo.
    /// </summary>
    public void DeleteCompany(CompanyRecord company)
    {
        // Buscar por RowIndex si está disponible
        CompanyRecord? existing = null;

        if (company.RowIndex > 0)
        {
            existing = _companies.FirstOrDefault(c => c.RowIndex == company.RowIndex);
        }

        // Fallback por combinación de CUIT + Empleador + Localidad
        if (existing == null)
        {
            var cuitClean = new string((company.CUIT ?? string.Empty).Where(char.IsDigit).ToArray());

            existing = _companies.FirstOrDefault(c =>
            {
                var ec = new string((c.CUIT ?? string.Empty).Where(char.IsDigit).ToArray());
                return ec.Equals(cuitClean, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(c.Empleador, company.Empleador, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(c.Localidad, company.Localidad, StringComparison.OrdinalIgnoreCase);
            });
        }

        if (existing == null)
        {
            return; // Nada que borrar
        }

        // Backup antes de modificar el archivo físico
        try
        {
            if (File.Exists(_filePath))
            {
                var dir = Path.GetDirectoryName(_filePath) ?? string.Empty;
                var name = Path.GetFileNameWithoutExtension(_filePath);
                var ext = Path.GetExtension(_filePath);
                var backupName = $"{name}_backup_{DateTime.Now:yyyyMMdd_HHmmss}{ext}";
                var backupPath = Path.Combine(dir, backupName);

                File.Copy(_filePath, backupPath, overwrite: false);
            }
        }
        catch
        {
            // Si falla el backup, no impedir la operación, pero idealmente se loguearía.
        }

        _companies.Remove(existing);
        SaveToFile();
    }

    private void LoadCompanies()
    {
        _companies.Clear();

        if (!File.Exists(_filePath))
        {
            // Crear archivo vacío
            CreateEmptyFile();
            return;
        }

        try
        {
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = FindCompaniesWorksheet(workbook);
            // Detectar formato de Empresas.xlsx
            // Formato 1 (DB externa):
            // A: Razon Social, B: Cuit, C: Calle, D: Localidad, E: Provincia,
            // F: Telefono, G: Email, H: Referente, I: Registro
            // Formato 2 (generado por la app):
            // 1: CUIT, 2: CIIU, 3: Empleador, 4: Calle, 5: CodPostal,
            // 6: Localidad, 7: Provincia, 8: Telefono, 9: Fax, 10: Mail

            var headerA = worksheet.Cell(1, 1).GetString().Trim();
            var headerB = worksheet.Cell(1, 2).GetString().Trim();

            int colCuit, colRazon, colCiiu, colCalle, colCodPostal, colLocalidad, colProvincia, colTelefono, colFax, colMail;

            if (headerA.Contains("RAZON", StringComparison.OrdinalIgnoreCase) ||
                headerA.Contains("RAZÓN", StringComparison.OrdinalIgnoreCase))
            {
                // Formato DB externa descrito por el usuario
                colRazon = 1;  // Razon Social
                colCuit = 2;   // Cuit
                colCalle = 3;  // Calle
                colLocalidad = 4; // Localidad
                colProvincia = 5; // Provincia
                colTelefono = 6;  // Telefono
                colMail = 7;      // Email

                colCiiu = 0;      // No existe
                colCodPostal = 0; // No existe
                colFax = 0;       // No existe
            }
            else
            {
                // Asumir formato estándar de la app (encabezado CUIT en col 1)
                colCuit = 1;
                colCiiu = 2;
                colRazon = 3;
                colCalle = 4;
                colCodPostal = 5;
                colLocalidad = 6;
                colProvincia = 7;
                colTelefono = 8;
                colFax = 9;
                colMail = 10;
            }

            // Leer desde fila 2 (saltando encabezado)
            var rows = worksheet.RowsUsed().Skip(1);

            int rowIndex = 2;
            foreach (var row in rows)
            {
                string Get(int col) => col > 0 ? row.Cell(col).GetString().Trim() : string.Empty;

                var company = new CompanyRecord
                {
                    RowIndex = rowIndex,
                    CUIT = Get(colCuit),
                    CIIU = Get(colCiiu),
                    Empleador = Get(colRazon),
                    Calle = Get(colCalle),
                    CodPostal = Get(colCodPostal),
                    Localidad = Get(colLocalidad),
                    Provincia = Get(colProvincia),
                    Telefono = Get(colTelefono),
                    Fax = Get(colFax),
                    Mail = Get(colMail)
                };

                _companies.Add(company);
                rowIndex++;
            }
        }
        catch (Exception ex)
        {
            // No sobrescribir un archivo real si hubo error de lectura
            throw new Exception($"Error leyendo Empresas.xlsx en '{_filePath}': {ex.Message}", ex);
        }
    }

    private static IXLWorksheet FindCompaniesWorksheet(XLWorkbook workbook)
    {
        // Priorizar una hoja que tenga encabezados típicos en la fila 1
        foreach (var ws in workbook.Worksheets)
        {
            var a1 = ws.Cell(1, 1).GetString().Trim();
            var b1 = ws.Cell(1, 2).GetString().Trim();

            if (a1.Contains("RAZON", StringComparison.OrdinalIgnoreCase) ||
                a1.Contains("RAZÓN", StringComparison.OrdinalIgnoreCase) ||
                a1.Contains("EMPLEADOR", StringComparison.OrdinalIgnoreCase) ||
                a1.Contains("CUIT", StringComparison.OrdinalIgnoreCase) ||
                b1.Contains("CUIT", StringComparison.OrdinalIgnoreCase))
            {
                return ws;
            }
        }

        return workbook.Worksheets.First();
    }

    /// <summary>
    /// Resuelve la ruta real de Empresas.xlsx buscando hacia arriba desde baseDirectory
    /// </summary>
    private static string ResolveCompaniesFilePath(string baseDirectory)
    {
        try
        {
            // Buscar todas las copias de Empresas.xlsx hacia arriba desde baseDirectory
            // y elegir la más probable (por tamaño de archivo). Esto evita quedarse con
            // una copia vacía en bin/Debug si existe una base real en una carpeta superior.
            var candidates = new List<string>();

            var dir = new DirectoryInfo(baseDirectory);
            while (dir != null)
            {
                var candidate = Path.Combine(dir.FullName, "Empresas.xlsx");
                if (File.Exists(candidate))
                    candidates.Add(candidate);

                dir = dir.Parent;
            }

            if (candidates.Count > 0)
            {
                string best = candidates[0];
                long bestLen = new FileInfo(best).Length;
                foreach (var c in candidates)
                {
                    var len = new FileInfo(c).Length;
                    if (len > bestLen)
                    {
                        best = c;
                        bestLen = len;
                    }
                }

                return best;
            }
        }
        catch
        {
            // Ignorar y caer al fallback
        }

        // Fallback: usar baseDirectory aunque no exista todavía
        return Path.Combine(baseDirectory, "Empresas.xlsx");
    }

    private static string NormalizeName(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;

        text = text.Trim().ToUpperInvariant();

        // Normalizar acentos
        var normalized = text.Normalize(System.Text.NormalizationForm.FormD);
        var sb = new System.Text.StringBuilder();
        foreach (var c in normalized)
        {
            var uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
            if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                sb.Append(c);
        }

        // Quitar caracteres no alfanuméricos
        var only = new System.Text.StringBuilder();
        foreach (var c in sb.ToString())
        {
            if (char.IsLetterOrDigit(c) || char.IsWhiteSpace(c))
                only.Append(c);
        }

        // Colapsar espacios múltiples
        var collapsed = new System.Text.StringBuilder();
        bool lastWasSpace = false;
        foreach (var c in only.ToString().Trim())
        {
            var isSpace = char.IsWhiteSpace(c);
            if (isSpace)
            {
                if (!lastWasSpace)
                    collapsed.Append(' ');
                lastWasSpace = true;
            }
            else
            {
                collapsed.Append(c);
                lastWasSpace = false;
            }
        }

        return collapsed.ToString().Trim();
    }

    private void SaveToFile()
    {
        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Empresas");

            // Encabezados
            worksheet.Cell(1, 1).Value = "CUIT";
            worksheet.Cell(1, 2).Value = "CIIU";
            worksheet.Cell(1, 3).Value = "Empleador";
            worksheet.Cell(1, 4).Value = "Calle";
            worksheet.Cell(1, 5).Value = "CodPostal";
            worksheet.Cell(1, 6).Value = "Localidad";
            worksheet.Cell(1, 7).Value = "Provincia";
            worksheet.Cell(1, 8).Value = "Telefono";
            worksheet.Cell(1, 9).Value = "Fax";
            worksheet.Cell(1, 10).Value = "Mail";

            // Datos
            int row = 2;
            foreach (var company in _companies)
            {
                worksheet.Cell(row, 1).Value = company.CUIT;
                worksheet.Cell(row, 2).Value = company.CIIU;
                worksheet.Cell(row, 3).Value = company.Empleador;
                worksheet.Cell(row, 4).Value = company.Calle;
                worksheet.Cell(row, 5).Value = company.CodPostal;
                worksheet.Cell(row, 6).Value = company.Localidad;
                worksheet.Cell(row, 7).Value = company.Provincia;
                worksheet.Cell(row, 8).Value = company.Telefono;
                worksheet.Cell(row, 9).Value = company.Fax;
                worksheet.Cell(row, 10).Value = company.Mail;
                row++;
            }

            workbook.SaveAs(_filePath);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error guardando Empresas.xlsx: {ex.Message}", ex);
        }
    }

    private void CreateEmptyFile()
    {
        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Empresas");

            // Solo encabezados
            worksheet.Cell(1, 1).Value = "CUIT";
            worksheet.Cell(1, 2).Value = "CIIU";
            worksheet.Cell(1, 3).Value = "Empleador";
            worksheet.Cell(1, 4).Value = "Calle";
            worksheet.Cell(1, 5).Value = "CodPostal";
            worksheet.Cell(1, 6).Value = "Localidad";
            worksheet.Cell(1, 7).Value = "Provincia";
            worksheet.Cell(1, 8).Value = "Telefono";
            worksheet.Cell(1, 9).Value = "Fax";
            worksheet.Cell(1, 10).Value = "Mail";

            workbook.SaveAs(_filePath);
        }
        catch
        {
            // Si no se puede crear, continuar sin archivo
        }
    }
}
