using ConvertidorDeOrdenes.Core.Models;
using ConvertidorDeOrdenes.Core.Parsers;
using ConvertidorDeOrdenes.Core.Services;
using System.Text.RegularExpressions;

namespace ConvertidorDeOrdenes.Desktop.Forms;

/// <summary>
/// Formulario principal de selección de archivo y preview
/// </summary>
public partial class MainForm : Form
{
    private readonly WizardForm.TipoCarga _tipoCarga;
    private readonly string _frecuencia;
    private readonly string _referente;
    private readonly string _art;
    private readonly string _baseDirectory;

    private ParseResult? _parseResult;
    private CompanyRepositoryExcel? _companyRepository;
    private Normalizer? _normalizer;
    private Validator? _validator;
    private Logger? _logger;

    private Label lblArchivo = null!;
    private TextBox txtArchivo = null!;
    private Button btnSeleccionar = null!;
    private Button btnAnalizar = null!;
    private DataGridView dgvPreview = null!;
    private Label lblEstadisticas = null!;
    private Button btnExportar = null!;
    private Button btnNuevo = null!;
    private TextBox txtLog = null!;

    public MainForm(WizardForm.TipoCarga tipoCarga, string frecuencia, string referente, string art)
    {
        _tipoCarga = tipoCarga;
        _frecuencia = frecuencia;
        _referente = referente;
        _art = art;
        _baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

        InitializeComponent();
        InitializeServices();
    }

    private void InitializeComponent()
    {
        this.Text = "ConvertidorDeOrdenes - Procesamiento";
        this.Size = new Size(1250, 750);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.WhiteSmoke;

        // Menú principal (Empresas)
        var menuStrip = new MenuStrip();
        menuStrip.BackColor = Color.FromArgb(0, 120, 215);
        menuStrip.ForeColor = Color.White;
        menuStrip.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        
        var empresasMenu = new ToolStripMenuItem("Empresas");
        empresasMenu.ForeColor = Color.White;
        
        var administrarEmpresasItem = new ToolStripMenuItem("Administrar...");
        administrarEmpresasItem.Click += (_, _) =>
        {
            if (_companyRepository == null)
            {
                MessageBox.Show("El repositorio de empresas no está disponible.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var form = new CompanyListForm(_companyRepository);
            form.ShowDialog(this);
        };
        empresasMenu.DropDownItems.Add(administrarEmpresasItem);
        menuStrip.Items.Add(empresasMenu);

        this.MainMenuStrip = menuStrip;
        this.Controls.Add(menuStrip);

        // Archivo
        lblArchivo = new Label
        {
            Text = "Archivo de entrada:",
            Font = new Font("Segoe UI", 10, FontStyle.Regular),
            Location = new Point(25, 45),
            Size = new Size(140, 24)
        };

        txtArchivo = new TextBox
        {
            Location = new Point(165, 42),
            Size = new Size(780, 28),
            ReadOnly = true,
            Font = new Font("Segoe UI", 9),
            BackColor = Color.White
        };

        btnSeleccionar = new Button
        {
            Text = "Elegir archivo...",
            Location = new Point(955, 40),
            Size = new Size(130, 35),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnSeleccionar.FlatAppearance.BorderSize = 0;
        btnSeleccionar.Click += BtnSeleccionar_Click;

        btnAnalizar = new Button
        {
            Text = "Analizar",
            Location = new Point(1095, 40),
            Size = new Size(110, 35),
            Enabled = false,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(16, 124, 16),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnAnalizar.FlatAppearance.BorderSize = 0;
        btnAnalizar.Click += BtnAnalizar_Click;

        // DataGridView
        dgvPreview = new DataGridView
        {
            Location = new Point(25, 90),
            Size = new Size(1180, 430),
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Padding = new Padding(5)
            },
            AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(245, 245, 245)
            }
        };

        // Estadísticas
        lblEstadisticas = new Label
        {
            Location = new Point(25, 535),
            Size = new Size(1180, 24),
            Text = "Seleccione un archivo y pulse 'Analizar'",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(80, 80, 80)
        };

        // Log
        txtLog = new TextBox
        {
            Location = new Point(25, 570),
            Size = new Size(1180, 70),
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            ReadOnly = true,
            Font = new Font("Consolas", 8),
            BackColor = Color.FromArgb(240, 240, 240),
            BorderStyle = BorderStyle.FixedSingle
        };

        // Botones
        btnExportar = new Button
        {
            Text = "Exportar XLS",
            Location = new Point(1060, 650),
            Size = new Size(145, 40),
            Enabled = false,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(16, 124, 16),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnExportar.FlatAppearance.BorderSize = 0;
        btnExportar.Click += BtnExportar_Click;

        btnNuevo = new Button
        {
            Text = "Nueva Conversión",
            Location = new Point(900, 650),
            Size = new Size(150, 40),
            Font = new Font("Segoe UI", 9),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(100, 100, 100),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnNuevo.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        btnNuevo.Click += BtnNuevo_Click;

        this.Controls.Add(lblArchivo);
        this.Controls.Add(txtArchivo);
        this.Controls.Add(btnSeleccionar);
        this.Controls.Add(btnAnalizar);
        this.Controls.Add(dgvPreview);
        this.Controls.Add(lblEstadisticas);
        this.Controls.Add(txtLog);
        this.Controls.Add(btnExportar);
        this.Controls.Add(btnNuevo);

        ConfigureDataGridView();
    }

    private void ConfigureDataGridView()
    {
        // Configurar columnas A-X
        var columns = new[]
        {
            "A-CuitEmpleador", "B-CIIU", "C-Empleador", "D-Calle", "E-CodPostal",
            "F-Localidad", "G-Provincia", "H-ABMlocProv", "I-Telefono", "J-Fax",
            "K-Contrato", "L-NroEstablecimiento", "M-Frecuencia", "N-Cuil", "O-NroDocumento",
            "P-TrabajadorApellidoNombre", "Q-Riesgo", "R-DescripcionRiesgo", "S-ABMRiesgo",
            "T-Prestacion", "U-HistoriaClinica", "V-Mail", "W-Referente", "X-DescripcionError", "Y-Id"
        };

        foreach (var col in columns)
        {
            dgvPreview.Columns.Add(col, col);
            dgvPreview.Columns[dgvPreview.Columns.Count - 1].Width = 100;
        }
    }

    private void InitializeServices()
    {
        try
        {
            _logger = new Logger(_baseDirectory);
            _logger.LogInfo("=== Inicio de sesión ===");
            _logger.LogInfo($"Tipo de carga: {_tipoCarga}");
            _logger.LogInfo($"Frecuencia: {_frecuencia}");
            _logger.LogInfo($"ART: {_art}");
            _logger.LogInfo($"Referente: {_referente}");

            _companyRepository = new CompanyRepositoryExcel(_baseDirectory);
            _logger.LogInfo($"Empresas.xlsx: {_companyRepository.FilePath}");
            _logger.LogInfo($"Empresas cargadas: {_companyRepository.CompaniesCount}");
            var prestacionMapper = new PrestacionMapper(_baseDirectory);
            _normalizer = new Normalizer(prestacionMapper);
            _validator = new Validator();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error inicializando servicios: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BtnSeleccionar_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog();
        
        if (_tipoCarga == WizardForm.TipoCarga.AnualesSemestrales)
        {
            dialog.Filter = "Archivos Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
        }
        else
        {
            dialog.Filter = "Archivos CSV (*.csv)|*.csv|Todos los archivos (*.*)|*.*";
        }

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtArchivo.Text = dialog.FileName;
            btnAnalizar.Enabled = true;
            LogMessage($"Archivo seleccionado: {dialog.FileName}");
        }
    }

    private void BtnAnalizar_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtArchivo.Text))
            return;

        try
        {
            LogMessage("Iniciando análisis...");
            _logger?.LogInfo($"Analizando archivo: {txtArchivo.Text}");

            // Parsear archivo
            if (_tipoCarga == WizardForm.TipoCarga.ReconfirmatoriosReevaluaciones)
            {
                var parser = new CsvOrderParser();
                _parseResult = parser.Parse(txtArchivo.Text, _frecuencia, _referente);
            }
            else
            {
                var parser = new XlsxOrderParser();
                _parseResult = parser.Parse(txtArchivo.Text, _frecuencia, _referente);
            }

            LogMessage($"Parseado: {_parseResult.TotalRows} filas");
            _logger?.LogInfo($"Filas parseadas: {_parseResult.TotalRows}");

            if (_parseResult.TotalRows == 0)
            {
                MessageBox.Show("El archivo no contiene filas válidas para procesar.", "Información",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Completar datos de empresa automáticamente usando la base, sin mostrar diálogos
            AutoResolveCompanies();

            // Abrir ventana de revisión/edición de datos cargados
            if (_companyRepository != null)
            {
                using var resolutionForm = new CompanyResolutionForm(_parseResult.Rows, _companyRepository);
                resolutionForm.ShowDialog(this);
            }

            // Normalizar y validar luego de que el usuario complete datos
            NormalizeAndValidate(out var warnings, out var errors);

            // Mostrar en grid final
            LoadDataIntoGrid();

            // Actualizar estadísticas y habilitar exportación
            UpdateStatistics(warnings, errors);
            btnExportar.Enabled = errors.Count == 0;
            
            LogMessage("Análisis completado.");
        }
        catch (Exception ex)
        {
            LogMessage($"ERROR: {ex.Message}");
            _logger?.LogError($"Error en análisis: {ex.Message}");
            MessageBox.Show($"Error analizando archivo: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void AutoResolveCompanies()
    {
        if (_parseResult == null || _companyRepository == null)
            return;

        foreach (var row in _parseResult.Rows)
        {
            // Extraer CP desde localidad si viene como "(6034) LOCALIDAD-B A" para que
            // el modal de revisión muestre ya el Código Postal.
            ExtractCodPostalFromLocalidad(row);

            if (!string.IsNullOrWhiteSpace(row.CuitEmpleador))
            {
                var companies = _companyRepository.SearchByCuit(row.CuitEmpleador);

                if (companies.Count == 1)
                {
                    var company = companies[0];
                    // Cuando hay match único por CUIT, la DB es la fuente de verdad.
                    row.CuitEmpleador = company.CUIT;
                    row.CIIU = company.CIIU;
                    row.Empleador = company.Empleador;
                    row.Calle = company.Calle;
                    if (!string.IsNullOrWhiteSpace(company.CodPostal))
                        row.CodPostal = company.CodPostal;
                    if (!string.IsNullOrWhiteSpace(company.Localidad))
                        row.Localidad = company.Localidad;
                    if (!string.IsNullOrWhiteSpace(company.Provincia))
                        row.Provincia = company.Provincia;
                    if (!string.IsNullOrWhiteSpace(company.Telefono))
                        row.Telefono = company.Telefono;
                    if (!string.IsNullOrWhiteSpace(company.Fax))
                        row.Fax = company.Fax;
                    if (!string.IsNullOrWhiteSpace(company.Mail))
                        row.Mail = company.Mail;
                }
            }
            else if (!string.IsNullOrWhiteSpace(row.Empleador))
            {
                // Intentar resolver por nombre de empresa cuando no tenemos CUIT
                var companies = _companyRepository.SearchByName(row.Empleador);

                if (companies.Count == 1)
                {
                    var company = companies[0];
                    // Cuando matchea por nombre de empresa, también dejamos que DB gane.
                    row.CuitEmpleador = company.CUIT;
                    row.CIIU = company.CIIU;
                    row.Empleador = company.Empleador;
                    row.Calle = company.Calle;
                    if (!string.IsNullOrWhiteSpace(company.CodPostal))
                        row.CodPostal = company.CodPostal;
                    if (!string.IsNullOrWhiteSpace(company.Localidad))
                        row.Localidad = company.Localidad;
                    if (!string.IsNullOrWhiteSpace(company.Provincia))
                        row.Provincia = company.Provincia;
                    if (!string.IsNullOrWhiteSpace(company.Telefono))
                        row.Telefono = company.Telefono;
                    if (!string.IsNullOrWhiteSpace(company.Fax))
                        row.Fax = company.Fax;
                    if (!string.IsNullOrWhiteSpace(company.Mail))
                        row.Mail = company.Mail;
                }
            }
        }
    }

    private static void ExtractCodPostalFromLocalidad(OutputRow row)
    {
        if (!string.IsNullOrWhiteSpace(row.CodPostal))
            return;

        if (string.IsNullOrWhiteSpace(row.Localidad))
            return;

        var match = Regex.Match(row.Localidad, "^\\((\\d{3,5})\\)\\s*(.+)$");
        if (!match.Success)
            return;

        row.CodPostal = match.Groups[1].Value.Trim();
        row.Localidad = match.Groups[2].Value.Trim();
    }

    private void NormalizeAndValidate(out List<string> warnings, out List<string> errors)
    {
        warnings = new List<string>();
        errors = new List<string>();

        if (_parseResult == null)
            return;

        foreach (var row in _parseResult.Rows)
        {
            _normalizer?.NormalizeRow(row, warnings);
            var validation = _validator?.Validate(row);

            if (validation != null)
            {
                errors.AddRange(validation.Errors);
                warnings.AddRange(validation.Warnings);

                if (!validation.IsValid)
                {
                    row.DescripcionError = string.Join("; ", validation.Errors);
                }
                else
                {
                    row.DescripcionError = string.Empty;
                }
            }
        }
    }

    private void LoadDataIntoGrid()
    {
        if (_parseResult == null)
            return;

        dgvPreview.Rows.Clear();

        foreach (var row in _parseResult.Rows)
        {
            dgvPreview.Rows.Add(
                row.CuitEmpleador, row.CIIU, row.Empleador, row.Calle, row.CodPostal,
                row.Localidad, row.Provincia, row.ABMlocProv, row.Telefono, row.Fax,
                row.Contrato, row.NroEstablecimiento, row.Frecuencia, row.Cuil, row.NroDocumento,
                row.TrabajadorApellidoNombre, row.Riesgo, row.DescripcionRiesgo, row.ABMRiesgo,
                row.Prestacion, row.HistoriaClinica, row.Mail, row.Referente, row.DescripcionError,
                row.Id
            );

            // Colorear filas con error
            if (!string.IsNullOrWhiteSpace(row.DescripcionError))
            {
                dgvPreview.Rows[dgvPreview.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightCoral;
            }
        }
    }

    private void UpdateStatistics(List<string> warnings, List<string> errors)
    {
        if (_parseResult == null)
            return;

        var uniqueEmployees = _parseResult.Rows
            .Select(r => new string((r.Cuil ?? string.Empty).Where(char.IsDigit).ToArray()))
            .Where(cuilDigits => !string.IsNullOrWhiteSpace(cuilDigits))
            .Distinct()
            .Count();

        var stats = $"Total filas: {_parseResult.TotalRows} | " +
                   $"Empresas: {_parseResult.UniqueCompanies} | " +
                   $"Empleados únicos: {uniqueEmployees} | " +
                   $"Warnings: {warnings.Count} | " +
                   $"Errores: {errors.Count}";

        lblEstadisticas.Text = stats;
        lblEstadisticas.ForeColor = errors.Count > 0 ? Color.Red : Color.Green;

        foreach (var warning in warnings)
        {
            _logger?.LogWarning(warning);
        }

        foreach (var error in errors)
        {
            _logger?.LogError(error);
        }
    }

    private void BtnExportar_Click(object? sender, EventArgs e)
    {
        if (_parseResult == null || _parseResult.Rows.Count == 0)
        {
            MessageBox.Show("No hay datos para exportar.", "Advertencia",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var dialog = new SaveFileDialog
        {
            Filter = "Archivos XLS (*.xls)|*.xls",
            FileName = $"SALIDA_{DateTime.Now:yyyyMMdd_HHmmss}.xls"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                var exporter = new XlsExporter();
                exporter.Export(_parseResult.Rows, dialog.FileName);

                _logger?.LogInfo($"Archivo exportado: {dialog.FileName}");
                LogMessage($"Exportado exitosamente: {dialog.FileName}");

                MessageBox.Show($"Archivo exportado exitosamente:\n{dialog.FileName}",
                    "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR al exportar: {ex.Message}");
                _logger?.LogError($"Error exportando: {ex.Message}");
                MessageBox.Show($"Error exportando archivo: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void BtnNuevo_Click(object? sender, EventArgs e)
    {
        var result = MessageBox.Show("¿Desea iniciar una nueva conversión?", "Confirmar",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            this.Close();
            Application.Restart();
        }
    }

    private void LogMessage(string message)
    {
        txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}" + Environment.NewLine);
    }
}
