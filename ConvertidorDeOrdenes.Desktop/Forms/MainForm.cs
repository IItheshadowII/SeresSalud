using ConvertidorDeOrdenes.Core.Models;
using ConvertidorDeOrdenes.Core.Parsers;
using ConvertidorDeOrdenes.Core.Services;
using ConvertidorDeOrdenes.Desktop.Services;
using ConvertidorDeOrdenes.Desktop.Services.Updates;
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
    private readonly string _installDirectory;
    private readonly string _dataDirectory;

    private ParseResult? _parseResult;
    private CompanyRepositoryExcel? _companyRepository;
    private Normalizer? _normalizer;
    private Validator? _validator;
    private Logger? _logger;

    private UpdateService? _updateService;

    private Task? _initializationTask;
    private Exception? _initializationException;

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
        _installDirectory = AppPaths.InstallDirectory;
        _dataDirectory = AppPaths.DataRootDirectory;

        InitializeComponent();
        AppIcon.Apply(this);

        // Cargar servicios pesados (Empresas.xlsx, mapeos, etc.) en segundo plano
        _initializationTask = Task.Run(InitializeServicesBackground);
    }

    private void InitializeComponent()
    {
        var version = AppPaths.GetCurrentVersion();
        this.Text = $"ConvertidorDeOrdenes v{version} - Procesamiento";
        this.Icon = AppIcon.TryGet();
        this.Size = new Size(1250, 750);
        this.MinimumSize = new Size(1250, 750);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.WhiteSmoke;

        // Menú principal (Empresas)
        var menuStrip = new MenuStrip();
        menuStrip.BackColor = Color.FromArgb(70, 130, 180);
        menuStrip.ForeColor = Color.White;
        menuStrip.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        
        var empresasMenu = new ToolStripMenuItem("Empresas");
        empresasMenu.ForeColor = Color.White;
        
        var administrarEmpresasItem = new ToolStripMenuItem("Administrar...");
        administrarEmpresasItem.Click += async (_, _) =>
        {
            try
            {
                await EnsureInitializedAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inicializando la base de empresas: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (_companyRepository == null)
            {
                MessageBox.Show("El repositorio de empresas no está disponible todavía. Espere unos segundos y vuelva a intentar.", "Empresas",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var form = new CompanyListForm(_companyRepository);
            form.ShowDialog(this);
        };
        empresasMenu.DropDownItems.Add(administrarEmpresasItem);
        menuStrip.Items.Add(empresasMenu);

        var ayudaMenu = new ToolStripMenuItem("Ayuda") { ForeColor = Color.White };
        var buscarActualizacionesItem = new ToolStripMenuItem("Buscar actualizaciones...");
        buscarActualizacionesItem.Click += async (_, _) =>
        {
            using var dlg = new UpdateSettingsDialog();
            var result = dlg.ShowDialog(this);

            if (result == DialogResult.OK && dlg.StartUpdateCheck)
            {
                await CheckForUpdatesAsync(interactive: true);
            }
        };
        ayudaMenu.DropDownItems.Add(buscarActualizacionesItem);
        menuStrip.Items.Add(ayudaMenu);

        this.MainMenuStrip = menuStrip;
        this.Controls.Add(menuStrip);

        this.Shown += async (_, _) => await CheckForUpdatesAsync(interactive: false);

        // Título y subtítulo
        var lblTitulo = new Label
        {
            Text = "📂 Procesamiento de Órdenes Médicas",
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50),
            Location = new Point(25, 40),
            Size = new Size(600, 30)
        };

        var lblSubtitulo = new Label
        {
            Text = "Analice archivos de órdenes, verifique los datos y exporte al formato requerido",
            Font = new Font("Segoe UI", 9, FontStyle.Regular),
            ForeColor = Color.FromArgb(90, 90, 90),
            Location = new Point(25, 72),
            Size = new Size(900, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        // Archivo
        lblArchivo = new Label
        {
            Text = "📄 Archivo de entrada:",
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = Color.FromArgb(60, 60, 60),
            Location = new Point(25, 108),
            Size = new Size(170, 24)
        };

        txtArchivo = new TextBox
        {
            Location = new Point(200, 105),
            Size = new Size(745, 28),
            ReadOnly = true,
            Font = new Font("Segoe UI", 9),
            BackColor = Color.White,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        btnSeleccionar = new Button
        {
            Text = "📁 Elegir archivo...",
            Location = new Point(955, 103),
            Size = new Size(140, 35),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(70, 130, 180),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        btnSeleccionar.FlatAppearance.BorderSize = 0;
        btnSeleccionar.Click += BtnSeleccionar_Click;

        btnAnalizar = new Button
        {
            Text = "▶ Analizar",
            Location = new Point(1105, 103),
            Size = new Size(100, 35),
            Enabled = false,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(60, 140, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        btnAnalizar.FlatAppearance.BorderSize = 0;
        btnAnalizar.Click += BtnAnalizar_Click;

        // DataGridView
        dgvPreview = new DataGridView
        {
            Location = new Point(25, 155),
            Size = new Size(1180, 365),
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            ColumnHeadersHeight = 35,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(70, 130, 180),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Padding = new Padding(5),
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                WrapMode = DataGridViewTriState.False
            },
            AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(245, 245, 245)
            }
        };

        // Separador visual
        var separator1 = new Label
        {
            Location = new Point(25, 528),
            Size = new Size(1180, 1),
            BackColor = Color.FromArgb(200, 200, 200),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };

        // Estadísticas
        lblEstadisticas = new Label
        {
            Location = new Point(25, 535),
            Size = new Size(1180, 26),
            Text = "ℹ️ Seleccione un archivo y pulse 'Analizar' para comenzar el procesamiento",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(80, 80, 80),
            TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };

        // Título del log
        var lblLog = new Label
        {
            Text = "📋 Registro de actividad:",
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = Color.FromArgb(80, 80, 80),
            Location = new Point(25, 570),
            Size = new Size(200, 20),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        };

        // Log
        txtLog = new TextBox
        {
            Location = new Point(25, 595),
            Size = new Size(1180, 55),
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            ReadOnly = true,
            Font = new Font("Consolas", 8),
            BackColor = Color.FromArgb(250, 250, 250),
            BorderStyle = BorderStyle.FixedSingle,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };

        // Botones
        btnExportar = new Button
        {
            Text = "💾 Exportar XLS",
            Location = new Point(1060, 660),
            Size = new Size(145, 40),
            Enabled = false,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(60, 140, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        btnExportar.FlatAppearance.BorderSize = 0;
        btnExportar.Click += BtnExportar_Click;

        var tooltipExportar = new ToolTip();
        tooltipExportar.SetToolTip(btnExportar, "Exportar los datos procesados a formato XLS (Excel 97-2003)");

        btnNuevo = new Button
        {
            Text = "🔄 Nueva Conversión",
            Location = new Point(895, 660),
            Size = new Size(155, 40),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(100, 100, 100),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        btnNuevo.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        btnNuevo.Click += BtnNuevo_Click;

        var tooltipNuevo = new ToolTip();
        tooltipNuevo.SetToolTip(btnNuevo, "Reiniciar el proceso con un nuevo archivo");

        this.Controls.Add(lblTitulo);
        this.Controls.Add(lblSubtitulo);
        this.Controls.Add(lblArchivo);
        this.Controls.Add(txtArchivo);
        this.Controls.Add(btnSeleccionar);
        this.Controls.Add(btnAnalizar);
        this.Controls.Add(dgvPreview);
        this.Controls.Add(separator1);
        this.Controls.Add(lblEstadisticas);
        this.Controls.Add(lblLog);
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

    private void InitializeServicesBackground()
    {
        try
        {
            var logger = new Logger(_dataDirectory);
            logger.LogInfo("=== Inicio de sesión ===");
            logger.LogInfo($"Tipo de carga: {_tipoCarga}");
            logger.LogInfo($"Frecuencia: {_frecuencia}");
            logger.LogInfo($"ART: {_art}");
            logger.LogInfo($"Referente: {_referente}");
            logger.LogInfo($"InstallDir: {_installDirectory}");
            logger.LogInfo($"DataDir: {_dataDirectory}");

            var companyRepository = new CompanyRepositoryExcel(_dataDirectory);
            logger.LogInfo($"Empresas.xlsx: {companyRepository.FilePath}");
            logger.LogInfo($"Empresas cargadas: {companyRepository.CompaniesCount}");
            var prestacionMapper = new PrestacionMapper(_installDirectory);
            var normalizer = new Normalizer(prestacionMapper);
            var validator = new Validator();

            var updateService = new UpdateService(new HttpClient(), new UpdateStateStore(AppPaths.UpdateStatePath));

            // Publicar instancias listas
            _logger = logger;
            _companyRepository = companyRepository;
            _normalizer = normalizer;
            _validator = validator;
            _updateService = updateService;
        }
        catch (Exception ex)
        {
            _initializationException = ex;
        }
    }

    private async Task EnsureInitializedAsync()
    {
        var task = _initializationTask;
        if (task == null)
            return;

        try
        {
            await task.ConfigureAwait(true);
        }
        catch
        {
            // La excepción real se guarda en _initializationException
        }

        if (_initializationException != null)
        {
            throw _initializationException;
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

    private async Task CheckForUpdatesAsync(bool interactive)
    {
        if (_updateService == null)
            return;

        try
        {
            if (!interactive && !_updateService.ShouldAutoCheck())
                return;

            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
            var update = await _updateService.CheckLatestAsync(cts.Token);

            if (update == null)
            {
                if (interactive)
                {
                    MessageBox.Show("No hay actualizaciones disponibles.", "Actualizaciones",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return;
            }

            if (!interactive && !_updateService.ShouldNotify(update))
                return;

            _updateService.MarkNotified(update);

            var result = MessageBox.Show(
                $"Hay una nueva versión disponible: v{update.Version}.\n\n¿Desea descargar e instalar ahora?\n\nLa aplicación se cerrará para completar la actualización.",
                "Actualización disponible",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);

            if (result != DialogResult.Yes)
                return;

            UseWaitCursor = true;
            Enabled = false;

            var installerPath = await _updateService.DownloadInstallerAsync(update, progress: null, cts.Token);
            if (string.IsNullOrWhiteSpace(installerPath) || !File.Exists(installerPath))
            {
                MessageBox.Show(
                    "No se pudo descargar el instalador.\n\nPuede descargarlo manualmente desde la página de Releases.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (!_updateService.RunInstallerAndExit(installerPath))
            {
                MessageBox.Show(
                    "No se pudo ejecutar el instalador automáticamente.\n\nPuede ejecutarlo manualmente desde: " + installerPath,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            Application.Exit();
        }
        catch (Exception ex)
        {
            _logger?.LogError($"Error durante la verificación/descarga de actualización: {ex}");

            if (interactive)
            {
                MessageBox.Show(
                    "No se pudo completar la actualización.\n\nDetalle técnico: " + ex.Message,
                    "Actualizaciones",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
        finally
        {
            UseWaitCursor = false;
            Enabled = true;
        }
    }

    private async void BtnAnalizar_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtArchivo.Text))
            return;

        try
        {
            await EnsureInitializedAsync();

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
                row.CuitEmpleador, row.CIIU, row.Empleador, row.Calle, string.Empty,
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

        var statusIcon = errors.Count > 0 ? "❌" : warnings.Count > 0 ? "⚠️" : "✅";
        var stats = $"{statusIcon} Total filas: {_parseResult.TotalRows} | " +
                   $"🏢 Empresas: {_parseResult.UniqueCompanies} | " +
                   $"👥 Empleados únicos: {uniqueEmployees} | " +
                   $"⚠️ Warnings: {warnings.Count} | " +
                   $"❌ Errores: {errors.Count}";

        lblEstadisticas.Text = stats;
        lblEstadisticas.ForeColor = errors.Count > 0 ? Color.FromArgb(180, 60, 60) : 
                                    warnings.Count > 0 ? Color.FromArgb(180, 120, 0) : 
                                    Color.FromArgb(60, 140, 60);

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
