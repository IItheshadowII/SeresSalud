using ConvertidorDeOrdenes.Core.Models;
using ConvertidorDeOrdenes.Core.Services;

namespace ConvertidorDeOrdenes.Desktop.Forms;

/// <summary>
/// Diálogo para editar o crear registro de empresa
/// </summary>
public partial class CompanyEditDialog : Form
{
    public CompanyRecord Company { get; private set; }

    private TextBox txtCUIT = null!;
    private TextBox txtCIIU = null!;
    private TextBox txtEmpleador = null!;
    private TextBox txtCalle = null!;
    private TextBox txtCodPostal = null!;
    private TextBox txtLocalidad = null!;
    private TextBox txtProvincia = null!;
    private TextBox txtTelefono = null!;
    private TextBox txtFax = null!;
    private TextBox txtMail = null!;
    private Button btnGuardar = null!;
    private Button btnCancelar = null!;

    private readonly bool _cuitRequired;

    public CompanyEditDialog(CompanyRecord? company = null, bool cuitRequired = true)
    {
        Company = company ?? new CompanyRecord();
        _cuitRequired = cuitRequired;
        
        InitializeComponent();
        LoadData();
    }

    private void InitializeComponent()
    {
        this.Text = "Datos de Empresa";
        this.Size = new Size(550, 560);
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.BackColor = Color.WhiteSmoke;

        var lblTitulo = new Label
        {
            Text = "Datos de Empresa",
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 120, 215),
            Location = new Point(25, 15),
            Size = new Size(480, 30)
        };
        this.Controls.Add(lblTitulo);

        int y = 60;
        int labelWidth = 130;
        int textBoxWidth = 340;
        int spacing = 42;

        // CUIT
        AddField("CUIT:", ref txtCUIT, y, labelWidth, textBoxWidth);
        y += spacing;

        // CIIU
        AddField("CIIU:", ref txtCIIU, y, labelWidth, textBoxWidth);
        y += spacing;

        // Empleador
        AddField("Empleador:", ref txtEmpleador, y, labelWidth, textBoxWidth);
        y += spacing;

        // Calle
        AddField("Calle:", ref txtCalle, y, labelWidth, textBoxWidth);
        y += spacing;

        // Código Postal
        AddField("Código Postal:", ref txtCodPostal, y, labelWidth, textBoxWidth);
        y += spacing;

        // Localidad
        AddField("Localidad:", ref txtLocalidad, y, labelWidth, textBoxWidth);
        y += spacing;

        // Provincia
        AddField("Provincia:", ref txtProvincia, y, labelWidth, textBoxWidth);
        y += spacing;

        // Teléfono
        AddField("Teléfono:", ref txtTelefono, y, labelWidth, textBoxWidth);
        y += spacing;

        // Fax
        AddField("Fax:", ref txtFax, y, labelWidth, textBoxWidth);
        y += spacing;

        // Mail
        AddField("Mail:", ref txtMail, y, labelWidth, textBoxWidth);
        y += spacing + 20;

        // Botones
        btnGuardar = new Button
        {
            Text = "Guardar",
            Location = new Point(240, y),
            Size = new Size(110, 38),
            DialogResult = DialogResult.OK,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnGuardar.FlatAppearance.BorderSize = 0;
        btnGuardar.Click += BtnGuardar_Click;

        btnCancelar = new Button
        {
            Text = "Cancelar",
            Location = new Point(360, y),
            Size = new Size(110, 38),
            DialogResult = DialogResult.Cancel,
            Font = new Font("Segoe UI", 10),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(100, 100, 100),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnCancelar.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);

        this.Controls.Add(btnGuardar);
        this.Controls.Add(btnCancelar);

        this.AcceptButton = btnGuardar;
        this.CancelButton = btnCancelar;
    }

    private void AddField(string labelText, ref TextBox textBox, int y, int labelWidth, int textBoxWidth)
    {
        var label = new Label
        {
            Text = labelText,
            Font = new Font("Segoe UI", 9, FontStyle.Regular),
            Location = new Point(25, y + 3),
            Size = new Size(labelWidth, 22)
        };

        textBox = new TextBox
        {
            Location = new Point(160, y),
            Size = new Size(textBoxWidth, 28),
            Font = new Font("Segoe UI", 9)
        };

        this.Controls.Add(label);
        this.Controls.Add(textBox);
    }

    private void LoadData()
    {
        txtCUIT.Text = Company.CUIT;
        txtCIIU.Text = Company.CIIU;
        txtEmpleador.Text = Company.Empleador;
        txtCalle.Text = Company.Calle;
        txtCodPostal.Text = Company.CodPostal;
        txtLocalidad.Text = Company.Localidad;
        txtProvincia.Text = Company.Provincia;
        txtTelefono.Text = Company.Telefono;
        txtFax.Text = Company.Fax;
        txtMail.Text = Company.Mail;
    }

    private void BtnGuardar_Click(object? sender, EventArgs e)
    {
        // Validar CUIT si es obligatorio
        if (_cuitRequired && string.IsNullOrWhiteSpace(txtCUIT.Text))
        {
            MessageBox.Show("El CUIT es obligatorio.", "Validación",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Guardar datos
        Company.CUIT = txtCUIT.Text.Trim();
        Company.CIIU = txtCIIU.Text.Trim();
        Company.Empleador = txtEmpleador.Text.Trim();
        Company.Calle = txtCalle.Text.Trim();
        Company.CodPostal = txtCodPostal.Text.Trim();
        Company.Localidad = txtLocalidad.Text.Trim();
        Company.Provincia = txtProvincia.Text.Trim();
        Company.Telefono = txtTelefono.Text.Trim();
        Company.Fax = txtFax.Text.Trim();
        Company.Mail = txtMail.Text.Trim();

        this.DialogResult = DialogResult.OK;
        this.Close();
    }
}

/// <summary>
/// Diálogo para seleccionar entre múltiples empresas
/// </summary>
public partial class CompanySelectDialog : Form
{
    public CompanyRecord? SelectedCompany { get; private set; }

    private ListBox lstCompanies = null!;
    private TextBox txtBuscar = null!;
    private Button btnSeleccionar = null!;
    private Button btnNuevo = null!;
    private Button btnCancelar = null!;

    private readonly List<CompanyRecord> _companies;          // lista completa
    private readonly List<CompanyRecord> _filteredCompanies;  // lista filtrada para mostrar
    private readonly CompanyRepositoryExcel _repository;

    public CompanySelectDialog(List<CompanyRecord> companies, CompanyRepositoryExcel repository)
    {
        _companies = companies;
        _filteredCompanies = new List<CompanyRecord>(companies);
        _repository = repository;
        InitializeComponent();
        LoadCompanies();
    }

    private void InitializeComponent()
    {
        this.Text = "Seleccionar Empresa";
        this.Size = new Size(680, 480);
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.BackColor = Color.WhiteSmoke;

        var lblInstrucciones = new Label
        {
            Text = "Se encontraron múltiples empresas. Seleccione una:",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 120, 215),
            Location = new Point(25, 20),
            Size = new Size(620, 24)
        };

        txtBuscar = new TextBox
        {
            PlaceholderText = "🔍 Buscar por CUIT, nombre o localidad...",
            Font = new Font("Segoe UI", 10),
            Location = new Point(25, 50),
            Size = new Size(620, 28)
        };
        txtBuscar.TextChanged += (_, _) => ApplyFilter();

        lstCompanies = new ListBox
        {
            Location = new Point(25, 90),
            Size = new Size(620, 280),
            Font = new Font("Segoe UI", 9),
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        lstCompanies.DoubleClick += (s, e) => BtnSeleccionar_Click(s, e);

        btnSeleccionar = new Button
        {
            Text = "Seleccionar",
            Location = new Point(135, 390),
            Size = new Size(110, 38),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnSeleccionar.FlatAppearance.BorderSize = 0;
        btnSeleccionar.Click += BtnSeleccionar_Click;

        var btnEliminar = new Button
        {
            Text = "Eliminar",
            Location = new Point(255, 390),
            Size = new Size(110, 38),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(200, 50, 50),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnEliminar.FlatAppearance.BorderSize = 0;
        btnEliminar.Click += BtnEliminar_Click;

        btnNuevo = new Button
        {
            Text = "Crear Nuevo",
            Location = new Point(375, 390),
            Size = new Size(110, 38),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(16, 124, 16),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnNuevo.FlatAppearance.BorderSize = 0;
        btnNuevo.Click += BtnNuevo_Click;

        btnCancelar = new Button
        {
            Text = "Cancelar",
            Location = new Point(535, 390),
            Size = new Size(110, 38),
            Font = new Font("Segoe UI", 9),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(100, 100, 100),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            DialogResult = DialogResult.Cancel
        };
        btnCancelar.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);

        this.Controls.Add(lblInstrucciones);
        this.Controls.Add(txtBuscar);
        this.Controls.Add(lstCompanies);
        this.Controls.Add(btnSeleccionar);
        this.Controls.Add(btnEliminar);
        this.Controls.Add(btnNuevo);
        this.Controls.Add(btnCancelar);

        this.CancelButton = btnCancelar;
    }

    private void LoadCompanies()
    {
        lstCompanies.Items.Clear();
        
        foreach (var company in _filteredCompanies)
        {
            var displayText = $"{company.CUIT} - {company.Empleador} - {company.Localidad}, {company.Provincia}";
            lstCompanies.Items.Add(displayText);
        }

        if (lstCompanies.Items.Count > 0)
        {
            lstCompanies.SelectedIndex = 0;
        }
    }

    private void BtnSeleccionar_Click(object? sender, EventArgs e)
    {
        if (lstCompanies.SelectedIndex >= 0)
        {
            SelectedCompany = _filteredCompanies[lstCompanies.SelectedIndex];
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        else
        {
            MessageBox.Show("Debe seleccionar una empresa.", "Validación",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void BtnNuevo_Click(object? sender, EventArgs e)
    {
        SelectedCompany = new CompanyRecord();
        this.DialogResult = DialogResult.Yes; // Indica que se creará nuevo
        this.Close();
    }

    private void ApplyFilter()
    {
        var term = (txtBuscar.Text ?? string.Empty).Trim().ToUpperInvariant();

        _filteredCompanies.Clear();

        if (string.IsNullOrWhiteSpace(term))
        {
            _filteredCompanies.AddRange(_companies);
        }
        else
        {
            foreach (var c in _companies)
            {
                var composite = $"{c.CUIT} {c.Empleador} {c.Localidad} {c.Provincia}".ToUpperInvariant();
                if (composite.Contains(term))
                    _filteredCompanies.Add(c);
            }
        }

        LoadCompanies();
    }

    private void BtnEliminar_Click(object? sender, EventArgs e)
    {
        if (lstCompanies.SelectedIndex < 0)
        {
            MessageBox.Show("Debe seleccionar una empresa para eliminar.", "Validación",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var company = _filteredCompanies[lstCompanies.SelectedIndex];

        var confirm = MessageBox.Show(
            $"¿Confirma que desea eliminar la empresa:\n\n{company.CUIT} - {company.Empleador}?\n\nSe realizará un backup previo de Empresas.xlsx.",
            "Confirmar eliminación",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (confirm != DialogResult.Yes)
            return;

        try
        {
            _repository.DeleteCompany(company);

            _companies.Remove(company);
            _filteredCompanies.Remove(company);
            LoadCompanies();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error eliminando empresa: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
