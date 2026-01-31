namespace ConvertidorDeOrdenes.Desktop.Forms;

/// <summary>
/// Formulario de bienvenida y configuración inicial (Wizard)
/// </summary>
public partial class WizardForm : Form
{
    public enum TipoCarga
    {
        AnualesSemestrales,
        ReconfirmatoriosReevaluaciones
    }

    public TipoCarga TipoCargaSeleccionado { get; private set; }
    public string FrecuenciaSeleccionada { get; private set; } = string.Empty;
    public string Referente { get; private set; } = string.Empty;

    private bool _allowClose;

    private RadioButton rbAnualesSemestrales = null!;
    private RadioButton rbReconfirmatorios = null!;
    private ComboBox cbFrecuencia = null!;
    private Button btnSiguiente = null!;
    private Button btnCancelar = null!;
    private Label lblTitulo = null!;
    private Label lblTipoCarga = null!;
    private Label lblFrecuencia = null!;

    public WizardForm()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        this.Text = "ConvertidorDeOrdenes - Configuración Inicial";
        this.Size = new Size(580, 420);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.ControlBox = false;
        this.FormClosing += WizardForm_FormClosing;
        this.BackColor = Color.WhiteSmoke;

        // Título
        lblTitulo = new Label
        {
            Text = "⚙️ Configuración de Nueva Conversión",
            Font = new Font("Segoe UI", 16, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 120, 215),
            Location = new Point(30, 25),
            Size = new Size(520, 35)
        };

        var lblDescripcion = new Label
        {
            Text = "Configure los parámetros de procesamiento según el tipo de archivo que desea convertir",
            Font = new Font("Segoe UI", 9, FontStyle.Regular),
            ForeColor = Color.FromArgb(100, 100, 100),
            Location = new Point(30, 62),
            Size = new Size(520, 18)
        };

        // Tipo de carga
        lblTipoCarga = new Label
        {
            Text = "Seleccione el tipo de carga:",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Location = new Point(30, 95),
            Size = new Size(520, 24)
        };

        rbAnualesSemestrales = new RadioButton
        {
            Text = "📊 Anuales/Semestrales (XLSX con múltiples solapas)",
            Font = new Font("Segoe UI", 10),
            Location = new Point(50, 125),
            Size = new Size(480, 28),
            Checked = true
        };
        rbAnualesSemestrales.CheckedChanged += Radio_CheckedChanged;

        rbReconfirmatorios = new RadioButton
        {
            Text = "📄 Reconfirmatorios/Reevaluaciones (CSV)",
            Font = new Font("Segoe UI", 10),
            Location = new Point(50, 160),
            Size = new Size(480, 28)
        };
        rbReconfirmatorios.CheckedChanged += Radio_CheckedChanged;

        // Frecuencia
        lblFrecuencia = new Label
        {
            Text = "Frecuencia (obligatorio):",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Location = new Point(30, 210),
            Size = new Size(520, 24)
        };

        cbFrecuencia = new ComboBox
        {
            Location = new Point(50, 240),
            Size = new Size(280, 28),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 10)
        };
        PopulateFrecuencias();

        // Botones
        btnSiguiente = new Button
        {
            Text = "Siguiente",
            Location = new Point(330, 320),
            Size = new Size(110, 40),
            DialogResult = DialogResult.OK,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnSiguiente.FlatAppearance.BorderSize = 0;
        btnSiguiente.Click += BtnSiguiente_Click;

        btnCancelar = new Button
        {
            Text = "Cancelar",
            Location = new Point(450, 320),
            Size = new Size(100, 40),
            DialogResult = DialogResult.Cancel,
            Font = new Font("Segoe UI", 10),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(100, 100, 100),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnCancelar.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        btnCancelar.Click += (_, _) => { _allowClose = true; Close(); };

        // Agregar controles
        this.Controls.Add(lblTitulo);
        this.Controls.Add(lblDescripcion);
        this.Controls.Add(lblTipoCarga);
        this.Controls.Add(rbAnualesSemestrales);
        this.Controls.Add(rbReconfirmatorios);
        this.Controls.Add(lblFrecuencia);
        this.Controls.Add(cbFrecuencia);
        this.Controls.Add(btnSiguiente);
        this.Controls.Add(btnCancelar);

        this.AcceptButton = btnSiguiente;
        this.CancelButton = btnCancelar;
    }

    private void BtnSiguiente_Click(object? sender, EventArgs e)
    {
        // Validar frecuencia
        if (cbFrecuencia.SelectedIndex < 0)
        {
            MessageBox.Show("Debe seleccionar una frecuencia.", "Validación",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Guardar selecciones
        TipoCargaSeleccionado = rbAnualesSemestrales.Checked
            ? TipoCarga.AnualesSemestrales
            : TipoCarga.ReconfirmatoriosReevaluaciones;

        var frecuenciaText = cbFrecuencia.SelectedItem?.ToString() ?? string.Empty;
        FrecuenciaSeleccionada = frecuenciaText.Split(' ')[0]; // Extraer "A", "S" o "R"

        // Requisito actual: no pedir Referente al inicio y dejarlo vacío
        Referente = string.Empty;

        _allowClose = true;
        this.DialogResult = DialogResult.OK;
        this.Close();
    }

    private void WizardForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (!_allowClose)
        {
            e.Cancel = true;
        }
    }

    private void Radio_CheckedChanged(object? sender, EventArgs e)
    {
        PopulateFrecuencias();
    }

    private void PopulateFrecuencias()
    {
        if (cbFrecuencia == null) return;

        cbFrecuencia.Items.Clear();

        if (rbReconfirmatorios.Checked)
        {
            cbFrecuencia.Items.Add("R - Reconfirmatorio");
            cbFrecuencia.SelectedIndex = 0;
            cbFrecuencia.Enabled = false;
        }
        else
        {
            cbFrecuencia.Items.Add("A - Anual");
            cbFrecuencia.Items.Add("S - Semestral");
            cbFrecuencia.Enabled = true;
            cbFrecuencia.SelectedIndex = 0;
        }
    }
}
