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
        this.Size = new Size(500, 360);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.ControlBox = false;
        this.FormClosing += WizardForm_FormClosing;

        // Título
        lblTitulo = new Label
        {
            Text = "Configuración de Nueva Conversión",
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            Location = new Point(20, 20),
            Size = new Size(450, 30)
        };

        // Tipo de carga
        lblTipoCarga = new Label
        {
            Text = "Seleccione el tipo de carga:",
            Location = new Point(20, 70),
            Size = new Size(450, 20)
        };

        rbAnualesSemestrales = new RadioButton
        {
            Text = "Anuales/Semestrales (XLSX con múltiples solapas)",
            Location = new Point(40, 95),
            Size = new Size(400, 25),
            Checked = true
        };
        rbAnualesSemestrales.CheckedChanged += Radio_CheckedChanged;

        rbReconfirmatorios = new RadioButton
        {
            Text = "Reconfirmatorios/Reevaluaciones (CSV)",
            Location = new Point(40, 125),
            Size = new Size(400, 25)
        };
        rbReconfirmatorios.CheckedChanged += Radio_CheckedChanged;

        // Frecuencia
        lblFrecuencia = new Label
        {
            Text = "Frecuencia (obligatorio):",
            Location = new Point(20, 165),
            Size = new Size(450, 20)
        };

        cbFrecuencia = new ComboBox
        {
            Location = new Point(40, 190),
            Size = new Size(200, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        PopulateFrecuencias();

        // Botones
        btnSiguiente = new Button
        {
            Text = "Siguiente",
            Location = new Point(260, 280),
            Size = new Size(90, 30),
            DialogResult = DialogResult.OK
        };
        btnSiguiente.Click += BtnSiguiente_Click;

        btnCancelar = new Button
        {
            Text = "Cancelar",
            Location = new Point(360, 280),
            Size = new Size(90, 30),
            DialogResult = DialogResult.Cancel
        };
        btnCancelar.Click += (_, _) => { _allowClose = true; Close(); };

        // Agregar controles
        this.Controls.Add(lblTitulo);
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
