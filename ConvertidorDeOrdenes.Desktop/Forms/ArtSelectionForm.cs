namespace ConvertidorDeOrdenes.Desktop.Forms;

/// <summary>
/// Primera ventana: selección de ART. De esta elección depende el flujo de la app.
/// </summary>
public sealed class ArtSelectionForm : Form
{
    public string ArtSeleccionada { get; private set; } = string.Empty;

    private ComboBox _cbArt = null!;
    private Button _btnSiguiente = null!;
    private Button _btnCancelar = null!;

    private bool _allowClose;

    public ArtSelectionForm()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "ConvertidorDeOrdenes - Selección de ART";
        Size = new Size(500, 280);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ControlBox = false;
        FormClosing += ArtSelectionForm_FormClosing;
        BackColor = Color.WhiteSmoke;

        var lblTitulo = new Label
        {
            Text = "🏥 Seleccione la ART",
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50),
            Location = new Point(30, 25),
            Size = new Size(440, 35)
        };

        var lblDescripcion = new Label
        {
            Text = "Elija la Aseguradora de Riesgos del Trabajo para configurar el procesamiento",
            Font = new Font("Segoe UI", 9, FontStyle.Regular),
            ForeColor = Color.FromArgb(90, 90, 90),
            Location = new Point(30, 62),
            Size = new Size(440, 18)
        };

        var lblArt = new Label
        {
            Text = "ART (obligatorio):",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Location = new Point(30, 95),
            Size = new Size(440, 22)
        };

        _cbArt = new ComboBox
        {
            Location = new Point(30, 120),
            Size = new Size(340, 28),
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 10)
        };

        // Para futuro: agregar más ARTs acá.
        _cbArt.Items.Add("La Segunda");
        _cbArt.SelectedIndex = 0;

        _btnSiguiente = new Button
        {
            Text = "✓ Siguiente",
            Location = new Point(250, 170),
            Size = new Size(110, 40),
            DialogResult = DialogResult.OK,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(70, 130, 180),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnSiguiente.FlatAppearance.BorderSize = 0;
        _btnSiguiente.Click += BtnSiguiente_Click;

        _btnCancelar = new Button
        {
            Text = "✕ Cancelar",
            Location = new Point(370, 170),
            Size = new Size(100, 40),
            DialogResult = DialogResult.Cancel,
            Font = new Font("Segoe UI", 10),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(100, 100, 100),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnCancelar.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        _btnCancelar.Click += (_, _) => { _allowClose = true; Close(); };

        Controls.Add(lblTitulo);
        Controls.Add(lblDescripcion);
        Controls.Add(lblArt);
        Controls.Add(_cbArt);
        Controls.Add(_btnSiguiente);
        Controls.Add(_btnCancelar);

        AcceptButton = _btnSiguiente;
        CancelButton = _btnCancelar;
    }

    private void BtnSiguiente_Click(object? sender, EventArgs e)
    {
        if (_cbArt.SelectedIndex < 0)
        {
            MessageBox.Show("Debe seleccionar una ART.", "Validación",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        ArtSeleccionada = _cbArt.SelectedItem?.ToString() ?? string.Empty;

        _allowClose = true;
        DialogResult = DialogResult.OK;
        Close();
    }

    private void ArtSelectionForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (!_allowClose)
        {
            e.Cancel = true;
        }
    }
}
