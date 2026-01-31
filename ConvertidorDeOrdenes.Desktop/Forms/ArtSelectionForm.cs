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
        Size = new Size(420, 220);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ControlBox = false;
        FormClosing += ArtSelectionForm_FormClosing;

        var lblTitulo = new Label
        {
            Text = "Seleccione la ART",
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            Location = new Point(20, 20),
            Size = new Size(360, 30)
        };

        var lblArt = new Label
        {
            Text = "ART (obligatorio):",
            Location = new Point(20, 70),
            Size = new Size(360, 20)
        };

        _cbArt = new ComboBox
        {
            Location = new Point(20, 95),
            Size = new Size(260, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };

        // Para futuro: agregar más ARTs acá.
        _cbArt.Items.Add("La Segunda");
        _cbArt.SelectedIndex = 0;

        _btnSiguiente = new Button
        {
            Text = "Siguiente",
            Location = new Point(200, 140),
            Size = new Size(90, 30),
            DialogResult = DialogResult.OK
        };
        _btnSiguiente.Click += BtnSiguiente_Click;

        _btnCancelar = new Button
        {
            Text = "Cancelar",
            Location = new Point(300, 140),
            Size = new Size(90, 30),
            DialogResult = DialogResult.Cancel
        };
        _btnCancelar.Click += (_, _) => { _allowClose = true; Close(); };

        Controls.Add(lblTitulo);
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
