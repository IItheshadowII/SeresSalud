using System.Drawing;
using System.Windows.Forms;
using ConvertidorDeOrdenes.Desktop.Services;

namespace ConvertidorDeOrdenes.Desktop.Forms;

/// <summary>
/// Diálogo nativo para recoger credenciales del portal de La Segunda.
/// Las credenciales se ingresan aquí y se inyectan en el formulario de login
/// del portal vía WebView2, sin instalar hooks en el DOM externo.
/// </summary>
internal sealed class PortalCredentialDialog : Form
{
    private readonly TextBox _txtUsername = new();
    private readonly TextBox _txtPassword = new();
    private readonly CheckBox _chkRemember = new();
    private readonly Button _btnOk = new();
    private readonly Button _btnCancel = new();

    public string Username => _txtUsername.Text.Trim();
    public string Password => _txtPassword.Text;
    public bool RememberCredentials => _chkRemember.Checked;

    public PortalCredentialDialog(string prefilledUsername = "")
    {
        InitializeComponent();
        _txtUsername.Text = prefilledUsername;
        ActiveControl = string.IsNullOrEmpty(prefilledUsername) ? _txtUsername : _txtPassword;
    }

    private void InitializeComponent()
    {
        Text = "Credenciales del portal - La Segunda";
        Icon = AppIcon.TryGet();
        Size = new Size(370, 240);
        MinimumSize = new Size(370, 240);
        MaximumSize = new Size(370, 240);
        MaximizeBox = false;
        MinimizeBox = false;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(16, 12, 16, 12),
            RowCount = 5,
            ColumnCount = 2,
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var lblInfo = new Label
        {
            Text = "Ingresá tus credenciales del portal La Segunda:",
            AutoSize = false,
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 9),
            TextAlign = ContentAlignment.MiddleLeft,
        };
        layout.SetColumnSpan(lblInfo, 2);
        layout.Controls.Add(lblInfo, 0, 0);

        var lblUser = new Label
        {
            Text = "Usuario:",
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new Font("Segoe UI", 9),
        };
        _txtUsername.Dock = DockStyle.Fill;
        _txtUsername.Font = new Font("Segoe UI", 9);
        layout.Controls.Add(lblUser, 0, 1);
        layout.Controls.Add(_txtUsername, 1, 1);

        var lblPass = new Label
        {
            Text = "Contraseña:",
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new Font("Segoe UI", 9),
        };
        _txtPassword.Dock = DockStyle.Fill;
        _txtPassword.Font = new Font("Segoe UI", 9);
        _txtPassword.UseSystemPasswordChar = true;
        layout.Controls.Add(lblPass, 0, 2);
        layout.Controls.Add(_txtPassword, 1, 2);

        _chkRemember.Text = "Recordar credenciales";
        _chkRemember.Checked = true;
        _chkRemember.Dock = DockStyle.Fill;
        _chkRemember.Font = new Font("Segoe UI", 9);
        layout.SetColumnSpan(_chkRemember, 2);
        layout.Controls.Add(_chkRemember, 0, 3);

        var btnPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(0, 4, 0, 0),
        };
        _btnCancel.Text = "Cancelar";
        _btnCancel.Size = new Size(80, 28);
        _btnCancel.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };

        _btnOk.Text = "Aceptar";
        _btnOk.Size = new Size(80, 28);
        _btnOk.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        _btnOk.Click += BtnOk_Click;

        btnPanel.Controls.Add(_btnCancel);
        btnPanel.Controls.Add(_btnOk);
        layout.SetColumnSpan(btnPanel, 2);
        layout.Controls.Add(btnPanel, 0, 4);

        Controls.Add(layout);
        AcceptButton = _btnOk;
        CancelButton = _btnCancel;
    }

    private void BtnOk_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(Username))
        {
            MessageBox.Show("Ingresá el usuario.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            _txtUsername.Focus();
            return;
        }

        if (string.IsNullOrEmpty(Password))
        {
            MessageBox.Show("Ingresá la contraseña.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            _txtPassword.Focus();
            return;
        }

        DialogResult = DialogResult.OK;
        Close();
    }
}
