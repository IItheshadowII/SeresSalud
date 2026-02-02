using ConvertidorDeOrdenes.Desktop.Services;
using ConvertidorDeOrdenes.Desktop.Services.Updates;

namespace ConvertidorDeOrdenes.Desktop.Forms;

public sealed class UpdateSettingsDialog : Form
{
    private readonly TextBox _txtToken;
    private readonly Label _lblStatus;
    private readonly Label _lblPath;
    private readonly CheckBox _chkShow;

    public bool StartUpdateCheck { get; private set; }

    public UpdateSettingsDialog()
    {
        Text = "Actualizaciones";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(12);

        var info = new Label
        {
            AutoSize = true,
            Text = "Si el repositorio de GitHub pasa a ser privado,\npuede necesitar un token para buscar y descargar releases.",
        };

        _lblStatus = new Label { AutoSize = true };
        _lblPath = new Label { AutoSize = true };

        var tokenLabel = new Label
        {
            AutoSize = true,
            Text = "Token de GitHub (PAT):",
            Margin = new Padding(0, 10, 0, 4)
        };

        _txtToken = new TextBox
        {
            Width = 360,
            UseSystemPasswordChar = true,
            PlaceholderText = "ghp_...",
        };

        _chkShow = new CheckBox
        {
            AutoSize = true,
            Text = "Mostrar token",
            Margin = new Padding(0, 6, 0, 0)
        };
        _chkShow.CheckedChanged += (_, _) => _txtToken.UseSystemPasswordChar = !_chkShow.Checked;

        var btnSave = new Button { Text = "Guardar token", AutoSize = true };
        btnSave.Click += (_, _) => SaveToken();

        var btnClear = new Button { Text = "Borrar token", AutoSize = true };
        btnClear.Click += (_, _) => ClearToken();

        var btnCheck = new Button { Text = "Buscar actualizaciones", AutoSize = true };
        btnCheck.Click += (_, _) =>
        {
            StartUpdateCheck = true;
            DialogResult = DialogResult.OK;
            Close();
        };

        var btnClose = new Button { Text = "Cerrar", AutoSize = true, DialogResult = DialogResult.Cancel };

        var buttons = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            WrapContents = false,
            Margin = new Padding(0, 12, 0, 0)
        };
        buttons.Controls.Add(btnSave);
        buttons.Controls.Add(btnClear);
        buttons.Controls.Add(btnCheck);
        buttons.Controls.Add(btnClose);

        var layout = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            AutoSize = true,
            WrapContents = false,
            Dock = DockStyle.Fill
        };

        layout.Controls.Add(info);
        layout.Controls.Add(_lblPath);
        layout.Controls.Add(_lblStatus);
        layout.Controls.Add(tokenLabel);
        layout.Controls.Add(_txtToken);
        layout.Controls.Add(_chkShow);
        layout.Controls.Add(buttons);

        Controls.Add(layout);

        RefreshStatus();
    }

    private void RefreshStatus()
    {
        var path = AppPaths.UpdateTokenPath;
        _lblPath.Text = $"Ubicación: {path}";

        var exists = File.Exists(path);
        var loaded = GitHubTokenStore.TryLoad(path, out var token, out var error);

        if (!exists)
        {
            _lblStatus.Text = "Estado: no existe archivo de token para este usuario.";
            return;
        }

        if (loaded)
        {
            _lblStatus.Text = "Estado: hay un token guardado y se puede leer.";
            return;
        }

        _lblStatus.Text = string.IsNullOrWhiteSpace(error)
            ? "Estado: existe archivo de token, pero no se pudo leer (posible corrupción)."
            : "Estado: existe archivo de token, pero no se pudo leer. Detalle: " + error;
    }

    private void SaveToken()
    {
        var token = _txtToken.Text;
        if (string.IsNullOrWhiteSpace(token))
        {
            MessageBox.Show(this, "Ingrese un token antes de guardar.", "Actualizaciones",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var ok = GitHubTokenStore.TrySave(AppPaths.UpdateTokenPath, token, out var error);
        if (!ok)
        {
            var detail = string.IsNullOrWhiteSpace(error) ? string.Empty : "\n\nDetalle: " + error;
            MessageBox.Show(this, "No se pudo guardar el token." + detail, "Actualizaciones",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _txtToken.Clear();
        RefreshStatus();

        // Verificación inmediata: leer lo que acabamos de guardar.
        var verifyOk = GitHubTokenStore.TryLoad(AppPaths.UpdateTokenPath, out var _, out var verifyError);
        if (!verifyOk)
        {
            var detail = string.IsNullOrWhiteSpace(verifyError) ? string.Empty : "\n\nDetalle: " + verifyError;
            MessageBox.Show(this, "Se guardó el archivo, pero luego no se pudo leer." + detail, "Actualizaciones",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show(this, "Token guardado para este usuario.", "Actualizaciones",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void ClearToken()
    {
        var ok = GitHubTokenStore.TryClear(AppPaths.UpdateTokenPath, out var error);
        if (!ok)
        {
            var detail = string.IsNullOrWhiteSpace(error) ? string.Empty : "\n\nDetalle: " + error;
            MessageBox.Show(this, "No se pudo borrar el token." + detail, "Actualizaciones",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _txtToken.Clear();
        RefreshStatus();

        MessageBox.Show(this, "Token borrado.", "Actualizaciones",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
