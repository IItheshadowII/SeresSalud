namespace ConvertidorDeOrdenes.Desktop.Forms;

public sealed class UpdateSettingsDialog : Form
{
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
            Text = "Buscar actualizaciones y descargar el instalador más reciente desde GitHub Releases.",
        };

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
        layout.Controls.Add(buttons);

        Controls.Add(layout);
    }

    /*
     * ==============================
     *  TOKEN DE GITHUB (DESHABILITADO)
     * ==============================
     *
     * Se ocultó toda la UI y el flujo de guardado/borrado de token porque,
     * mientras el repositorio sea público, no es necesario pedir un PAT.
     *
     * Si en el futuro el repo pasa a privado, se puede reutilizar este bloque.
     *
     * Código previo (dejar comentado para reactivar):
     *
     * using ConvertidorDeOrdenes.Desktop.Services;
     * using ConvertidorDeOrdenes.Desktop.Services.Updates;
     *
     * private readonly TextBox _txtToken;
     * private readonly Label _lblStatus;
     * private readonly Label _lblPath;
     * private readonly CheckBox _chkShow;
     *
     * // En el constructor:
     * // var info = new Label { Text = "Si el repositorio de GitHub pasa a ser privado,..." };
     * // _lblStatus/_lblPath
     * // tokenLabel + _txtToken + _chkShow
     * // botones: Guardar token / Borrar token
     * // layout.Controls.Add(...) para esos controles
     * // RefreshStatus();
     *
     * private void RefreshStatus() { ... AppPaths.UpdateTokenPath ... GitHubTokenStore.TryLoad ... }
     * private void SaveToken() { ... GitHubTokenStore.TrySave ... }
     * private void ClearToken() { ... GitHubTokenStore.TryClear ... }
     */
}
