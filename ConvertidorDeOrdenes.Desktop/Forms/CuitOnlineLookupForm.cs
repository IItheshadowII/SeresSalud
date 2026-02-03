using System;
using System.Drawing;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConvertidorDeOrdenes.Core.Services;
using ConvertidorDeOrdenes.Desktop.Services;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ConvertidorDeOrdenes.Desktop.Forms;

/// <summary>
/// Formulario con navegador embebido para buscar el CUIT en el portal de La Segunda
/// a partir de un número de contrato ya mostrado en pantalla.
/// </summary>
public sealed class CuitOnlineLookupForm : Form
{
    private readonly string _contractNumber;

    private readonly WebView2 _webView = new();
    private readonly Button _btnExtraerCuit = new();
    private readonly Label _lblInfo = new();

    private bool _loginAttempted;
    private bool _autoFlowStarted;

    /// <summary>
    /// CUIT encontrado (solo dígitos). Se establece cuando el usuario pulsa "Extraer CUIT".
    /// </summary>
    public string? FoundCuit { get; private set; }

    public CuitOnlineLookupForm(string contractNumber)
    {
        _contractNumber = contractNumber.Trim();
        InitializeComponent();
    }

    private async void InitializeComponent()
    {
        Text = "Buscar CUIT online - La Segunda";
        Icon = AppIcon.TryGet();
        Size = new Size(1100, 700);
        MinimumSize = new Size(900, 600);
        StartPosition = FormStartPosition.CenterParent;

        // Panel superior que contiene el texto de ayuda y el botón "Extraer CUIT".
        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 80,
            BackColor = SystemColors.Control
        };

        _lblInfo.Text =
            "1) Inicie sesión en el portal (usuario/contraseña no se guardan)." + Environment.NewLine +
            "2) Vaya a 'Consulta de Operativos' / bandeja / tray." + Environment.NewLine +
            $"3) Busque el N° de Contrato {_contractNumber} usando el buscador del portal." + Environment.NewLine +
            "4) Cuando vea el operativo en la lista, pulse 'Extraer CUIT' para copiarlo en esta orden.";
        _lblInfo.AutoSize = false;
        _lblInfo.Dock = DockStyle.Fill;
        _lblInfo.Padding = new Padding(10, 5, 10, 5);
        _lblInfo.Font = new Font("Segoe UI", 9);
        _lblInfo.ForeColor = Color.FromArgb(50, 50, 50);

        _btnExtraerCuit.Text = "Extraer CUIT";
        _btnExtraerCuit.Size = new Size(140, 32);
        _btnExtraerCuit.Dock = DockStyle.Right;
        _btnExtraerCuit.Margin = new Padding(0, 20, 20, 20);
        _btnExtraerCuit.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        _btnExtraerCuit.BackColor = Color.FromArgb(60, 140, 60);
        _btnExtraerCuit.ForeColor = Color.White;
        _btnExtraerCuit.FlatStyle = FlatStyle.Flat;
        _btnExtraerCuit.FlatAppearance.BorderSize = 0;
        _btnExtraerCuit.Cursor = Cursors.Hand;
        _btnExtraerCuit.Click += BtnExtraerCuit_Click;

        _webView.Dock = DockStyle.Fill;

        headerPanel.Controls.Add(_lblInfo);
        headerPanel.Controls.Add(_btnExtraerCuit);

        Controls.Add(_webView);
        Controls.Add(headerPanel);

        try
        {
            // IMPORTANTE: En instalaciones corporativas, la app suele vivir en Program Files.
            // WebView2 necesita escribir en su UserDataFolder, así que lo movemos a AppData.
            Directory.CreateDirectory(AppPaths.WebView2UserDataDirectory);
            _webView.CreationProperties = new CoreWebView2CreationProperties
            {
                UserDataFolder = AppPaths.WebView2UserDataDirectory
            };

            await _webView.EnsureCoreWebView2Async();
            _webView.CoreWebView2.Settings.IsStatusBarEnabled = true;
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

            _webView.CoreWebView2.NavigationCompleted += CoreWebView2OnNavigationCompleted;

            // Intentar ir directo a /tray; si no hay sesión, el servidor redirigirá a /login.
            _webView.Source = new Uri("https://lasegundaart-ml.conexia.com.ar/tray");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"No se pudo inicializar el navegador embebido (WebView2).\n\n{ex.Message}",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            DialogResult = DialogResult.Cancel;
            Close();
        }
    }

    private async void CoreWebView2OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        if (_webView.CoreWebView2 == null)
            return;

        var url = _webView.Source?.ToString() ?? string.Empty;

        // Si estamos en login y todavía no intentamos login automático, hacerlo.
        if (url.Contains("/login", StringComparison.OrdinalIgnoreCase) && !_loginAttempted)
        {
            _loginAttempted = true;
            await TryAutoLoginAsync();
            return;
        }

        // Cuando llegamos a /tray por primera vez, lanzar el flujo automático de búsqueda y extracción.
        if (url.Contains("/tray", StringComparison.OrdinalIgnoreCase) && !_autoFlowStarted)
        {
            _autoFlowStarted = true;

            // Esperar un poco a que cargue la grilla de operativos.
            await Task.Delay(3000);

            await AutoExtractCuitAndCloseAsync();
        }
    }

    private async void BtnExtraerCuit_Click(object? sender, EventArgs e)
    {
        await AutoExtractCuitAndCloseAsync(manual: true);
    }

    private async Task TryAutoLoginAsync()
    {
        try
        {
            if (_webView.CoreWebView2 == null)
                return;

            // Seguridad: NO hardcodear credenciales. Si se desea auto-login corporativo,
            // proveerlas por variables de entorno en la PC del usuario.
            var usuario = Environment.GetEnvironmentVariable("SERESSALUD_PORTAL_USER");
            var password = Environment.GetEnvironmentVariable("SERESSALUD_PORTAL_PASS");

            if (string.IsNullOrWhiteSpace(usuario) || string.IsNullOrWhiteSpace(password))
            {
                // Sin credenciales configuradas, el usuario se loguea manualmente.
                return;
            }

            var jsLogin =
                "(function() {" +
                "  try {" +
                "    var user = '" + EscapeForJsString(usuario) + "';" +
                "    var pass = '" + EscapeForJsString(password) + "';" +
                "    var userInput = document.querySelector('input[type=email],input[name*=" + "\"user\"" + "],input[name*=" + "\"Usuario\"" + "],input[name*=" + "\"usuario\"" + "]');" +
                "    var passInput = document.querySelector('input[type=password]');" +
                "    if (!userInput || !passInput) return 'no-fields';" +
                "    userInput.value = user;" +
                "    userInput.dispatchEvent(new Event('input', { bubbles: true }));" +
                "    passInput.value = pass;" +
                "    passInput.dispatchEvent(new Event('input', { bubbles: true }));" +
                "    var form = userInput.form || document.querySelector('form');" +
                "    if (form) { form.submit(); return 'submitted'; }" +
                "    var btn = document.querySelector('button[type=submit],button');" +
                "    if (btn) { btn.click(); return 'clicked'; }" +
                "    return 'no-submit';" +
                "  } catch (e) { return 'error'; }" +
                "})();";

            await _webView.CoreWebView2.ExecuteScriptAsync(jsLogin);
        }
        catch
        {
            // Si el login automático falla, el usuario aún puede loguearse manualmente.
        }
    }

    private async Task AutoExtractCuitAndCloseAsync(bool manual = false)
    {
        if (_webView.CoreWebView2 == null)
        {
            if (manual)
            {
                MessageBox.Show("El navegador todavía no está listo.", "Información",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return;
        }

        // JavaScript muy genérico: busca en el texto de la página una línea que contenga
        // "Nº de Contrato: <N>" y en ese mismo bloque busca la línea con "CUIT: XXXXX".
        var js =
            "(function() {" +
            "  try {" +
            "    var contract = '" + EscapeForJsString(_contractNumber) + "';" +
            "    var text = (document.body && document.body.innerText) ? document.body.innerText : '';" +
            "    if (!text) return null;" +
            "    var blocks = text.split('Operativo N°');" +
            "    for (var i = 0; i < blocks.length; i++) {" +
            "      var block = blocks[i];" +
            "      if (!block) continue;" +
            "      if (block.indexOf('Nº de Contrato: ' + contract) !== -1 ||" +
            "          block.indexOf('N° de Contrato: ' + contract) !== -1) {" +
            "        var m = block.match(/CUIT:\\s*(\\d+)/i);" +
            "        if (m && m[1]) return m[1];" +
            "      }" +
            "    }" +
            "    var m2 = text.match(/CUIT:\\s*(\\d+)/i);" +
            "    if (m2 && m2[1]) return m2[1];" +
            "    return null;" +
            "  } catch (e) { return null; }" +
            "})();";

        try
        {
            var raw = await _webView.ExecuteScriptAsync(js);
            // ExecuteScriptAsync devuelve un JSON string (entre comillas) o "null".
            string? cuit = null;
            if (!string.IsNullOrWhiteSpace(raw) && raw != "null")
            {
                try
                {
                    cuit = JsonSerializer.Deserialize<string>(raw);
                }
                catch
                {
                    // fallback: quitar comillas manualmente
                    cuit = raw.Trim('"');
                }
            }

            if (string.IsNullOrWhiteSpace(cuit))
            {
                if (manual)
                {
                    MessageBox.Show(
                        "No se pudo encontrar un CUIT en la página actual para ese N° de contrato.\n" +
                        "Asegúrese de que el operativo esté visible en la lista.",
                        "CUIT no encontrado",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                return;
            }

            // Normalizar a solo dígitos
            cuit = new string(cuit.Where(char.IsDigit).ToArray());
            string resultCuit = cuit;

            if (cuit.Length == 11)
            {
                // Forzar formato estándar xx-xxxxxxxx-x
                resultCuit = CuitUtils.FormatOrKeep(cuit);
            }
            else if (manual)
            {
                var msg = "Se encontró un valor con formato inesperado como CUIT: '" + cuit + "'." +
                          "\nDe todas formas se copiará tal cual para que pueda revisarlo.";
                MessageBox.Show(msg, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            FoundCuit = resultCuit;
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Exception ex)
        {
            if (manual)
            {
                MessageBox.Show(
                    "Ocurrió un error al intentar leer el contenido de la página.\n\n" + ex.Message,
                    "Error al extraer CUIT",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }

    private static string EscapeForJsString(string value)
    {
        return value
            .Replace("\\", "\\\\")
            .Replace("'", "\\'")
            .Replace("\"", "\\\"")
            .Replace("\r", " ")
            .Replace("\n", " ");
    }
}
