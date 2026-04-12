using System;
using System.Collections.Generic;
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

public sealed class BulkCuitOnlineLookupForm : Form
{
    private readonly List<string> _targetContracts;
    public Dictionary<string, string> FoundCuits { get; private set; } = new();

    private readonly WebView2 _webView = new();
    private readonly Button _btnStart = new();
    private readonly ProgressBar _progress = new();
    private readonly Label _lblInfo = new();

    private bool _isScanning = false;

    public BulkCuitOnlineLookupForm(IEnumerable<string> targetContracts)
    {
        _targetContracts = targetContracts.Distinct().Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
        InitializeComponent();
    }

    private async void InitializeComponent()
    {
        Text = "Auto-Scanner Masivo de CUITs - La Segunda";
        Icon = AppIcon.TryGet();
        Size = new Size(1100, 700);
        MinimumSize = new Size(900, 600);
        StartPosition = FormStartPosition.CenterParent;

        var headerPanel = new Panel { Dock = DockStyle.Top, Height = 100, BackColor = SystemColors.Control };
        
        _lblInfo.Text = 
            $"Se encontraron {_targetContracts.Count} contratos únicos sin CUIT.\n" +
            "1. Inicia sesión en el portal si se solicita.\n" +
            "2. En la pantalla de /tray presione 'Iniciar Robot Scanner'.\n" +
            "El sistema tiperará a alta velocidad cada contrato filtrándolos y capturará el CUIT automáticamente.";
        _lblInfo.AutoSize = false;
        _lblInfo.Dock = DockStyle.Fill;
        _lblInfo.Padding = new Padding(10);
        _lblInfo.Font = new Font("Segoe UI", 9);

        _progress.Dock = DockStyle.Bottom;
        _progress.Height = 25;
        _progress.Maximum = _targetContracts.Count > 0 ? _targetContracts.Count : 1;

        _btnStart.Text = "▶ Iniciar Robot Scanner";
        _btnStart.Size = new Size(180, 40);
        _btnStart.Dock = DockStyle.Right;
        _btnStart.Margin = new Padding(0, 15, 20, 15);
        _btnStart.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        _btnStart.BackColor = Color.Navy;
        _btnStart.ForeColor = Color.White;
        _btnStart.FlatStyle = FlatStyle.Flat;
        _btnStart.Cursor = Cursors.Hand;
        _btnStart.Click += BtnStart_Click;

        headerPanel.Controls.Add(_lblInfo);
        headerPanel.Controls.Add(_btnStart);
        headerPanel.Controls.Add(_progress);
        
        _webView.Dock = DockStyle.Fill;
        Controls.Add(_webView);
        Controls.Add(headerPanel);

        try
        {
            Directory.CreateDirectory(AppPaths.WebView2UserDataDirectory);
            _webView.CreationProperties = new CoreWebView2CreationProperties
            {
                UserDataFolder = AppPaths.WebView2UserDataDirectory
            };
            await _webView.EnsureCoreWebView2Async();

            _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
            _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

            _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
            _webView.Source = new Uri("https://lasegundaart-ml.conexia.com.ar/tray");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error de WebView2: {ex.Message}", "Error");
            Close();
        }
    }

    private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            var payload = e.TryGetWebMessageAsString();
            if (string.IsNullOrWhiteSpace(payload)) return;
            var msg = JsonSerializer.Deserialize<JsonElement>(payload);
            if(msg.TryGetProperty("type", out var tProp) && tProp.GetString() == "progress")
            {
                var cuit = msg.GetProperty("cuit").GetString();
                var contract = msg.GetProperty("contract").GetString();
                int current = msg.GetProperty("current").GetInt32();
                
                if(!string.IsNullOrWhiteSpace(cuit) && cuit.Any(char.IsDigit) && contract != null)
                {
                    var cleanCuit = new string(cuit.Where(char.IsDigit).ToArray());
                    FoundCuits[contract] = cleanCuit.Length == 11 ? CuitUtils.FormatOrKeep(cleanCuit) : cleanCuit;
                }
                
                _progress.Value = current;
                _lblInfo.Text = $"Escaneando contrato {contract}... Encontrado: {cuit}";
            }
            else if (msg.TryGetProperty("type", out var typeDone) && typeDone.GetString() == "done")
            {
                DialogResult = DialogResult.OK;
                Close();
            }
            else if (msg.TryGetProperty("type", out var typeErr) && typeErr.GetString() == "error")
            {
                MessageBox.Show("Error en robot: " + msg.GetProperty("message").GetString());
                _isScanning = false;
                _btnStart.Text = "▶ Iniciar Robot Scanner";
                _btnStart.BackColor = Color.Navy;
            }
        }
        catch { }
    }

    private async void BtnStart_Click(object? sender, EventArgs e)
    {
        if (_targetContracts.Count == 0 || _webView.CoreWebView2 == null) return;
        if (_isScanning)
        {
            _btnStart.Text = "Deteniendo...";
            try { await _webView.CoreWebView2.ExecuteScriptAsync("window.__cancelRobot = true;"); } catch { }
            return;
        }

        var url = _webView.Source?.ToString() ?? string.Empty;
        if(!url.Contains("/tray", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show("Por favor, inicia sesión y dirígete a la bandeja /tray para comenzar.");
            return;
        }

        _isScanning = true;
        try { await _webView.CoreWebView2.ExecuteScriptAsync("window.__cancelRobot = false;"); } catch { }
        _btnStart.Text = "🛑 Detener Robot";
        _btnStart.BackColor = Color.DarkRed;
        _progress.Value = 0;

        var contractsJson = JsonSerializer.Serialize(_targetContracts);

        var script = $@"
            (async function() {{
                try {{
                    var contracts = {contractsJson};
                    var results = {{}};
                    var inputSearch = document.querySelector('input[placeholder*=""Filtro""], input[type=""search""], input[type=""text""]');
                    if(!inputSearch) return JSON.stringify({{ error: 'No input search found' }});
                    
                    var nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;

                    for(var i=0; i<contracts.length; i++) {{
                        if (window.__cancelRobot) break;

                        var c = contracts[i];
                        
                        nativeSetter.call(inputSearch, c);
                        inputSearch.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        inputSearch.dispatchEvent(new Event('change', {{ bubbles: true }}));

                        var cuitFound = '';
                        var maxWait = 15; // 15 intentos de 300ms (hasta 4.5s) para esperar a que cargue la lista
                        
                        for (var wait=0; wait < maxWait; wait++) {{
                            await new Promise(r => setTimeout(r, 300));
                            if (window.__cancelRobot) break;

                            var text = document.body.innerText || '';
                            var blocks = text.split('Operativo N°');
                            for(var j=0; j<blocks.length; j++) {{
                                var block = blocks[j];
                                if(block.indexOf('Nº de Contrato: ' + c) !== -1 || block.indexOf('N° de Contrato: ' + c) !== -1) {{
                                    var textMatch = block.match(/CUIT:\s*(\d+)/i);
                                    if(textMatch && textMatch[1]) {{
                                        cuitFound = textMatch[1];
                                        break;
                                    }}
                                }}
                            }}
                            if(cuitFound) break; // Terminar de esperar, ya lo encontramos
                        }}

                        if(cuitFound) results[c] = cuitFound;
                        
                        window.chrome.webview.postMessage(JSON.stringify({{ type: 'progress', contract: c, cuit: cuitFound || 'Nada', current: i+1 }}));
                    }}
                    window.chrome.webview.postMessage(JSON.stringify({{ type: 'done' }}));
                    return JSON.stringify({{ ok: true, dict: results }});
                }} catch(err) {{ 
                    window.chrome.webview.postMessage(JSON.stringify({{ type: 'error', message: err.message }}));
                    return JSON.stringify({{ error: err.message }}); 
                }}
            }})();
        ";

        try
        {
            await _webView.CoreWebView2.ExecuteScriptAsync(script);
            // Form is closed by WebMessageReceived when type == 'done'
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error al inyectar escaneo: " + ex.Message);
            _isScanning = false;
        }
    }
}
