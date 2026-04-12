using System.Diagnostics;
using System.Net.Http.Headers;

namespace ConvertidorDeOrdenes.Desktop.Services.Updates;

public sealed class UpdateService
{
    private const string Owner = "IItheshadowII";
    private const string Repo = "SeresSalud";

    private const string PreferredInstallerName = "ConvertidorDeOrdenes-Setup.exe";

    private readonly HttpClient _http;
    private readonly GitHubReleaseClient _client;
    private readonly UpdateStateStore _stateStore;

    public UpdateService(HttpClient http, UpdateStateStore stateStore)
    {
        _http = http;
        _client = new GitHubReleaseClient(http);
        _stateStore = stateStore;
    }

    public bool ShouldAutoCheck()
    {
        var state = _stateStore.Load();
        if (state.LastCheckedUtc == null)
            return true;

        // Corporate-friendly: una vez por día
        return DateTimeOffset.UtcNow - state.LastCheckedUtc.Value > TimeSpan.FromHours(24);
    }

    public async Task<UpdateInfo?> CheckLatestAsync(CancellationToken ct)
    {
        var token = GetGitHubToken();
        var release = await _client.GetLatestReleaseAsync(Owner, Repo, token, ct);

        var state = _stateStore.Load();
        state.LastCheckedUtc = DateTimeOffset.UtcNow;
        _stateStore.Save(state);

        if (release?.tag_name == null)
            return null;

        var latest = ParseVersion(release.tag_name);
        if (latest == null)
            return null;

        var current = AppPaths.GetCurrentVersion();
        if (latest <= current)
            return null;

        var asset = SelectInstallerAsset(release);
        if (asset?.browser_download_url == null)
            return null;

        var hashAsset = SelectHashAsset(release, asset.name);

        return new UpdateInfo
        {
            Version = latest,
            InstallerUrl = asset.browser_download_url,
            InstallerApiUrl = asset.url,
            ReleaseUrl = release.html_url ?? $"https://github.com/{Owner}/{Repo}/releases",
            InstallerHashUrl = hashAsset?.browser_download_url,
            InstallerHashApiUrl = hashAsset?.url,
        };
    }

    public bool ShouldNotify(UpdateInfo update)
    {
        var state = _stateStore.Load();
        if (string.Equals(state.LastNotifiedVersion, update.Version.ToString(), StringComparison.OrdinalIgnoreCase))
            return false;

        return true;
    }

    public void MarkNotified(UpdateInfo update)
    {
        var state = _stateStore.Load();
        state.LastNotifiedVersion = update.Version.ToString();
        _stateStore.Save(state);
    }

    public async Task<(string? path, string? error)> DownloadInstallerAsync(UpdateInfo update, IProgress<(long received, long? total)>? progress, CancellationToken ct)
    {
        var token = GetGitHubToken();

        // Para repos privados, conviene descargar assets vía API con Accept: application/octet-stream
        // usando el token, para evitar flujos de descarga que requieren sesión/cookies.
        var downloadUrl = (!string.IsNullOrWhiteSpace(token) && !string.IsNullOrWhiteSpace(update.InstallerApiUrl))
            ? update.InstallerApiUrl
            : update.InstallerUrl;

        using var request = new HttpRequestMessage(HttpMethod.Get, downloadUrl);
        request.Headers.UserAgent.Add(new ProductInfoHeaderValue("ConvertidorDeOrdenes", "1.0"));

        if (!string.IsNullOrWhiteSpace(token))
        {
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

            if (!string.IsNullOrWhiteSpace(update.InstallerApiUrl))
            {
                request.Headers.Accept.Clear();
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/octet-stream"));
            }
        }

        using var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct);
        if (!response.IsSuccessStatusCode)
            return (null, $"El servidor devolvió {(int)response.StatusCode} al descargar el instalador.");

        var tempPath = Path.Combine(Path.GetTempPath(), $"ConvertidorDeOrdenes-Setup-{update.Version}.exe");

        await using var input = await response.Content.ReadAsStreamAsync(ct);
        await using (var output = File.Create(tempPath))
        {
            var buffer = new byte[81920];
            long received = 0;
            var total = response.Content.Headers.ContentLength;

            while (true)
            {
                var read = await input.ReadAsync(buffer, ct);
                if (read <= 0)
                    break;

                await output.WriteAsync(buffer.AsMemory(0, read), ct);
                received += read;
                progress?.Report((received, total));
            }
        }

        // Verificar integridad SHA-256 si hay un asset de hash disponible.
        var hashUrl = (!string.IsNullOrWhiteSpace(token) && !string.IsNullOrWhiteSpace(update.InstallerHashApiUrl))
            ? update.InstallerHashApiUrl
            : update.InstallerHashUrl;

        if (!string.IsNullOrWhiteSpace(hashUrl))
        {
            var verifyError = await VerifyInstallerHashAsync(tempPath, hashUrl, token, ct);
            if (verifyError != null)
            {
                try { File.Delete(tempPath); } catch { }
                return (null, verifyError);
            }
        }

        return (tempPath, null);
    }

    private async Task<string?> VerifyInstallerHashAsync(string installerPath, string hashUrl, string? token, CancellationToken ct)
    {
        try
        {
            using var req = new HttpRequestMessage(HttpMethod.Get, hashUrl);
            req.Headers.UserAgent.Add(new ProductInfoHeaderValue("ConvertidorDeOrdenes", "1.0"));
            if (!string.IsNullOrWhiteSpace(token))
            {
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/octet-stream"));
            }

            using var resp = await _http.SendAsync(req, ct);
            if (!resp.IsSuccessStatusCode)
                return $"No se pudo descargar el archivo de verificación de integridad (HTTP {(int)resp.StatusCode}).";

            var hashFileContent = (await resp.Content.ReadAsStringAsync(ct)).Trim();
            // Formato: "<hex>  <filename>" o simplemente "<hex>"
            var expectedHex = hashFileContent.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries)[0].ToLowerInvariant();

            if (expectedHex.Length != 64 || !System.Text.RegularExpressions.Regex.IsMatch(expectedHex, "^[0-9a-f]{64}$"))
                return "El archivo de verificación de integridad tiene un formato inesperado.";

            using var sha = System.Security.Cryptography.SHA256.Create();
            await using var fileStream = File.OpenRead(installerPath);
            var hashBytes = await sha.ComputeHashAsync(fileStream, ct);
            var actualHex = Convert.ToHexString(hashBytes).ToLowerInvariant();

            if (!string.Equals(actualHex, expectedHex, StringComparison.Ordinal))
                return "La verificación de integridad del instalador falló. El archivo puede estar corrupto o fue alterado.";

            return null; // OK
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return $"Error al verificar la integridad del instalador: {ex.Message}";
        }
    }

    public bool RunInstallerAndExit(string installerPath)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = installerPath,
                UseShellExecute = true,
            };

            Process.Start(psi);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static Version? ParseVersion(string tag)
    {
        var clean = tag.Trim().TrimStart('v', 'V');
        return Version.TryParse(clean, out var v) ? v : null;
    }

    private static GitHubAssetDto? SelectInstallerAsset(GitHubReleaseDto release)
    {
        var assets = release.assets;
        if (assets == null || assets.Count == 0)
            return null;

        // Preferir nombre exacto
        var exact = assets.FirstOrDefault(a => string.Equals(a.name, PreferredInstallerName, StringComparison.OrdinalIgnoreCase));
        if (exact != null)
            return exact;

        // Fallback: cualquier .exe que contenga "Setup"
        var fallback = assets.FirstOrDefault(a =>
            !string.IsNullOrWhiteSpace(a.name) &&
            a.name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) &&
            a.name.Contains("setup", StringComparison.OrdinalIgnoreCase));

        return fallback;
    }

    private static GitHubAssetDto? SelectHashAsset(GitHubReleaseDto release, string? installerName)
    {
        var assets = release.assets;
        if (assets == null || assets.Count == 0 || string.IsNullOrWhiteSpace(installerName))
            return null;

        // Buscar el archivo "<nombre_instalador>.sha256"
        var expected = installerName + ".sha256";
        return assets.FirstOrDefault(a => string.Equals(a.name, expected, StringComparison.OrdinalIgnoreCase));
    }

    private static string? GetGitHubToken()
    {
        var stored = GitHubTokenStore.TryLoad(AppPaths.UpdateTokenPath);
        if (!string.IsNullOrWhiteSpace(stored))
            return stored;

        // Fallback: variable de entorno específica de la aplicación.
        // No se usa GITHUB_TOKEN genérico para evitar que tokens de CI/CD con permisos
        // amplios sean consumidos involuntariamente por la app de escritorio.
        return Environment.GetEnvironmentVariable("SERESSALUD_GITHUB_TOKEN");
    }
}
