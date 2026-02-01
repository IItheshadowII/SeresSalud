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

        return new UpdateInfo
        {
            Version = latest,
            InstallerUrl = asset.browser_download_url,
            ReleaseUrl = release.html_url ?? $"https://github.com/{Owner}/{Repo}/releases"
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

    public async Task<string?> DownloadInstallerAsync(UpdateInfo update, IProgress<(long received, long? total)>? progress, CancellationToken ct)
    {
        var token = GetGitHubToken();

        using var request = new HttpRequestMessage(HttpMethod.Get, update.InstallerUrl);
        request.Headers.UserAgent.Add(new ProductInfoHeaderValue("ConvertidorDeOrdenes", "1.0"));

        if (!string.IsNullOrWhiteSpace(token))
        {
            // En releases privadas puede requerirse auth.
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }

        using var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct);
        if (!response.IsSuccessStatusCode)
            return null;

        var tempPath = Path.Combine(Path.GetTempPath(), $"ConvertidorDeOrdenes-Setup-{update.Version}.exe");

        await using var input = await response.Content.ReadAsStreamAsync(ct);
        await using var output = File.Create(tempPath);

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

        return tempPath;
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

    private static string? GetGitHubToken()
    {
        // No hardcodear tokens. Para releases privadas, setear env var en la PC corporativa.
        return Environment.GetEnvironmentVariable("SERESSALUD_GITHUB_TOKEN")
            ?? Environment.GetEnvironmentVariable("GITHUB_TOKEN");
    }
}
