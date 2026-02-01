using System.Net.Http.Headers;
using System.Text.Json;

namespace ConvertidorDeOrdenes.Desktop.Services.Updates;

internal sealed class GitHubReleaseClient
{
    private readonly HttpClient _http;

    public GitHubReleaseClient(HttpClient http)
    {
        _http = http;

        // Requerido por la API de GitHub
        if (!_http.DefaultRequestHeaders.UserAgent.Any())
        {
            _http.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("ConvertidorDeOrdenes", "1.0"));
        }

        _http.DefaultRequestHeaders.Accept.Clear();
        _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
    }

    public async Task<GitHubReleaseDto?> GetLatestReleaseAsync(string owner, string repo, string? token, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, $"https://api.github.com/repos/{owner}/{repo}/releases/latest");

        if (!string.IsNullOrWhiteSpace(token))
        {
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }

        using var response = await _http.SendAsync(request, ct);
        if (!response.IsSuccessStatusCode)
            return null;

        await using var stream = await response.Content.ReadAsStreamAsync(ct);
        return await JsonSerializer.DeserializeAsync<GitHubReleaseDto>(stream, cancellationToken: ct);
    }
}

internal sealed class GitHubReleaseDto
{
    public string? tag_name { get; set; }
    public string? html_url { get; set; }
    public bool prerelease { get; set; }
    public List<GitHubAssetDto>? assets { get; set; }
}

internal sealed class GitHubAssetDto
{
    public string? name { get; set; }
    public string? browser_download_url { get; set; }
}
