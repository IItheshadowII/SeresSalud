using System.Text.Json.Serialization;

namespace ConvertidorDeOrdenes.Desktop.Services.Updates;

public sealed class UpdateInfo
{
    public required Version Version { get; init; }
    public required string InstallerUrl { get; init; }
    public string? InstallerApiUrl { get; init; }
    public required string ReleaseUrl { get; init; }
    /// <summary>URL pública (browser_download_url) del archivo .sha256 del instalador.</summary>
    public string? InstallerHashUrl { get; init; }
    /// <summary>URL de la API de GitHub del archivo .sha256 (para repos privados).</summary>
    public string? InstallerHashApiUrl { get; init; }
}

public sealed class UpdateState
{
    public DateTimeOffset? LastCheckedUtc { get; set; }
    public string? LastNotifiedVersion { get; set; }
}

[JsonSerializable(typeof(UpdateState))]
internal partial class UpdateJsonContext : JsonSerializerContext
{
}
