using System.Text.Json.Serialization;

namespace ConvertidorDeOrdenes.Desktop.Services.Updates;

public sealed class UpdateInfo
{
    public required Version Version { get; init; }
    public required string InstallerUrl { get; init; }
    public string? InstallerApiUrl { get; init; }
    public required string ReleaseUrl { get; init; }
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
