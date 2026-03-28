using System.Text.Json;

namespace ConvertidorDeOrdenes.Desktop.Services;

public sealed class PortalLoginStateStore
{
    private readonly string _path;

    public PortalLoginStateStore(string path)
    {
        _path = path;
    }

    public PortalLoginState Load()
    {
        try
        {
            if (!File.Exists(_path))
                return new PortalLoginState();

            var json = File.ReadAllText(_path);
            var state = JsonSerializer.Deserialize<PortalLoginState>(json);
            return state ?? new PortalLoginState();
        }
        catch
        {
            return new PortalLoginState();
        }
    }

    public void Save(PortalLoginState state)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_path) ?? string.Empty);
            var json = JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_path, json);
        }
        catch
        {
            // no-op
        }
    }
}

public sealed class PortalLoginState
{
    public string RememberedUsername { get; set; } = string.Empty;
}