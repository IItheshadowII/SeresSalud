using System.Text.Json;

namespace ConvertidorDeOrdenes.Desktop.Services.Updates;

public sealed class UpdateStateStore
{
    private readonly string _path;

    public UpdateStateStore(string path)
    {
        _path = path;
    }

    public UpdateState Load()
    {
        try
        {
            if (!File.Exists(_path))
                return new UpdateState();

            var json = File.ReadAllText(_path);
            var state = JsonSerializer.Deserialize(json, UpdateJsonContext.Default.UpdateState);
            return state ?? new UpdateState();
        }
        catch
        {
            return new UpdateState();
        }
    }

    public void Save(UpdateState state)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_path) ?? string.Empty);
            var json = JsonSerializer.Serialize(state, UpdateJsonContext.Default.UpdateState);
            File.WriteAllText(_path, json);
        }
        catch
        {
            // no-op
        }
    }
}
