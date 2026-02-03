using System.Text.Json;
using ConvertidorDeOrdenes.Core.Services;

namespace ConvertidorDeOrdenes.Desktop.Services;

public enum CompanyResolutionDecisionKind
{
    Unify = 1,
    Keep = 2,
    Ignore = 3
}

public sealed class CompanyResolutionDecision
{
    public CompanyResolutionDecisionKind Kind { get; set; }

    // Solo para Unify
    public int? TargetRowIndex { get; set; }
}

public sealed class CompanyResolutionDecisionStore
{
    private readonly string _path;
    private readonly object _gate = new();
    private Dictionary<string, CompanyResolutionDecision> _cache = new();
    private bool _loaded;

    public CompanyResolutionDecisionStore(string path)
    {
        _path = path;
    }

    public bool TryGet(string cuit, string sedeKey, out CompanyResolutionDecision decision)
    {
        EnsureLoaded();

        var key = MakeKey(cuit, sedeKey);
        lock (_gate)
        {
            return _cache.TryGetValue(key, out decision!);
        }
    }

    public void Save(string cuit, string sedeKey, CompanyResolutionDecision decision)
    {
        EnsureLoaded();

        var key = MakeKey(cuit, sedeKey);
        lock (_gate)
        {
            _cache[key] = decision;
            PersistUnsafe();
        }
    }

    private void EnsureLoaded()
    {
        lock (_gate)
        {
            if (_loaded)
                return;

            _loaded = true;
            _cache = new Dictionary<string, CompanyResolutionDecision>(StringComparer.OrdinalIgnoreCase);

            try
            {
                if (!File.Exists(_path))
                    return;

                var json = File.ReadAllText(_path);
                var data = JsonSerializer.Deserialize<Dictionary<string, CompanyResolutionDecision>>(json);
                if (data != null)
                    _cache = new Dictionary<string, CompanyResolutionDecision>(data, StringComparer.OrdinalIgnoreCase);
            }
            catch
            {
                // No bloquear si falla el load.
                _cache = new Dictionary<string, CompanyResolutionDecision>(StringComparer.OrdinalIgnoreCase);
            }
        }
    }

    private void PersistUnsafe()
    {
        try
        {
            var dir = Path.GetDirectoryName(_path);
            if (!string.IsNullOrWhiteSpace(dir))
                Directory.CreateDirectory(dir);

            var json = JsonSerializer.Serialize(_cache, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_path, json);
        }
        catch
        {
            // No bloquear si falla el persist.
        }
    }

    private static string MakeKey(string cuit, string sedeKey)
    {
        var digits = CuitUtils.ExtractDigits(cuit);
        return $"{digits}|{(sedeKey ?? string.Empty).Trim()}";
    }
}
