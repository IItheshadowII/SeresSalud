using System.Drawing;

namespace ConvertidorDeOrdenes.Desktop.Services;

public static class AppIcon
{
    private static Icon? _cached;

    public static Icon? TryGet()
    {
        if (_cached != null)
            return _cached;

        try
        {
            var iconPath = Path.Combine(AppPaths.InstallDirectory, "icon.ico");
            if (File.Exists(iconPath))
            {
                _cached = new Icon(iconPath);
                return _cached;
            }
        }
        catch
        {
            // ignore
        }

        try
        {
            // Fallback al ícono embebido del exe
            _cached = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            return _cached;
        }
        catch
        {
            return null;
        }
    }

    public static void Apply(Form form)
    {
        try
        {
            var icon = TryGet();
            if (icon != null)
                form.Icon = icon;
        }
        catch
        {
            // ignore
        }
    }
}
