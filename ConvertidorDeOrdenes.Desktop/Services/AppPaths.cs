using System.Reflection;

namespace ConvertidorDeOrdenes.Desktop.Services;

public static class AppPaths
{
    public const string VendorName = "Seres Salud";
    public const string AppName = "ConvertidorDeOrdenes";

    public static string InstallDirectory => AppDomain.CurrentDomain.BaseDirectory;

    // Per-user (evita problemas de permisos en Program Files)
    public static string DataRootDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        VendorName,
        AppName);

    public static string DbDirectory => Path.Combine(DataRootDirectory, "DB");
    public static string CompaniesFilePath => Path.Combine(DbDirectory, "Empresas.xlsx");

    // Archivo "semilla" (se instala junto al exe)
    public static string SeedCompaniesFilePath => Path.Combine(InstallDirectory, "DB", "Empresas.xlsx");

    public static string LogsDirectory => Path.Combine(DataRootDirectory, "logs");
    public static string UpdateStatePath => Path.Combine(DataRootDirectory, "update_state.json");

    public static Version GetCurrentVersion()
    {
        // Preferir AssemblyInformationalVersion si existe
        var informational = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion;

        if (!string.IsNullOrWhiteSpace(informational))
        {
            // Puede venir con metadata (ej: 1.2.3+sha). Cortar en '+'
            var clean = informational.Split('+')[0];
            if (Version.TryParse(clean.TrimStart('v', 'V'), out var vInfo))
                return vInfo;
        }

        return Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0, 0, 0);
    }
}
