namespace ConvertidorDeOrdenes.Desktop.Services;

public static class AppInitializer
{
    public static void Initialize()
    {
        Directory.CreateDirectory(AppPaths.DataRootDirectory);
        Directory.CreateDirectory(AppPaths.DbDirectory);
        Directory.CreateDirectory(AppPaths.LogsDirectory);

        EnsureSeeded(AppPaths.SeedCompaniesFilePath, AppPaths.CompaniesFilePath);
    }

    private static void EnsureSeeded(string seedFile, string targetFile)
    {
        try
        {
            if (File.Exists(targetFile))
                return;

            if (File.Exists(seedFile))
            {
                File.Copy(seedFile, targetFile, overwrite: false);
                return;
            }

            // Si no hay semilla, dejar que el repositorio cree el archivo vacío.
        }
        catch
        {
            // No bloquear la app si falla el seed.
        }
    }
}
