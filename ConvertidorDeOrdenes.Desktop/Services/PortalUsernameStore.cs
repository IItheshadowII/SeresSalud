using System.Security.Cryptography;
using System.Text;

namespace ConvertidorDeOrdenes.Desktop.Services;

/// <summary>
/// Almacena el nombre de usuario del portal cifrado con DPAPI (igual que PortalPasswordStore).
/// </summary>
internal static class PortalUsernameStore
{
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("ConvertidorDeOrdenes.PortalUsername.v1");

    public static string? TryLoad(string path)
    {
        try
        {
            if (!File.Exists(path))
                return null;

            var protectedBytes = File.ReadAllBytes(path);
            if (protectedBytes.Length == 0)
                return null;

            var bytes = ProtectedData.Unprotect(protectedBytes, Entropy, DataProtectionScope.CurrentUser);
            var username = Encoding.UTF8.GetString(bytes);
            return string.IsNullOrWhiteSpace(username) ? null : username.Trim();
        }
        catch
        {
            return null;
        }
    }

    public static bool TrySave(string path, string username)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path) ?? string.Empty);

            var bytes = Encoding.UTF8.GetBytes(username.Trim());
            var protectedBytes = ProtectedData.Protect(bytes, Entropy, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(path, protectedBytes);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static bool TryClear(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
