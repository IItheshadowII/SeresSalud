using System.Security.Cryptography;
using System.Text;

namespace ConvertidorDeOrdenes.Desktop.Services;

internal static class PortalPasswordStore
{
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("ConvertidorDeOrdenes.PortalPassword.v1");

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
            var password = Encoding.UTF8.GetString(bytes);
            return string.IsNullOrWhiteSpace(password) ? null : password;
        }
        catch
        {
            return null;
        }
    }

    public static bool TrySave(string path, string password)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path) ?? string.Empty);

            var bytes = Encoding.UTF8.GetBytes(password);
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