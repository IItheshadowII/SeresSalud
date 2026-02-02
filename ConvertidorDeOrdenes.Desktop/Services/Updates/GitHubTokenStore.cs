using System.Security.Cryptography;
using System.Text;

namespace ConvertidorDeOrdenes.Desktop.Services.Updates;

internal static class GitHubTokenStore
{
    // DPAPI (Windows) - cifra por usuario. No es "secreto perfecto" pero evita texto plano.
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("ConvertidorDeOrdenes.GitHubToken.v1");

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
            var token = Encoding.UTF8.GetString(bytes);
            return string.IsNullOrWhiteSpace(token) ? null : token.Trim();
        }
        catch
        {
            return null;
        }
    }

    public static bool TryLoad(string path, out string? token, out string? errorMessage)
    {
        try
        {
            errorMessage = null;
            token = TryLoad(path);
            return token != null;
        }
        catch (Exception ex)
        {
            token = null;
            errorMessage = ex.Message;
            return false;
        }
    }

    public static bool TrySave(string path, string token)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path) ?? string.Empty);

            var bytes = Encoding.UTF8.GetBytes(token.Trim());
            var protectedBytes = ProtectedData.Protect(bytes, Entropy, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(path, protectedBytes);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static bool TrySave(string path, string token, out string? errorMessage)
    {
        try
        {
            errorMessage = null;
            return TrySave(path, token);
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
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

    public static bool TryClear(string path, out string? errorMessage)
    {
        try
        {
            errorMessage = null;
            return TryClear(path);
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            return false;
        }
    }
}
