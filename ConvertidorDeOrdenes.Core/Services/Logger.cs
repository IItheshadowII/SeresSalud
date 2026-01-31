namespace ConvertidorDeOrdenes.Core.Services;

/// <summary>
/// Servicio de logging a archivos
/// </summary>
public class Logger
{
    private readonly string _logDirectory;
    private readonly string _logFilePath;

    public Logger(string baseDirectory)
    {
        _logDirectory = Path.Combine(baseDirectory, "logs");
        Directory.CreateDirectory(_logDirectory);

        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        _logFilePath = Path.Combine(_logDirectory, $"log_{timestamp}.txt");
    }

    public void LogInfo(string message)
    {
        WriteLog("INFO", message);
    }

    public void LogWarning(string message)
    {
        WriteLog("WARNING", message);
    }

    public void LogError(string message)
    {
        WriteLog("ERROR", message);
    }

    private void WriteLog(string level, string message)
    {
        try
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            var logEntry = $"[{timestamp}] [{level}] {message}";
            
            File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
        }
        catch
        {
            // Silenciar errores de escritura de log
        }
    }
}
