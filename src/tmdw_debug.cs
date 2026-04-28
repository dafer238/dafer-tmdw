using System;
using System.IO;
using ExcelDna.Integration;

public static partial class CoolPropWrapper
{
    private static volatile bool _debugEnabled = false;
    private static string _logFilePath = null;
    private static readonly object _logLock = new object();

    [ExcelFunction(Name = "CP_EnableDebug",
        Description = "Enable or disable debug logging to CoolPropWrapper_debug.log in the add-in directory. " +
                      "When enabled, DLL load status, input errors, and computation failures are written to the log. " +
                      "Pass TRUE to enable, FALSE to disable.")]
    public static object CP_EnableDebug(bool enable)
    {
        if (enable)
        {
            try
            {
                string dir;
                try { dir = Path.GetDirectoryName(ExcelDnaUtil.XllPath) ?? Directory.GetCurrentDirectory(); }
                catch { dir = Directory.GetCurrentDirectory(); }

                _logFilePath = Path.Combine(dir, "CoolPropWrapper_debug.log");
                _debugEnabled = true;

                lock (_logLock)
                {
                    File.AppendAllText(_logFilePath,
                        $"{Environment.NewLine}[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ========== Debug session started =========={Environment.NewLine}");
                }

                LogDebug($"XLL path: {ExcelDnaUtil.XllPath ?? "unavailable"}");
                LogDebug($"CoolProp DLL function pointers loaded: {(_PropsSI != null ? "Yes" : "Not yet (lazy-loaded on first call)")}");
                return $"Debug logging enabled. Log: {_logFilePath}";
            }
            catch (Exception ex)
            {
                _debugEnabled = false;
                return $"Error enabling debug logging: {ex.Message}";
            }
        }
        else
        {
            LogDebug("========== Debug session ended ==========");
            _debugEnabled = false;
            return "Debug logging disabled.";
        }
    }

    [ExcelFunction(Name = "CP_GetLogPath", IsVolatile = true,
        Description = "Returns the path to the current debug log file, or a notice if logging is not enabled. Recalculates automatically on every F9 / auto-calc cycle.")]
    public static object CP_GetLogPath()
    {
        if (!_debugEnabled || string.IsNullOrEmpty(_logFilePath))
            return "Debug logging is not enabled. Call CP_EnableDebug(TRUE) first.";
        return _logFilePath;
    }

    internal static void LogDebug(string message)
    {
        if (!_debugEnabled || string.IsNullOrEmpty(_logFilePath)) return;
        try
        {
            lock (_logLock)
            {
                File.AppendAllText(_logFilePath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}");
            }
        }
        catch { /* Never let logging crash the add-in */ }
    }
}
