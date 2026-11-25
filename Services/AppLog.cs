using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace ProjectionManager.Services;

/// <summary>
/// 全程式共用的事件與錯誤日誌。寫入 %LOCALAPPDATA%\Glory\logs，
/// 每次啟動寫一個 session 標記，錯誤含完整 Exception.ToString() 堆疊。
/// 單一日誌檔超過 8 MB 會輪替，並自動清除 30 天前的舊檔。
/// </summary>
public static class AppLog
{
    private static readonly object Sync = new();
    private static readonly UTF8Encoding Utf8NoBom = new(false);
    private const long MaxFileBytes = 8L * 1024 * 1024;
    private const int KeepDays = 30;

    private static bool _initialized;
    private static string _logDir = "";

    public static string DirectoryPath => _logDir;

    public static string CurrentFilePath
    {
        get
        {
            EnsureDirectory();
            return Path.Combine(_logDir, $"glory-{DateTime.Now:yyyyMMdd}.log");
        }
    }

    public static void Initialize()
    {
        if (_initialized)
        {
            return;
        }

        try
        {
            _logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Glory", "logs");
            EnsureDirectory();
            CleanupOldFiles();
            _initialized = true;

            var version = Assembly.GetExecutingAssembly()
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
                ?? Assembly.GetExecutingAssembly().GetName().Version?.ToString()
                ?? "unknown";
            Info($"=== Glory 啟動 · PID {Environment.ProcessId} · 版本 {version} ===", "START");
            Info($"OS {RuntimeInformation.OSDescription} ({RuntimeInformation.OSArchitecture})；" +
                 $".NET {RuntimeInformation.FrameworkDescription}", "START");
            Info($"執行檔 {Environment.ProcessPath}", "START");
            Info($"命令列 {string.Join(" ", Environment.GetCommandLineArgs())}", "START");
            Info($"日誌檔 {CurrentFilePath}", "START");
        }
        catch
        {
            // 日誌不能阻斷主程式啟動。
            _initialized = false;
        }
    }

    public static void Info(string message, string category = "APP") =>
        Write("INFO", category, message, null);

    public static void Warn(string message, Exception? exception = null, string category = "APP") =>
        Write("WARN", category, message, exception);

    public static void Error(string message, Exception? exception = null, string category = "APP") =>
        Write("ERROR", category, message, exception);

    private static void Write(string level, string category, string message, Exception? exception)
    {
        if (!_initialized)
        {
            return;
        }

        try
        {
            lock (Sync)
            {
                EnsureDirectory();
                var path = CurrentFilePath;
                RotateIfNeeded(path);

                var builder = new StringBuilder()
                    .Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                    .Append(" [").Append(level).Append(']')
                    .Append(" [").Append(category).Append("] ")
                    .Append(message);

                if (exception is not null)
                {
                    builder.AppendLine().Append(exception);
                }

                builder.AppendLine();
                File.AppendAllText(path, builder.ToString(), Utf8NoBom);
            }
        }
        catch
        {
            // 寫入失敗時不可再拋出，避免掩蓋原始錯誤。
        }
    }

    private static void EnsureDirectory()
    {
        if (string.IsNullOrWhiteSpace(_logDir))
        {
            _logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Glory", "logs");
        }

        Directory.CreateDirectory(_logDir);
    }

    private static void RotateIfNeeded(string path)
    {
        try
        {
            if (!File.Exists(path) || new FileInfo(path).Length < MaxFileBytes)
            {
                return;
            }

            var archive = Path.Combine(
                Path.GetDirectoryName(path) ?? _logDir,
                $"{Path.GetFileNameWithoutExtension(path)}-{DateTime.Now:HHmmss}.log");
            File.Move(path, archive, overwrite: true);
        }
        catch
        {
            // 輪替失敗時繼續 append，不能中斷程式。
        }
    }

    private static void CleanupOldFiles()
    {
        try
        {
            var cutoff = DateTime.Now.AddDays(-KeepDays);
            foreach (var file in Directory.EnumerateFiles(_logDir, "glory-*.log"))
            {
                if (File.GetLastWriteTime(file) < cutoff)
                {
                    File.Delete(file);
                }
            }
        }
        catch
        {
            // 清理失敗不影響啟動。
        }
    }
}
