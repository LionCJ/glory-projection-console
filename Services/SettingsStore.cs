using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ProjectionManager.Services;

/// <summary>本機設定：綁定的翻頁筆、投影螢幕、上次使用的簡報。</summary>
public sealed class AppSettings
{
    public string? BoundDeviceName { get; set; }
    public string? TargetMonitorDevice { get; set; }
    public string? LastPresentation { get; set; }

    /// <summary>以「視窗中瀏覽」模式放映：PowerPoint 失焦時動畫仍會繼續播放。</summary>
    public bool WindowedShow { get; set; } = true;

    /// <summary>放映期間控制台不搶焦點：點按鈕不會讓 PowerPoint 失去前台，動畫不中斷。</summary>
    public bool NoActivateConsole { get; set; } = true;

    /// <summary>UI 主題：0 = 黑金（預設）、1 = 現有深灰青綠、2 = 白色背景。</summary>
    public int UiTheme { get; set; }

    /// <summary>調試模式：只有一台實體螢幕時，用右半螢幕模擬螢幕 2，方便在家測試推送與放映。</summary>
    public bool DebugMode { get; set; }

    /// <summary>上次「加入主 PPT」開啟的資料夾，作為下次選檔的起始位置。</summary>
    public string? LastPresentationFolder { get; set; }

    /// <summary>上次「加入資料夾」匯入素材的資料夾。</summary>
    public string? LastMaterialFolder { get; set; }

    /// <summary>左欄被收合的分組 Id 清單（重啟後保持收合狀態）。</summary>
    public List<string>? CollapsedGroups { get; set; }

    /// <summary>手機遙控器（V2.0）：是否啟動內建 HTTP 服務。</summary>
    public bool MobileControlEnabled { get; set; } = true;

    /// <summary>手機遙控器監聽的連接埠。</summary>
    public int MobileControlPort { get; set; } = 5180;

    /// <summary>已配對並綁定的手機。Token 撤銷後即失效。</summary>
    public List<MobileDevice>? MobileDevices { get; set; }
}

public static class SettingsStore
{
    private static readonly string Folder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ProjectionManager");

    private static readonly string FilePath =
        Environment.GetEnvironmentVariable("PROJECTIONMANAGER_SETTINGS")
        ?? Path.Combine(Folder, "settings.json");

    public static string AppFolder => Folder;

    public static string SettingsPath => FilePath;

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(FilePath))
            {
                var json = File.ReadAllText(FilePath);
                return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            }
        }
        catch
        {
            // 設定檔損壞時回到預設值，不阻斷啟動
        }

        return new AppSettings();
    }

    public static void Save(AppSettings settings)
    {
        try
        {
            Directory.CreateDirectory(Folder);
            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(FilePath, json);
        }
        catch
        {
            // 寫入失敗不影響使用，僅本次不保存
        }
    }
}
