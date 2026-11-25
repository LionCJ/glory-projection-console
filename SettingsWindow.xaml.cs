using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Windows;
using System.Windows.Controls;
using ProjectionManager.Services;

namespace ProjectionManager;

public partial class SettingsWindow : Window
{
    private readonly AppSettings _settings;
    private bool _loading = true;

    public SettingsWindow(AppSettings settings)
    {
        _settings = settings;
        InitializeComponent();
        LoadCurrent();
        _loading = false;
    }

    private void LoadCurrent()
    {
        switch (_settings.UiTheme)
        {
            case ThemeManager.Classic:
                ClassicRadio.IsChecked = true;
                break;
            case ThemeManager.Light:
                LightRadio.IsChecked = true;
                break;
            default:
                BlackGoldRadio.IsChecked = true;
                break;
        }

        PresentationFolderText.Text = FolderText(_settings.LastPresentationFolder);
        MaterialFolderText.Text = FolderText(_settings.LastMaterialFolder);
        DebugModeCheck.IsChecked = _settings.DebugMode;
        SettingsPathText.Text = SettingsStore.SettingsPath;
        MaterialIndexText.Text = MaterialStore.IndexPath;
        MaterialCacheText.Text = MaterialStore.ImageRoot;
        PreviewCacheText.Text = PreviewRoot();
        LogPathText.Text = AppLog.DirectoryPath;
    }

    private static string FolderText(string? path) =>
        string.IsNullOrWhiteSpace(path) ? "尚未設定，使用系統上次位置" : path;

    private static string PreviewRoot() =>
        Path.Combine(Path.GetTempPath(), "ProjectionManager", "preview");

    private void Theme_Checked(object sender, RoutedEventArgs e)
    {
        if (_loading || sender is not RadioButton radio)
        {
            return;
        }

        int theme = radio.Tag switch
        {
            "1" => ThemeManager.Classic,
            "2" => ThemeManager.Light,
            _ => ThemeManager.BlackGold,
        };

        _settings.UiTheme = theme;
        ThemeManager.Apply(theme);
        SettingsStore.Save(_settings);
        AppLog.Info($"介面主題切換為 {theme}。", "SETTINGS");
    }

    private void DebugMode_Changed(object sender, RoutedEventArgs e)
    {
        if (_loading)
        {
            return;
        }

        _settings.DebugMode = DebugModeCheck.IsChecked == true;
        SettingsStore.Save(_settings);
        AppLog.Info($"調試模式（模擬屏幕 2）已{(_settings.DebugMode ? "開啟" : "關閉")}。", "SETTINGS");
    }

    private void BrowsePresentationFolder_Click(object sender, RoutedEventArgs e)
    {
        string? selected = PickFolder(_settings.LastPresentationFolder, "選擇主 PPT 的起始資料夾");
        if (selected is null)
        {
            return;
        }

        _settings.LastPresentationFolder = selected;
        PresentationFolderText.Text = selected;
        SettingsStore.Save(_settings);
    }

    private void ClearPresentationFolder_Click(object sender, RoutedEventArgs e)
    {
        _settings.LastPresentationFolder = null;
        PresentationFolderText.Text = FolderText(null);
        SettingsStore.Save(_settings);
    }

    private void BrowseMaterialFolder_Click(object sender, RoutedEventArgs e)
    {
        string? selected = PickFolder(_settings.LastMaterialFolder, "選擇素材的起始資料夾");
        if (selected is null)
        {
            return;
        }

        _settings.LastMaterialFolder = selected;
        MaterialFolderText.Text = selected;
        SettingsStore.Save(_settings);
    }

    private void ClearMaterialFolder_Click(object sender, RoutedEventArgs e)
    {
        _settings.LastMaterialFolder = null;
        MaterialFolderText.Text = FolderText(null);
        SettingsStore.Save(_settings);
    }

    private void OpenFolder_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement element)
        {
            return;
        }

        string path = element.Tag switch
        {
            "materialIndex" => MaterialStore.IndexPath,
            "materialCache" => MaterialStore.ImageRoot,
            "previewCache" => PreviewRoot(),
            "logs" => AppLog.DirectoryPath,
            _ => SettingsStore.SettingsPath,
        };

        try
        {
            if (File.Exists(path))
            {
                path = Path.GetDirectoryName(path) ?? path;
            }

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{path}\"",
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"無法開啟資料夾：{ex.Message}", "Glory 榮耀",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void ExportDiagnostics_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var zipPath = Path.Combine(desktop, $"Glory_診斷包_{stamp}.zip");

            using var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create);

            var info = new System.Text.StringBuilder()
                .AppendLine("Glory 榮耀 · 投影控制台 診斷資訊")
                .AppendLine($"時間：{DateTime.Now:yyyy-MM-dd HH:mm:ss}")
                .AppendLine($"版本：{GetType().Assembly.GetName().Version}")
                .AppendLine($"OS：{Environment.OSVersion}")
                .AppendLine($"電腦：{Environment.MachineName}")
                .AppendLine($"使用者：{Environment.UserName}")
                .AppendLine($"執行檔：{Environment.ProcessPath}")
                .AppendLine($"設定檔：{SettingsStore.SettingsPath}")
                .AppendLine($"素材索引：{MaterialStore.IndexPath}")
                .AppendLine($"日誌目錄：{AppLog.DirectoryPath}")
                .ToString();
            using (var writer = new StreamWriter(
                zip.CreateEntry("system-info.txt").Open(), new System.Text.UTF8Encoding(false)))
            {
                writer.Write(info);
            }

            foreach (var log in Directory.EnumerateFiles(AppLog.DirectoryPath, "*.log")
                         .OrderByDescending(File.GetLastWriteTime)
                         .Take(10))
            {
                AddToZip(zip, log, Path.Combine("logs", Path.GetFileName(log)));
            }

            AddToZip(zip, SettingsStore.SettingsPath, "settings.json");
            AddToZip(zip, MaterialStore.IndexPath, "materials.json");

            AppLog.Info($"已匯出診斷包：{zipPath}", "DIAG");
            MessageBox.Show(this,
                $"診斷包已匯出到桌面：\n{Path.GetFileName(zipPath)}\n\n可直接把這個 ZIP 交給技術支援。",
                "Glory 榮耀 · 診斷包",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{zipPath}\"",
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            AppLog.Error("匯出診斷包失敗。", ex, "DIAG");
            MessageBox.Show(this, $"匯出診斷包失敗：{ex.Message}", "Glory 榮耀",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private static void AddToZip(ZipArchive zip, string source, string entryName)
    {
        if (!File.Exists(source))
        {
            return;
        }

        var entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);
        using var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var output = entry.Open();
        input.CopyTo(output);
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    private string? PickFolder(string? initialFolder, string title)
    {
        try
        {
            // 仍保留設定頁指定的起始資料夾，但對話框根目錄固定為「本機」，
            // 可自由瀏覽 C/D 槽與外接碟，不被舊 Shell 對話框的上次位置限制。
            return FolderPicker.Pick(this, title, initialFolder, startAtComputerRoot: false);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"無法開啟資料夾選擇窗：{ex.Message}", "Glory 榮耀",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return null;
        }
    }
}
