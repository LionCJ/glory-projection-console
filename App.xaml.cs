using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using ProjectionManager.Services;

namespace ProjectionManager;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        AppLog.Initialize();
        AppLog.Info("載入主題與主視窗。", "START");

        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnDomainUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

        try
        {
            // 先建立主題畫刷，再載入 MainWindow 的 XAML；DynamicResource 才會有值。
            ThemeManager.Initialize();
            base.OnStartup(e);

            var window = new MainWindow();
            MainWindow = window;
            window.Show();
            AppLog.Info("主視窗已顯示。", "START");
        }
        catch (Exception ex)
        {
            AppLog.Error("主視窗啟動失敗。", ex, "START");
            MessageBox.Show(
                $"Glory 啟動失敗：{ex.Message}\n\n詳細錯誤已寫入：\n{AppLog.DirectoryPath}",
                "Glory 榮耀",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(1);
        }
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        AppLog.Error("UI 執行緒發生未處理例外。", e.Exception, "UI");
        MessageBox.Show(
            $"操作發生錯誤：{e.Exception.Message}\n\n已記錄到日誌，可在「設定 → 診斷與日誌」開啟資料夾。\n{AppLog.DirectoryPath}",
            "Glory 榮耀 · 錯誤記錄",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        e.Handled = true;
    }

    private void OnDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        AppLog.Error(
            e.IsTerminating ? "程式發生致命錯誤，即將結束。" : "程式發生未處理例外。",
            e.ExceptionObject as Exception,
            "FATAL");
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        AppLog.Error("背景工作發生未觀察的例外。", e.Exception, "TASK");
        e.SetObserved();
    }
}
