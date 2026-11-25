using System.Windows;
using System.Windows.Interop;
using Microsoft.Web.WebView2.Wpf;
using ProjectionManager.Services;

namespace ProjectionManager;

/// <summary>
/// 網頁投影視窗：WebView2 全螢幕蓋在投影螢幕上，無邊框、置頂、不搶焦點。
/// 只做「顯示＋上下滾動」，不做互動輸入；操作仍由控制台按鈕驅動。
/// </summary>
public partial class WebWindow : Window
{
    private readonly WebView2 _browser;
    private bool _coreReady;

    public WebWindow()
    {
        InitializeComponent();

        // 由程式碼建立 WebView2，避免 XAML 編譯器在部分 SDK 版本下
        // 找不到 Microsoft.Web.WebView2.Wpf 的 XAML 型別。
        _browser = new WebView2();
        BrowserHost.Content = _browser;
        _browser.NavigationCompleted += (_, _) =>
        {
            Hint.Visibility = Visibility.Collapsed;
            KeepOnTop();
        };
    }

    public string? CurrentUrl { get; private set; }

    public async void ShowUrl(MonitorInfo target, string url)
    {
        CurrentUrl = url;
        Hint.Visibility = Visibility.Visible;

        if (!IsVisible)
        {
            Left = target.Left;
            Top = target.Top;
            Width = 640;
            Height = 360;
            Show();
        }

        PlaceOn(target);

        try
        {
            if (!_coreReady)
            {
                await _browser.EnsureCoreWebView2Async();
                _browser.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                _browser.CoreWebView2.Settings.IsStatusBarEnabled = false;
                _browser.CoreWebView2.Settings.AreDevToolsEnabled = false;
                _coreReady = true;
            }

            _browser.Source = new Uri(url);
        }
        catch (Exception ex)
        {
            Hint.Visibility = Visibility.Visible;
            Hint.Text = "無法載入網頁：" + ex.Message;
        }
    }

    public void Reload()
    {
        if (_coreReady)
        {
            _browser.Reload();
        }
    }

    public void Back()
    {
        if (_coreReady && _browser.CoreWebView2.CanGoBack)
        {
            _browser.CoreWebView2.GoBack();
        }
    }

    public void Forward()
    {
        if (_coreReady && _browser.CoreWebView2.CanGoForward)
        {
            _browser.CoreWebView2.GoForward();
        }
    }

    /// <summary>翻頁筆在網頁源時當作上下滾動（一頁約 80% 視窗高）。</summary>
    public void Scroll(bool down)
    {
        if (!_coreReady)
        {
            return;
        }

        double step = Math.Max(200, ActualHeight * 0.8);
        _browser.CoreWebView2.ExecuteScriptAsync(
            $"window.scrollBy(0, {(down ? step : -step):0})");
    }

    public void KeepOnTop()
    {
        if (IsVisible)
        {
            Native.SetTopmost(new WindowInteropHelper(this).Handle, true);
        }
    }

    public void HideWindow() => Hide();

    private void PlaceOn(MonitorInfo target)
    {
        IntPtr hwnd = new WindowInteropHelper(this).Handle;
        Native.SetNoActivate(hwnd, true);
        Native.SetTopmost(hwnd, true);
        Native.SetWindowPos(hwnd, Native.HWND_TOPMOST,
            target.Left, target.Top, target.Width, target.Height,
            Native.SWP_NOACTIVATE | Native.SWP_SHOWWINDOW);
    }

    protected override void OnClosed(EventArgs e)
    {
        _browser.Dispose();
        base.OnClosed(e);
    }
}
