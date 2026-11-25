using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using ProjectionManager.Services;

namespace ProjectionManager;

/// <summary>
/// 素材放映視窗：無邊框、釘在最上層、不啟用（不搶焦點）。
/// 用原生 SetWindowPos 定位而不是 WPF 的 Left/Top/Width/Height：
/// 投影螢幕的 DPI 可能與作業螢幕不同，WPF 的裝置獨立座標會縮放，原生座標才是實體像素。
/// </summary>
public partial class MaterialWindow : Window
{
    public MaterialWindow()
    {
        InitializeComponent();
    }

    /// <summary>把某一頁顯示在指定螢幕上；視窗未建立時先建立。</summary>
    public void ShowPage(MonitorInfo target, string? imagePath)
    {
        if (!IsVisible)
        {
            // 先給一個極小的起始位置再以原生座標定位，避免在作業螢幕閃一下
            Left = target.Left;
            Top = target.Top;
            Width = 320;
            Height = 180;
            Show();
        }

        IntPtr hwnd = new WindowInteropHelper(this).Handle;
        Native.SetNoActivate(hwnd, true);
        Native.SetTopmost(hwnd, true);

        var bitmap = imagePath is null ? null : ImageCache.Load(imagePath);
        SlideImage.Source = bitmap;
        EmptyText.Visibility = bitmap is null ? Visibility.Visible : Visibility.Collapsed;

        Native.SetWindowPos(hwnd, Native.HWND_TOPMOST,
            target.Left, target.Top, target.Width, target.Height,
            Native.SWP_NOACTIVATE | Native.SWP_SHOWWINDOW);
    }

    /// <summary>換頁但不重新定位（翻頁時避免每頁都動一次視窗）。</summary>
    public void SetPage(string? imagePath)
    {
        if (!IsVisible)
        {
            return;
        }

        var bitmap = imagePath is null ? null : ImageCache.Load(imagePath);
        SlideImage.Source = bitmap;
        EmptyText.Visibility = bitmap is null ? Visibility.Visible : Visibility.Collapsed;
        Native.SetTopmost(new WindowInteropHelper(this).Handle, true);
    }

    /// <summary>蓋回 PowerPoint 之上：PowerPoint 有時會在重繪時跑到最上層。</summary>
    public void KeepOnTop()
    {
        if (!IsVisible)
        {
            return;
        }

        Native.SetTopmost(new WindowInteropHelper(this).Handle, true);
    }

    public void HideWindow()
    {
        Hide();
    }
}
