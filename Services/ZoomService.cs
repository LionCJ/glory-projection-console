using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace ProjectionManager.Services;

/// <summary>
/// 把 Zoom 會議視窗搬到投影螢幕並鋪滿（Zoom 投影源）。
/// 不去改 Zoom 的分享設定：分享哪個螢幕仍由使用者在 Zoom 裡選。
/// </summary>
public sealed class ZoomService
{
    private IntPtr _hwnd = IntPtr.Zero;
    private long _style;
    private long _exStyle;
    private Native.RECT _rect;
    private bool _active;

    public bool IsActive => _active;
    public string WindowTitle { get; private set; } = string.Empty;
    public int WindowCount { get; private set; }

    public bool IsRunning => FindZoomProcesses().Count > 0;

    public string StatusText
    {
        get
        {
            if (_active)
            {
                return $"投影中：{WindowTitle}";
            }

            if (!IsRunning)
            {
                return "Zoom 未執行。按下方「開啟 Zoom」啟動，再進入會議。";
            }

            Refresh();
            return WindowCount > 0
                ? $"Zoom 執行中，找到 {WindowCount} 個可見視窗；按「推送投影」鋪到螢幕 2。"
                : "Zoom 執行中，但沒有可投影的會議視窗（可能還在大廳或視窗被最小化）。";
        }
    }

    /// <summary>重新盤點 Zoom 的本體與可見視窗，供右欄狀態與「重新抓取」使用。</summary>
    public void Refresh()
    {
        WindowCount = CountZoomWindows();
        var hwnd = FindZoomWindow();
        WindowTitle = hwnd == IntPtr.Zero ? string.Empty : TitleOf(hwnd);
    }

    public bool Launch(out string message)
    {
        var candidates = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Zoom", "bin", "Zoom.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Zoom", "bin", "Zoom.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Zoom", "bin", "Zoom.exe"),
        };

        foreach (var candidate in candidates)
        {
            if (!File.Exists(candidate))
            {
                continue;
            }

            try
            {
                Process.Start(new ProcessStartInfo(candidate) { UseShellExecute = true });
                message = "已啟動 Zoom，請登入並進入會議後再按「推送投影」。";
                return true;
            }
            catch (Exception ex)
            {
                message = "啟動 Zoom 失敗：" + ex.Message;
                return false;
            }
        }

        message = "找不到 Zoom 安裝路徑。請先安裝 Zoom，或手動開啟會議後再按「重新抓取」。";
        return false;
    }

    private static List<Process> FindZoomProcesses()
    {
        var result = new List<Process>();
        foreach (var process in Process.GetProcesses())
        {
            try
            {
                if (process.ProcessName.Equals("Zoom", StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(process);
                }
            }
            catch
            {
                // 存取失敗的程序直接略過
            }
        }

        return result;
    }

    private static int CountZoomWindows()
    {
        int count = 0;
        EnumWindows((h, _) =>
        {
            if (IsWindowVisible(h) && !string.IsNullOrWhiteSpace(TitleOf(h)) && IsZoomWindow(h))
            {
                count++;
            }

            return true;
        }, IntPtr.Zero);

        return count;
    }

    public bool Activate(MonitorInfo target, out string message)
    {
        if (_active && _hwnd != IntPtr.Zero && Native.IsWindow(_hwnd))
        {
            // 已經接管過：重新鋪滿（換螢幕或 Zoom 自己跑位時用）
            PlaceOn(target);
            message = "Zoom 已鋪在投影螢幕上。";
            return true;
        }

        _active = false;
        _hwnd = IntPtr.Zero;

        var hwnd = FindZoomWindow();
        if (hwnd == IntPtr.Zero)
        {
            message = "找不到 Zoom 會議視窗。請先開啟 Zoom 並進入會議（或把會議視窗還原成非最小化）。";
            return false;
        }

        _hwnd = hwnd;
        WindowTitle = TitleOf(hwnd);

        _style = GetWindowLong(hwnd, Native.GWL_STYLE);
        _exStyle = GetWindowLong(hwnd, Native.GWL_EXSTYLE);
        GetWindowRect(hwnd, out _rect);

        // 去標題列與邊框，讓它鋪滿整個投影螢幕；保留原本樣式以便還原
        long style = _style & ~(Native.WS_CAPTION | Native.WS_THICKFRAME
                                | Native.WS_SYSMENU | Native.WS_MINIMIZEBOX | Native.WS_MAXIMIZEBOX);
        SetWindowLong(hwnd, Native.GWL_STYLE, style);

        PlaceOn(target);

        _active = true;
        message = $"Zoom 已鋪在投影螢幕：{WindowTitle}";
        return true;
    }

    public void Restore()
    {
        if (!_active || _hwnd == IntPtr.Zero || !Native.IsWindow(_hwnd))
        {
            _active = false;
            _hwnd = IntPtr.Zero;
            return;
        }

        SetWindowLong(_hwnd, Native.GWL_STYLE, _style);
        SetWindowLong(_hwnd, Native.GWL_EXSTYLE, _exStyle);
        Native.SetTopmost(_hwnd, false);
        Native.SetNoActivate(_hwnd, false);
        Native.SetWindowPos(_hwnd, Native.HWND_NOTOPMOST,
            _rect.Left, _rect.Top, _rect.Right - _rect.Left, _rect.Bottom - _rect.Top,
            Native.SWP_FRAMECHANGED | Native.SWP_SHOWWINDOW);

        _active = false;
        _hwnd = IntPtr.Zero;
    }

    private void PlaceOn(MonitorInfo target)
    {
        Native.SetNoActivate(_hwnd, true);
        Native.SetTopmost(_hwnd, true);
        Native.SetWindowPos(_hwnd, Native.HWND_TOPMOST,
            target.Left, target.Top, target.Width, target.Height,
            Native.SWP_NOACTIVATE | Native.SWP_SHOWWINDOW | Native.SWP_FRAMECHANGED);
    }

    private static IntPtr FindZoomWindow()
    {
        IntPtr best = IntPtr.Zero;
        long bestArea = 0;

        EnumWindows((h, _) =>
        {
            if (!IsWindowVisible(h))
            {
                return true;
            }

            var title = TitleOf(h);
            if (string.IsNullOrWhiteSpace(title))
            {
                return true;
            }

            if (GetWindowRect(h, out var rect) is false)
            {
                return true;
            }

            long w = rect.Right - rect.Left;
            long hh = rect.Bottom - rect.Top;
            if (w < 320 || hh < 240)
            {
                return true;
            }

            if (!IsZoomWindow(h))
            {
                return true;
            }

            long area = w * hh;
            if (area > bestArea)
            {
                bestArea = area;
                best = h;
            }

            return true;
        }, IntPtr.Zero);

        return best;
    }

    private static bool IsZoomWindow(IntPtr hwnd)
    {
        int pid;
        GetWindowThreadProcessId(hwnd, out pid);
        if (pid == 0)
        {
            return false;
        }

        try
        {
            return Process.GetProcessById(pid).ProcessName
                .Equals("Zoom", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static string TitleOf(IntPtr hwnd)
    {
        var buffer = new StringBuilder(512);
        GetWindowText(hwnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder buffer, int maxCount);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out Native.RECT rect);

    [DllImport("user32.dll")]
    private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
    private static extern int GetWindowLong32(IntPtr hWnd, int index);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern IntPtr GetWindowLong64(IntPtr hWnd, int index);

    private static long GetWindowLong(IntPtr hWnd, int index) =>
        IntPtr.Size == 8 ? GetWindowLong64(hWnd, index).ToInt64() : GetWindowLong32(hWnd, index);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongW")]
    private static extern int SetWindowLong32(IntPtr hWnd, int index, int value);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
    private static extern IntPtr SetWindowLong64(IntPtr hWnd, int index, IntPtr value);

    private static void SetWindowLong(IntPtr hWnd, int index, long value)
    {
        if (IntPtr.Size == 8)
        {
            SetWindowLong64(hWnd, index, new IntPtr(value));
        }
        else
        {
            SetWindowLong32(hWnd, index, unchecked((int)value));
        }
    }
}
