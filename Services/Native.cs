using System.Runtime.InteropServices;
using System.Text;

namespace ProjectionManager.Services;

internal static class Native
{
    public const uint SWP_NOACTIVATE = 0x0010;
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const uint SWP_FRAMECHANGED = 0x0020;
    public const uint SWP_NOMOVE = 0x0002;
    public const uint SWP_NOSIZE = 0x0001;

    public const int SW_HIDE = 0;
    public const int SW_SHOWMINNOACTIVE = 7;

    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

    /// <summary>顯示器設定變更（插拔投影線、切換解析度）時系統送出的訊息。</summary>
    public const int WM_DISPLAYCHANGE = 0x007E;

    // 視窗樣式（用來把「視窗中瀏覽」模式的放映視窗去掉標題列，鋪成無邊框）
    public const int GWL_STYLE = -16;
    public const int GWL_EXSTYLE = -20;
    public const long WS_CAPTION = 0x00C00000L;
    public const long WS_THICKFRAME = 0x00040000L;
    public const long WS_SYSMENU = 0x00080000L;
    public const long WS_MINIMIZEBOX = 0x00020000L;
    public const long WS_MAXIMIZEBOX = 0x00010000L;
    public const long WS_EX_NOACTIVATE = 0x08000000L;
    public const long WS_EX_TOOLWINDOW = 0x00000080L;

    public const int WM_KEYDOWN = 0x0100;
    public const int WM_SYSKEYDOWN = 0x0104;
    public const int WM_INPUT = 0x00FF;

    public const int VK_PRIOR = 0x21;  // PageUp
    public const int VK_NEXT = 0x22;   // PageDown
    public const int VK_LEFT = 0x25;   // ←
    public const int VK_UP = 0x26;     // ↑
    public const int VK_RIGHT = 0x27;  // →
    public const int VK_DOWN = 0x28;   // ↓
    public const int VK_SPACE = 0x20;  // 空白鍵（部分翻頁筆的下一頁）
    public const int VK_B = 0x42;
    public const int VK_MEDIA_NEXT_TRACK = 0xB0;  // 部分簡報筆以多媒體鍵回報
    public const int VK_MEDIA_PREV_TRACK = 0xB1;

    /// <summary>HWND_MESSAGE 父視窗，用於建立訊息專用視窗。</summary>
    public static readonly IntPtr HWND_MESSAGE = new IntPtr(-3);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SetWindowPos(IntPtr hWnd, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true, EntryPoint = "GetWindowLongW")]
    private static extern int GetWindowLong32(IntPtr hWnd, int index);

    [DllImport("user32.dll", SetLastError = true, EntryPoint = "SetWindowLongW")]
    private static extern int SetWindowLong32(IntPtr hWnd, int index, int value);

    [DllImport("user32.dll", SetLastError = true, EntryPoint = "GetWindowLongPtrW")]
    private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int index);

    [DllImport("user32.dll", SetLastError = true, EntryPoint = "SetWindowLongPtrW")]
    private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int index, IntPtr value);

    public static long GetWindowLong(IntPtr hWnd, int index)
    {
        return IntPtr.Size == 8
            ? GetWindowLongPtr64(hWnd, index).ToInt64()
            : GetWindowLong32(hWnd, index);
    }

    public static void SetWindowLong(IntPtr hWnd, int index, long value)
    {
        if (IntPtr.Size == 8)
        {
            SetWindowLongPtr64(hWnd, index, new IntPtr(value));
        }
        else
        {
            SetWindowLong32(hWnd, index, unchecked((int)value));
        }
    }

    /// <summary>去掉視窗的標題列與可調整邊框，讓「視窗中瀏覽」的放映看起來像全螢幕。</summary>
    public static void StripWindowFrame(IntPtr hWnd)
    {
        long style = GetWindowLong(hWnd, GWL_STYLE);
        long target = style & ~(WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX);
        if (target != style)
        {
            SetWindowLong(hWnd, GWL_STYLE, target);
            SetWindowPos(hWnd, IntPtr.Zero, 0, 0, 0, 0,
                SWP_NOACTIVATE | SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE);
        }
    }

    /// <summary>把視窗釘在最上層：素材畫面要蓋住 PowerPoint 放映視窗，靠這一層決定誰在上。</summary>
    public static void SetTopmost(IntPtr hWnd, bool on)
    {
        if (hWnd == IntPtr.Zero || !IsWindow(hWnd))
        {
            return;
        }

        SetWindowPos(hWnd, on ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0,
            SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
    }

    /// <summary>開啟／關閉「不啟用」樣式：開啟後點這個視窗不會把前台搶走。</summary>
    public static void SetNoActivate(IntPtr hWnd, bool on)
    {
        if (hWnd == IntPtr.Zero || !IsWindow(hWnd))
        {
            return;
        }

        long ex = GetWindowLong(hWnd, GWL_EXSTYLE);
        long target = on ? (ex | WS_EX_NOACTIVATE) : (ex & ~WS_EX_NOACTIVATE);
        if (target != ex)
        {
            SetWindowLong(hWnd, GWL_EXSTYLE, target);
        }
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ClipCursor(ref RECT rect);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ClipCursor(IntPtr rect);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder buffer, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder buffer, int maxCount);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    /// <summary>PowerPoint 放映視窗的視窗類別名稱。</summary>
    private const string SlideShowWindowClass = "screenClass";

    /// <summary>
    /// 找出 PowerPoint 的放映視窗。
    /// 不能用 SlideShowWindow.HWND：在部分 PowerPoint 版本上該屬性會回傳空值。
    /// 放映視窗類別固定為 screenClass；開啟「簡報者檢視」時會有兩個同類別視窗，
    /// 需排除標題含「簡報者 / Presenter」的那一個。
    /// </summary>
    /// <summary>
    /// 判斷某個視窗「現在」是不是放映視窗（不只是曾經是）。
    /// 全螢幕放映類別是 screenClass；視窗化放映是 PPTFrameClass、但標題帶「投影片放映」。
    /// 放映結束後同一個視窗會退回簡報編輯視窗（標題形如 xxx.pptx:1 - PowerPoint），
    /// 因此可用它判斷放映是否已被 PowerPoint 端結束。
    /// </summary>
    public static bool IsSlideShowWindow(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero || !IsWindow(hWnd) || !IsWindowVisible(hWnd))
        {
            return false;
        }

        var cls = new StringBuilder(256);
        GetClassName(hWnd, cls, cls.Capacity);

        var title = new StringBuilder(512);
        GetWindowText(hWnd, title, title.Capacity);
        string text = title.ToString();

        if (text.Contains("簡報者", StringComparison.Ordinal) ||
            text.Contains("Presenter", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (string.Equals(cls.ToString(), SlideShowWindowClass, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return text.Contains("投影片放映", StringComparison.Ordinal)
            || text.Contains("Slide Show", StringComparison.OrdinalIgnoreCase)
            || text.Contains("Slideshow", StringComparison.OrdinalIgnoreCase);
    }

    public static IntPtr FindSlideShowWindow()
    {
        IntPtr best = IntPtr.Zero;
        long bestArea = -1;
        IntPtr largest = IntPtr.Zero;
        long largestArea = -1;

        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd))
            {
                return true;
            }

            var cls = new StringBuilder(256);
            GetClassName(hWnd, cls, cls.Capacity);
            string className = cls.ToString();

            // 全螢幕放映是 screenClass；視窗化放映（PPTFrameClass）標題仍帶「投影片放映」，
            // 所以不能只看類別，還要接受標題命中的視窗。
            bool isShowClass = string.Equals(className, SlideShowWindowClass, StringComparison.OrdinalIgnoreCase);

            var title = new StringBuilder(512);
            GetWindowText(hWnd, title, title.Capacity);
            string text = title.ToString();

            bool looksLikeShow =
                text.Contains("投影片放映", StringComparison.Ordinal) ||
                text.Contains("Slide Show", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("Slideshow", StringComparison.OrdinalIgnoreCase);

            if (!isShowClass && !looksLikeShow)
            {
                return true;
            }

            if (!GetWindowRect(hWnd, out var rect))
            {
                return true;
            }

            long area = (long)(rect.Right - rect.Left) * (rect.Bottom - rect.Top);

            if (area > largestArea)
            {
                largestArea = area;
                largest = hWnd;
            }

            // 簡報者檢視不是投影畫面，排除
            if (text.Contains("簡報者", StringComparison.Ordinal) ||
                text.Contains("Presenter", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (area > bestArea)
            {
                bestArea = area;
                best = hWnd;
            }

            return true;
        }, IntPtr.Zero);

        // 標題比對不到（非中文或其他語系介面）時，退而取面積最大的 screenClass 視窗
        return best != IntPtr.Zero ? best : largest;
    }

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int command);

    /// <summary>
    /// 找出指定 PowerPoint 進程的編輯主視窗（PPTFrameClass）。
    /// 放映視窗是 screenClass，不是這個；用它可以在放映結束時立刻藏住編輯視窗。
    /// </summary>
    public static IntPtr FindPowerPointFrameWindow(int processId)
    {
        IntPtr found = IntPtr.Zero;

        EnumWindows((hWnd, _) =>
        {
            uint pid;
            GetWindowThreadProcessId(hWnd, out pid);
            if (processId != 0 && pid != (uint)processId)
            {
                return true;
            }

            var cls = new StringBuilder(256);
            GetClassName(hWnd, cls, cls.Capacity);
            if (!string.Equals(cls.ToString(), "PPTFrameClass", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            found = hWnd;
            return false;
        }, IntPtr.Zero);

        return found;
    }
}
