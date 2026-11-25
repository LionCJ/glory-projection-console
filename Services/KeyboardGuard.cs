using System.Runtime.InteropServices;

namespace ProjectionManager.Services;

/// <summary>
/// 雙觸發防護：當放映視窗意外取得焦點時，吞掉導航鍵，僅由 COM 推進一次。
/// 只在本程式判定「前台視窗即放映視窗」時生效；平時不攔截任何按鍵。
/// </summary>
public sealed class KeyboardGuard : IDisposable
{
    private const int WH_KEYBOARD_LL = 13;

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public uint vkCode;
        public uint scanCode;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc proc, IntPtr module, uint threadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hook);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hook, int nCode, IntPtr wParam, IntPtr lParam);

    private readonly LowLevelKeyboardProc _proc;
    private readonly Func<IntPtr> _showWindowProvider;
    private readonly Action<int>? _onGuardedPageKey;
    private IntPtr _hook = IntPtr.Zero;

    public bool IsInstalled => _hook != IntPtr.Zero;

    /// <summary>
    /// 綁定翻頁筆期間暫停吞鍵，否則放映視窗在前台時 PageUp/PageDown
    /// 會在低階鉤子被吃掉，原始輸入永遠收不到，導致「PgDn 無法綁定」。
    /// </summary>
    public bool Suspended { get; set; }

    /// <param name="showWindowProvider">回傳目前放映視窗句柄；未放映時回傳 IntPtr.Zero。</param>
    /// <param name="onGuardedPageKey">
    /// 被防護層吞掉的 PageUp/PageDown 改由這裡通知控制台翻頁；
    /// 否則按鍵被吞後 Raw Input 也收不到，會出現「綁定成功但 PgUp/PgDn 沒反應」。
    /// </param>
    public KeyboardGuard(Func<IntPtr> showWindowProvider, Action<int>? onGuardedPageKey = null)
    {
        _showWindowProvider = showWindowProvider;
        _onGuardedPageKey = onGuardedPageKey;
        _proc = Hook;
    }

    public void Install()
    {
        if (_hook != IntPtr.Zero)
        {
            return;
        }

        _hook = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, IntPtr.Zero, 0);
    }

    private IntPtr Hook(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && (wParam == (IntPtr)Native.WM_KEYDOWN || wParam == (IntPtr)Native.WM_SYSKEYDOWN))
        {
            var data = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
            int vk = (int)data.vkCode;

            if (!Suspended && (vk == Native.VK_PRIOR || vk == Native.VK_NEXT) && ShowHasFocus())
            {
                // 吞掉按鍵，避免 PowerPoint 自己翻頁；同時由控制台執行唯一一次導航。
                _onGuardedPageKey?.Invoke(vk);
                return new IntPtr(1);
            }
        }

        return CallNextHookEx(_hook, nCode, wParam, lParam);
    }

    private bool ShowHasFocus()
    {
        IntPtr show = _showWindowProvider();
        return show != IntPtr.Zero && Native.GetForegroundWindow() == show;
    }

    public void Dispose()
    {
        if (_hook != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hook);
            _hook = IntPtr.Zero;
        }
    }
}
