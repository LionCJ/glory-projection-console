namespace ProjectionManager.Services;

/// <summary>把鼠標限制在單一螢幕（通常為作業螢幕），避免誤入投影螢幕。</summary>
public static class CursorLockService
{
    public static bool Lock(MonitorInfo monitor)
    {
        var rect = new Native.RECT
        {
            Left = monitor.Left,
            Top = monitor.Top,
            Right = monitor.Right,
            Bottom = monitor.Bottom,
        };

        return Native.ClipCursor(ref rect);
    }

    public static void Unlock()
    {
        Native.ClipCursor(IntPtr.Zero);
    }
}
