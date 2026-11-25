using System.Runtime.InteropServices;

namespace ProjectionManager.Services;

/// <summary>單一顯示器的資訊。座標為實體像素（行程已設為 PerMonitorV2）。</summary>
public sealed class MonitorInfo
{
    public IntPtr Handle = IntPtr.Zero;
    public string DeviceName = string.Empty;
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
    public bool IsPrimary;

    public int Width => Right - Left;
    public int Height => Bottom - Top;

    public override string ToString()
    {
        var flag = IsPrimary ? "　（主螢幕）" : string.Empty;
        return $"{DeviceName}　{Width}×{Height}{flag}";
    }
}

/// <summary>列舉顯示器。以 EnumDisplayMonitors 實作，不依賴 WinForms。</summary>
public static class DisplayService
{
    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct MONITORINFOEX
    {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string szDevice;
    }

    private delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdc, ref RECT rect, IntPtr data);

    [DllImport("user32.dll")]
    private static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr clip, MonitorEnumProc callback, IntPtr data);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX info);

    /// <summary>取得目前連接的所有顯示器，順序與系統列舉一致。</summary>
    public static IReadOnlyList<MonitorInfo> GetMonitors()
    {
        var list = new List<MonitorInfo>();

        bool OnMonitor(IntPtr hMonitor, IntPtr hdc, ref RECT rect, IntPtr data)
        {
            var info = new MONITORINFOEX
            {
                cbSize = Marshal.SizeOf<MONITORINFOEX>(),
                szDevice = string.Empty,
            };

            if (GetMonitorInfo(hMonitor, ref info))
            {
                list.Add(new MonitorInfo
                {
                    Handle = hMonitor,
                    DeviceName = info.szDevice,
                    Left = info.rcMonitor.Left,
                    Top = info.rcMonitor.Top,
                    Right = info.rcMonitor.Right,
                    Bottom = info.rcMonitor.Bottom,
                    IsPrimary = (info.dwFlags & 1) != 0,
                });
            }

            return true;
        }

        MonitorEnumProc callback = OnMonitor;
        EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, callback, IntPtr.Zero);
        GC.KeepAlive(callback);
        return list;
    }
}
