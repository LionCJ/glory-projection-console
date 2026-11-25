using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace ProjectionManager.Services;

/// <summary>一次原始鍵盤輸入。</summary>
public sealed class KeyInput
{
    public int VirtualKey { get; init; }
    public string DeviceName { get; init; } = string.Empty;
    public IntPtr DeviceHandle { get; init; }
}

/// <summary>
/// 以 Raw Input（RIDEV_INPUTSINK）旁聽實體鍵盤事件，即使本程式不在前台仍可收到。
/// 由事件附帶的裝置句柄取得裝置路徑，用以區分「講員的翻頁筆」與「操作員的鍵盤」。
/// 本服務只旁聽，不攔截；攔截由 KeyboardGuard 負責。
/// </summary>
public sealed class RawInputService : IDisposable
{
    private const int RIDEV_INPUTSINK = 0x00000100;
    private const uint RID_INPUT = 0x10000003;
    private const uint RIDI_DEVICENAME = 0x20000007;
    private const int RIM_TYPEKEYBOARD = 1;
    private const int RIM_TYPEHID = 2;

    // Consumer Control：部分簡報筆不是鍵盤集合，而是以多媒體鍵回報。
    private const ushort CONSUMER_NEXT_TRACK = 0x00B5;
    private const ushort CONSUMER_PREV_TRACK = 0x00B6;

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWINPUTDEVICE
    {
        public ushort usUsagePage;
        public ushort usUsage;
        public uint dwFlags;
        public IntPtr hwndTarget;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWINPUTHEADER
    {
        public uint dwType;
        public uint dwSize;
        public IntPtr hDevice;
        public IntPtr wParam;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWKEYBOARD
    {
        public ushort MakeCode;
        public ushort Flags;
        public ushort Reserved;
        public ushort VKey;
        public uint Message;
        public uint ExtraInformation;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWHID
    {
        public int SizeHid;
        public int Count;
    }

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool RegisterRawInputDevices(RAWINPUTDEVICE[] devices, int count, int size);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetRawInputData(IntPtr hRawInput, uint command, IntPtr data, ref int size, int headerSize);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int GetRawInputDeviceInfo(IntPtr hDevice, uint command, IntPtr data, ref int size);

    private HwndSource? _source;
    private readonly Dictionary<IntPtr, string> _deviceNames = new();
    private readonly int _headerSize = Marshal.SizeOf<RAWINPUTHEADER>();

    /// <summary>收到一次按鍵（僅 KeyDown，不含 KeyUp）。</summary>
    public event Action<KeyInput> KeyDown;

    public bool IsStarted => _source is not null;

    /// <summary>累計收到的原始按鍵事件數，用於診斷「有沒有收到」。</summary>
    public int ReceivedCount { get; private set; }

    /// <summary>
    /// 把原始輸入掛到指定的視窗上。必須用程式主視窗的 HWND：
    /// 自行建立隱藏視窗當目標時，WM_INPUT 在部分機器上不會派送，導致永遠收不到按鍵。
    /// </summary>
    public void Start(IntPtr targetWindow)
    {
        if (_source is not null)
        {
            return;
        }

        if (targetWindow == IntPtr.Zero)
        {
            throw new InvalidOperationException("目標視窗尚未建立，無法註冊原始輸入。");
        }

        _source = HwndSource.FromHwnd(targetWindow);
        if (_source is null)
        {
            throw new InvalidOperationException("無法取得目標視窗的訊息來源。");
        }

        _source.AddHook(WndProc);

        var keyboardDevices = new[]
        {
            new RAWINPUTDEVICE
            {
                usUsagePage = 0x01,     // Generic Desktop
                usUsage = 0x06,         // Keyboard
                dwFlags = RIDEV_INPUTSINK,
                hwndTarget = targetWindow,
            },
        };

        if (!RegisterRawInputDevices(keyboardDevices, keyboardDevices.Length, Marshal.SizeOf<RAWINPUTDEVICE>()))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error(), "註冊原始輸入裝置失敗。");
        }

        // Consumer Control 是加分項：部分簡報筆只以多媒體鍵回報。
        // 個別系統不支援時不可讓鍵盤監聽一起失敗。
        var consumerDevices = new[]
        {
            new RAWINPUTDEVICE
            {
                usUsagePage = 0x0C,     // Consumer Control
                usUsage = 0x01,
                dwFlags = RIDEV_INPUTSINK,
                hwndTarget = targetWindow,
            },
        };
        RegisterRawInputDevices(consumerDevices, consumerDevices.Length, Marshal.SizeOf<RAWINPUTDEVICE>());
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == Native.WM_INPUT)
        {
            HandleRawInput(lParam);
        }

        return IntPtr.Zero;
    }

    private void HandleRawInput(IntPtr rawInputHandle)
    {
        int size = 0;
        GetRawInputData(rawInputHandle, RID_INPUT, IntPtr.Zero, ref size, _headerSize);
        if (size <= 0)
        {
            return;
        }

        IntPtr buffer = Marshal.AllocHGlobal(size);
        try
        {
            int read = GetRawInputData(rawInputHandle, RID_INPUT, buffer, ref size, _headerSize);
            if (read <= 0)
            {
                return;
            }

            var header = Marshal.PtrToStructure<RAWINPUTHEADER>(buffer);
            if (header.dwType == RIM_TYPEHID)
            {
                HandleConsumerHid(buffer);
                return;
            }

            if (header.dwType != RIM_TYPEKEYBOARD)
            {
                return;
            }

            var keyboard = Marshal.PtrToStructure<RAWKEYBOARD>(
                new IntPtr(buffer.ToInt64() + _headerSize));

            if (keyboard.Message != Native.WM_KEYDOWN && keyboard.Message != Native.WM_SYSKEYDOWN)
            {
                return;
            }

            var input = new KeyInput
            {
                VirtualKey = keyboard.VKey,
                DeviceName = GetDeviceName(header.hDevice),
                DeviceHandle = header.hDevice,
            };

            ReceivedCount++;
            KeyDown?.Invoke(input);
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private void HandleConsumerHid(IntPtr buffer)
    {
        int hidOffset = _headerSize;
        var hid = Marshal.PtrToStructure<RAWHID>(new IntPtr(buffer.ToInt64() + hidOffset));
        if (hid.SizeHid <= 0 || hid.Count <= 0)
        {
            return;
        }

        int totalBytes = hid.SizeHid * hid.Count;
        if (totalBytes <= 0 || totalBytes > 256)
        {
            return;
        }

        int dataOffset = hidOffset + Marshal.SizeOf<RAWHID>();
        byte[] report = new byte[totalBytes];
        Marshal.Copy(new IntPtr(buffer.ToInt64() + dataOffset), report, 0, totalBytes);

        // Consumer Control 報告通常是 report id + 2 bytes usage；逐位移讀取
        // 0x00B5（下一首）／0x00B6（上一首），兩種簡報筆都能辨識。
        for (int i = 0; i + 1 < report.Length; i++)
        {
            ushort usage = (ushort)(report[i] | (report[i + 1] << 8));
            int vk = usage switch
            {
                CONSUMER_NEXT_TRACK => Native.VK_MEDIA_NEXT_TRACK,
                CONSUMER_PREV_TRACK => Native.VK_MEDIA_PREV_TRACK,
                _ => 0,
            };

            if (vk == 0)
            {
                continue;
            }

            var input = new KeyInput
            {
                VirtualKey = vk,
                DeviceName = GetDeviceName(Marshal.PtrToStructure<RAWINPUTHEADER>(buffer).hDevice),
                DeviceHandle = Marshal.PtrToStructure<RAWINPUTHEADER>(buffer).hDevice,
            };

            ReceivedCount++;
            KeyDown?.Invoke(input);
            return;
        }
    }

    private string GetDeviceName(IntPtr device)
    {
        if (_deviceNames.TryGetValue(device, out var cached))
        {
            return cached;
        }

        string name = string.Empty;
        int size = 0;

        GetRawInputDeviceInfo(device, RIDI_DEVICENAME, IntPtr.Zero, ref size);
        if (size > 0)
        {
            // RIDI_DEVICENAME 的 size 以字元計，配置時多留一個結尾
            IntPtr buffer = Marshal.AllocHGlobal((size + 1) * 2);
            try
            {
                GetRawInputDeviceInfo(device, RIDI_DEVICENAME, buffer, ref size);
                name = Marshal.PtrToStringUni(buffer) ?? string.Empty;
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        _deviceNames[device] = name;
        return name;
    }

    public void Dispose()
    {
        if (_source is not null)
        {
            // HwndSource 屬於主視窗，只移除掛鉤，不 Dispose，否則會連主視窗一起拆掉
            _source.RemoveHook(WndProc);
            _source = null;
        }

        // 還原為一般鍵盤註冊（不接收背景輸入）
        var devices = new[]
        {
            new RAWINPUTDEVICE
            {
                usUsagePage = 0x01,
                usUsage = 0x06,
                dwFlags = 0x00000001,   // RIDEV_REMOVE
                hwndTarget = IntPtr.Zero,
            },
            new RAWINPUTDEVICE
            {
                usUsagePage = 0x0C,
                usUsage = 0x01,
                dwFlags = 0x00000001,   // RIDEV_REMOVE
                hwndTarget = IntPtr.Zero,
            },
        };
        RegisterRawInputDevices(devices, devices.Length, Marshal.SizeOf<RAWINPUTDEVICE>());
    }
}
