using System.Windows;
using System.Windows.Media;

namespace ProjectionManager.Services;

/// <summary>
/// 全站 UI 主題。三套色票在同一個 Brush 物件上換色，
/// XAML 用 DynamicResource、程式碼用 ThemeManager.Brush(...)，切換時不需要重建視窗。
/// </summary>
public static class ThemeManager
{
    public const int BlackGold = 0;
    public const int Classic = 1;
    public const int Light = 2;

    public static readonly string[] Keys =
    {
        "WindowBrush", "PanelBrush", "Panel2Brush", "Panel3Brush", "LineBrush",
        "TextPrimaryBrush", "TextDimBrush", "TextFaintBrush",
        "TealBrush", "AccentSoftBrush", "AccentBorderBrush", "AccentTextBrush",
        "AmberBrush", "DangerBrush", "OperatorBrush", "SpeakerBrush",
        "ButtonBrush", "ButtonHoverBrush", "InputBrush",
        "PreviewBrush", "PreviewPanelBrush", "LogBrush", "LogTextBrush",
        "BadgeBrush", "MenuHoverBrush", "DisabledTextBrush"
    };

    private static readonly Dictionary<string, SolidColorBrush> Brushes = new();
    private static int _current = BlackGold;

    public static int Current => _current;

    public static event EventHandler? ThemeChanged;

    public static string DisplayName(int theme) => theme switch
    {
        Classic => "現有配色（深灰青綠）",
        Light => "白色背景（暖白 + 深金）",
        _ => "黑金（預設）",
    };

    public static void Initialize()
    {
        EnsureBrushes();
        Apply(SettingsStore.Load().UiTheme);
    }

    public static void Apply(int theme)
    {
        EnsureBrushes();
        _current = theme is Classic or Light ? theme : BlackGold;

        var palette = _current switch
        {
            Classic => ClassicPalette,
            Light => LightPalette,
            _ => BlackGoldPalette,
        };

        foreach (var pair in palette)
        {
            // ResourceDictionary 內的 Freezable 可能被 WPF 凍結，不能再改 Color；
            // 每次切換建立新 Brush 並覆蓋同一個 key，DynamicResource 會立即重新解析。
            var brush = new SolidColorBrush(Parse(pair.Value));
            Brushes[pair.Key] = brush;
            if (Application.Current is { } app)
            {
                app.Resources[pair.Key] = brush;
            }
        }

        ThemeChanged?.Invoke(null, EventArgs.Empty);
    }

    public static SolidColorBrush Brush(string key)
    {
        EnsureBrushes();
        return Brushes.TryGetValue(key, out var brush) ? brush : Brushes["TextPrimaryBrush"];
    }

    private static void EnsureBrushes()
    {
        if (Brushes.Count > 0)
        {
            return;
        }

        foreach (var key in Keys)
        {
            Brushes[key] = new SolidColorBrush(Colors.Transparent);
        }

        var app = Application.Current;
        if (app is null)
        {
            return;
        }

        foreach (var pair in Brushes)
        {
            app.Resources[pair.Key] = pair.Value;
        }
    }

    private static Color Parse(string hex)
    {
        var s = hex.TrimStart('#');
        if (s.Length == 6)
        {
            s = "FF" + s;
        }

        return Color.FromArgb(
            (byte)Convert.ToInt32(s.Substring(0, 2), 16),
            (byte)Convert.ToInt32(s.Substring(2, 2), 16),
            (byte)Convert.ToInt32(s.Substring(4, 2), 16),
            (byte)Convert.ToInt32(s.Substring(6, 2), 16));
    }

    // 黑金：純黑底、奶油金文字、古銅金邊框。文字與背景對比維持在 WCAG AA 以上。
    private static readonly Dictionary<string, string> BlackGoldPalette = new()
    {
        ["WindowBrush"] = "#000000",
        ["PanelBrush"] = "#0A0A0A",
        ["Panel2Brush"] = "#141414",
        ["Panel3Brush"] = "#1C1B18",
        ["LineBrush"] = "#3D3320",
        ["TextPrimaryBrush"] = "#F7ECD2",
        ["TextDimBrush"] = "#D8C39A",
        ["TextFaintBrush"] = "#A69064",
        ["TealBrush"] = "#E8B84A",
        ["AccentSoftBrush"] = "#2B2110",
        ["AccentBorderBrush"] = "#C9973A",
        ["AccentTextBrush"] = "#FFE7A3",
        ["AmberBrush"] = "#F4C95D",
        ["DangerBrush"] = "#FF6B5E",
        // 控制者角色：操作員＝主金色；講員＝天空藍，兩者在黑底上都能一眼分辨。
        ["OperatorBrush"] = "#E8B84A",
        ["SpeakerBrush"] = "#5FD3FF",
        ["ButtonBrush"] = "#191817",
        ["ButtonHoverBrush"] = "#2B2417",
        ["InputBrush"] = "#0D0D0C",
        ["PreviewBrush"] = "#000000",
        ["PreviewPanelBrush"] = "#0A0A09",
        ["LogBrush"] = "#050505",
        ["LogTextBrush"] = "#D8C39A",
        ["BadgeBrush"] = "#CC000000",
        ["MenuHoverBrush"] = "#2B2417",
        ["DisabledTextBrush"] = "#6F6249",
    };

    // 舊版 V2.2：深灰底 + 青綠重點色。
    private static readonly Dictionary<string, string> ClassicPalette = new()
    {
        ["WindowBrush"] = "#0F1216",
        ["PanelBrush"] = "#171B21",
        ["Panel2Brush"] = "#232932",
        ["Panel3Brush"] = "#1B2027",
        ["LineBrush"] = "#3A424D",
        ["TextPrimaryBrush"] = "#F5F7FA",
        ["TextDimBrush"] = "#C7D0DB",
        ["TextFaintBrush"] = "#96A1B0",
        ["TealBrush"] = "#47D1C4",
        ["AccentSoftBrush"] = "#12403B",
        ["AccentBorderBrush"] = "#2C6B62",
        ["AccentTextBrush"] = "#DFFFF9",
        ["AmberBrush"] = "#F0B54A",
        ["DangerBrush"] = "#F4695F",
        ["OperatorBrush"] = "#47D1C4",
        ["SpeakerBrush"] = "#F0B54A",
        ["ButtonBrush"] = "#2A313A",
        ["ButtonHoverBrush"] = "#323A45",
        ["InputBrush"] = "#1B2027",
        ["PreviewBrush"] = "#07090C",
        ["PreviewPanelBrush"] = "#12161B",
        ["LogBrush"] = "#0A0C0F",
        ["LogTextBrush"] = "#B9C2CE",
        ["BadgeBrush"] = "#D9101317",
        ["MenuHoverBrush"] = "#35404C",
        ["DisabledTextBrush"] = "#7C8796",
    };

    // 白色背景：暖白紙感、深墨文字、深金重點色；預覽區留白，投影內容仍完整顯示。
    private static readonly Dictionary<string, string> LightPalette = new()
    {
        ["WindowBrush"] = "#F7F5EF",
        ["PanelBrush"] = "#FFFFFF",
        ["Panel2Brush"] = "#F1EEE4",
        ["Panel3Brush"] = "#E8E3D6",
        ["LineBrush"] = "#CFC5AE",
        ["TextPrimaryBrush"] = "#1C1810",
        ["TextDimBrush"] = "#4C4232",
        ["TextFaintBrush"] = "#746854",
        ["TealBrush"] = "#9A6A12",
        ["AccentSoftBrush"] = "#F3E4C2",
        ["AccentBorderBrush"] = "#B98A2E",
        ["AccentTextBrush"] = "#4A3208",
        ["AmberBrush"] = "#8F6210",
        ["DangerBrush"] = "#B3261E",
        ["OperatorBrush"] = "#0E7C6B",
        ["SpeakerBrush"] = "#A85B00",
        ["ButtonBrush"] = "#EFEAE0",
        ["ButtonHoverBrush"] = "#E3DBC9",
        ["InputBrush"] = "#FFFFFF",
        ["PreviewBrush"] = "#FFFFFF",
        ["PreviewPanelBrush"] = "#F4F1E8",
        ["LogBrush"] = "#F1EEE4",
        ["LogTextBrush"] = "#4C4232",
        ["BadgeBrush"] = "#CCFFFFFF",
        ["MenuHoverBrush"] = "#F3E4C2",
        ["DisabledTextBrush"] = "#A39A88",
    };
}
