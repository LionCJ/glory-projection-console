using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using ProjectionManager.Services;

namespace ProjectionManager;

/// <summary>
/// 經文投影視窗：無邊框、置頂、不搶焦點，依字級與視窗大小自動拆頁。
/// 拆頁用 TextBlock 實際量高，不用字數估算，換字級或換投影螢幕都會重新排。
/// </summary>
public partial class ScriptureWindow : Window
{
    private readonly List<List<VerseLine>> _pages = new();
    private readonly List<VerseLine> _verses = new();
    private ScripturePassage? _passage;
    private int _index;
    private bool _paginating;

    public ScriptureWindow()
    {
        InitializeComponent();
        SizeChanged += (_, _) =>
        {
            if (_passage is not null && IsVisible && !_paginating)
            {
                Paginate();
                SetPage(Math.Min(_index, Math.Max(0, _pages.Count - 1)));
            }
        };
    }

    public int PageCount => _pages.Count;
    public int CurrentIndex => _index;


    public string CurrentPagePlain =>
        _pages.Count == 0 || _index < 0 || _index >= _pages.Count
            ? string.Empty
            : string.Join("\n", _pages[_index].Select(v => $"{v.Number}  {v.Text}"));

    /// <summary>取指定頁的純文字（供控制台右側「下一頁」預覽使用），索引超出範圍回傳空字串。</summary>
    public string GetPagePlain(int index)
    {
        return _pages.Count == 0 || index < 0 || index >= _pages.Count
            ? string.Empty
            : string.Join("\n", _pages[index].Select(v => $"{v.Number}  {v.Text}"));
    }

    public void ShowPassage(MonitorInfo target, ScripturePassage passage, List<VerseLine> verses)
    {
        _passage = passage;
        _verses.Clear();
        _verses.AddRange(verses);
        _index = 0;

        ApplyTheme();

        if (!IsVisible)
        {
            Left = target.Left;
            Top = target.Top;
            Width = 640;
            Height = 360;
            Show();
        }

        IntPtr hwnd = new WindowInteropHelper(this).Handle;
        Native.SetNoActivate(hwnd, true);
        Native.SetTopmost(hwnd, true);
        Native.SetWindowPos(hwnd, Native.HWND_TOPMOST,
            target.Left, target.Top, target.Width, target.Height,
            Native.SWP_NOACTIVATE | Native.SWP_SHOWWINDOW);

        UpdateLayout();
        Paginate();
        SetPage(0);
    }

    public void SetPage(int index)
    {
        if (_pages.Count == 0)
        {
            LinesHost.Children.Clear();
            return;
        }

        _index = Math.Clamp(index, 0, _pages.Count - 1);
        RenderPage();
    }

    /// <summary>改字級或主題後重新排版，盡量停在原來那一節。</summary>
    public void ReapplyStyle(ScripturePassage? passage = null)
    {
        if (passage is not null)
        {
            _passage = passage;
        }

        if (_passage is null)
        {
            return;
        }

        ApplyTheme();
        Paginate();
        SetPage(Math.Min(_index, Math.Max(0, _pages.Count - 1)));
    }

    public void KeepOnTop()
    {
        if (IsVisible)
        {
            Native.SetTopmost(new WindowInteropHelper(this).Handle, true);
        }
    }

    public void HideWindow() => Hide();

    private void ApplyTheme()
    {
        bool dark = _passage?.Theme == 0;
        RootGrid.Background = new SolidColorBrush(dark
            ? Color.FromRgb(0x0B, 0x12, 0x20)
            : Color.FromRgb(0xF6, 0xF2, 0xE8));
        RefText.Foreground = new SolidColorBrush(dark
            ? Color.FromRgb(0x8F, 0xB3, 0xD9)
            : Color.FromRgb(0x6B, 0x5A, 0x3E));
        PageText.Foreground = new SolidColorBrush(dark
            ? Color.FromRgb(0x5C, 0x6B, 0x7E)
            : Color.FromRgb(0x8C, 0x84, 0x74));
    }

    private double ContentWidth => Math.Max(220, ActualWidth - ActualWidth * 0.13);
    private double ContentHeight => Math.Max(200, ActualHeight - 230);

    private void Paginate()
    {
        if (_passage is null || _verses.Count == 0)
        {
            _pages.Clear();
            return;
        }

        _paginating = true;
        try
        {
            _pages.Clear();
            double width = ContentWidth;
            double avail = ContentHeight;
            double fs = _passage.FontSize;
            double gap = fs * 0.42;

            var page = new List<VerseLine>();
            double used = 0;

            void Flush()
            {
                if (page.Count > 0)
                {
                    _pages.Add(new List<VerseLine>(page));
                    page.Clear();
                }

                used = 0;
            }

            foreach (var verse in _verses)
            {
                string full = _passage.ShowNumbers ? $"{verse.Number}　{verse.Text}" : verse.Text;
                foreach (var chunk in SplitToFit(full, width, fs, avail))
                {
                    double h = MeasureHeight(chunk, width, fs);
                    if (page.Count > 0 && used + h > avail)
                    {
                        Flush();
                    }

                    page.Add(new VerseLine { Number = verse.Number, Text = chunk });
                    used += h + gap;
                }
            }

            Flush();
        }
        finally
        {
            _paginating = false;
        }
    }

    /// <summary>一節比整頁還高時，按字數切塊；其餘情況原樣一節一塊。</summary>
    private static List<string> SplitToFit(string text, double width, double fontSize, double avail)
    {
        double h = MeasureHeight(text, width, fontSize);
        if (h <= avail)
        {
            return new List<string> { text };
        }

        var result = new List<string>();
        int start = 0;

        while (start < text.Length)
        {
            int length = Math.Min(text.Length - start, 40);
            while (length > 1 && MeasureHeight(text.Substring(start, length), width, fontSize) > avail)
            {
                length = Math.Max(1, length - 1);
            }

            result.Add(text.Substring(start, length));
            start += length;
        }

        return result;
    }

    private static double MeasureHeight(string text, double width, double fontSize)
    {
        var block = new TextBlock
        {
            Text = text,
            FontFamily = new FontFamily("Microsoft JhengHei"),
            FontSize = fontSize,
            TextWrapping = TextWrapping.Wrap,
            Width = width,
            LineHeight = fontSize * 1.5,
            LineStackingStrategy = LineStackingStrategy.BlockLineHeight,
        };

        block.Measure(new Size(width, double.PositiveInfinity));
        return block.DesiredSize.Height;
    }

    private void RenderPage()
    {
        LinesHost.Children.Clear();

        if (_passage is null || _pages.Count == 0)
        {
            return;
        }

        bool dark = _passage.Theme == 0;
        var brush = new SolidColorBrush(dark
            ? Color.FromRgb(0xF2, 0xF5, 0xF8)
            : Color.FromRgb(0x18, 0x16, 0x12));

        RefText.Text = _passage.Reference;

        foreach (var line in _pages[_index])
        {
            LinesHost.Children.Add(new TextBlock
            {
                Text = line.Text,
                FontFamily = new FontFamily("Microsoft JhengHei"),
                FontSize = _passage.FontSize,
                Foreground = brush,
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Justify,
                LineHeight = _passage.FontSize * 1.5,
                LineStackingStrategy = LineStackingStrategy.BlockLineHeight,
                Margin = new Thickness(0, 0, 0, _passage.FontSize * 0.42),
            });
        }

        PageText.Text = _pages.Count > 1 ? $"第 {_index + 1} / {_pages.Count} 頁" : string.Empty;
    }
}
