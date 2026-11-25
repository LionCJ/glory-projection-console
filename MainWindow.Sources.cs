using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using ProjectionManager.Services;

namespace ProjectionManager;

/// <summary>
/// V1.2–V1.4 的額外投影源：經文、網頁、Zoom。
/// 全部沿用 V1.1 的三欄控制台：卡片照樣出現在左欄，翻頁照樣交給同一個 Navigate。
/// </summary>
public partial class MainWindow
{
    private readonly List<ScripturePassage> _scriptures = new();
    private readonly List<WebLink> _webs = new();
    private readonly List<SourceItem> _scriptureSources = new();
    private readonly List<SourceItem> _webSources = new();
    private readonly SourceItem _zoomSource = new() { Kind = "zoom", Name = "Zoom 會議投影" };

    private ScriptureWindow? _scriptureWindow;
    private WebWindow? _webWindow;
    private readonly ZoomService _zoomService = new();

    // ------------------------------------------------------------ 單螢幕診斷

    /// <summary>
    /// 只有一台螢幕時，正式投影流程不會開視窗（會蓋住控制台），
    /// 所以留兩個診斷參數，用來先在主螢幕驗證視窗本身：
    ///   ProjectionManager.exe --test-web [url]
    ///   ProjectionManager.exe --test-scripture
    /// </summary>
    private void TryRunDiagnostics(string[] args)
    {
        if (args.Length < 2 || _monitors.Count == 0)
        {
            return;
        }

        var primary = _monitors[0];

        try
        {
            switch (args[1].ToLowerInvariant())
            {
                case "--test-web":
                {
                    string url = args.Length > 2 ? args[2] : "https://example.com/";
                    _webWindow ??= new WebWindow();
                    _webWindow.ShowUrl(primary, url);
                    Log($"診斷：已在主螢幕開啟網頁投影視窗（{url}）。");
                    break;
                }

                case "--test-scripture":
                {
                    var passage = new ScripturePassage
                    {
                        Book = "約翰福音",
                        Chapter = 3,
                        StartVerse = 16,
                        EndVerse = 21,
                        FontSize = 48,
                        Theme = 0,
                    };

                    var verses = ScriptureStore.ReadPassage(passage);
                    _scriptureWindow ??= new ScriptureWindow();
                    _scriptureWindow.ShowPassage(primary, passage, verses);
                    Log($"診斷：已在主螢幕開啟經文投影視窗（拆成 {_scriptureWindow.PageCount} 頁）。");
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            Log("診斷失敗：" + ex.Message);
        }
    }

    // ------------------------------------------------------------ 載入與保存

    private void LoadSupplementalSources()
    {
        _scriptures.AddRange(ScriptureStore.LoadPassages());
        _webs.AddRange(ScriptureStore.LoadWebLinks());
        RebuildSupplementalSources();
    }

    private void RebuildSupplementalSources()
    {
        _scriptureSources.Clear();
        foreach (var passage in _scriptures)
        {
            _scriptureSources.Add(new SourceItem
            {
                Kind = "scripture",
                Name = passage.Reference,
                Scripture = passage,
            });
        }

        _webSources.Clear();
        foreach (var link in _webs)
        {
            _webSources.Add(new SourceItem
            {
                Kind = "web",
                Name = link.Name,
                Web = link,
            });
        }
    }

    private void SaveSupplemental()
    {
        ScriptureStore.SavePassages(_scriptures);
        ScriptureStore.SaveWebLinks(_webs);
    }

    // ------------------------------------------------------------ 切換來源

    /// <summary>
    /// 回傳 true 表示這個來源由本檔處理（經文／網頁／Zoom）；
    /// 回傳 false 就交回原本的主 PPT／素材流程。
    /// </summary>
    private bool BeforeActivate(SourceItem source)
    {
        // 單純選取投影源時不碰螢幕 2；上一份投影會一直保留，直到按「推送投影」才切換。
        if (source.Kind is "scripture" or "web" or "zoom")
        {
            _selected = source;

            switch (source.Kind)
            {
                case "scripture":
                    ActivateScripture(source);
                    return true;
                case "web":
                    ActivateWeb(source);
                    return true;
                default:
                    ActivateZoom(source);
                    return true;
            }
        }

        return false;
    }

    private void ActivateScripture(SourceItem source)
    {
        if (_target is null)
        {
            Log("尚未選擇投影螢幕，無法預覽經文。");
            return;
        }

        var passage = source.Scripture;
        if (passage is null)
        {
            return;
        }

        try
        {
            source.Verses = ScriptureStore.ReadPassage(passage);
        }
        catch (Exception ex)
        {
            Log("讀取經文失敗：" + ex.Message);
            return;
        }

        _live = source;
        _selected = source;

        if (_projectedSource != source)
        {
            source.PageIndex = 0;
        }

        Log(IsProjected(source)
            ? $"已選擇經文「{passage.Reference}」；此來源正在投影，按「斷開投影」可停止。"
            : $"已選擇經文「{passage.Reference}」；尚未投影，請按「推送投影」送到螢幕 2。");
        RenderAll();
    }

    private void ActivateWeb(SourceItem source)
    {
        if (_target is null)
        {
            Log("尚未選擇投影螢幕，無法預覽網頁。");
            return;
        }

        var link = source.Web;
        if (link is null)
        {
            return;
        }

        _live = source;
        _selected = source;

        Log(IsProjected(source)
            ? $"已選擇網頁「{link.Name}」；此來源正在投影，按「斷開投影」可停止。"
            : $"已選擇網頁「{link.Name}」；尚未投影，請按「推送投影」送到螢幕 2。");
        RenderAll();
    }

    private void ActivateZoom(SourceItem source)
    {
        if (_target is null)
        {
            Log("尚未選擇投影螢幕，無法使用 Zoom 投影。");
            return;
        }

        _live = source;
        _selected = source;

        Log(IsProjected(source)
            ? "已選擇 Zoom 投影；此來源正在投影，按「斷開投影」可停止。"
            : "已選擇 Zoom 投影；按「推送投影」把 Zoom 會議視窗鋪到螢幕 2。");
        RenderAll();
    }

    /// <summary>經文、網頁、Zoom 的通用推送：與素材一致，按下才真的上螢幕 2。</summary>
    private void PushExtraProjection()
    {
        if (_live is null || _target is null)
        {
            Log("尚未選擇投影源或投影螢幕。");
            return;
        }

        if (IsProjected(_live))
        {
            DisconnectProjection();
            return;
        }

        switch (_live.Kind)
        {
            case "scripture":
            {
                if (_live.Verses is null && _live.Scripture is not null)
                {
                    try
                    {
                        _live.Verses = ScriptureStore.ReadPassage(_live.Scripture);
                    }
                    catch (Exception ex)
                    {
                        Log("讀取經文失敗：" + ex.Message);
                        return;
                    }
                }

                StopProjectionForSwitch(_live);
                _projectedSource = _live;
                _black = false;
                _extraProjectionActive = true;

                if (SingleScreen)
                {
                    _scriptureWindow?.HideWindow();
                    Log("目前只有一台螢幕：經文只在控制台預覽；接上投影幕後才會送到螢幕 2。");
                }
                else
                {
                    ShowScriptureWindow(_live.Scripture!, _live.Verses);
                    _scriptureWindow?.SetPage(_live.PageIndex);
                    Log($"已推送經文「{_live.Scripture?.Reference}」（{_scriptureWindow?.PageCount ?? 0} 頁）到投影螢幕。");
                }

                RenderAll();
                break;
            }

            case "web":
            {
                StopProjectionForSwitch(_live);
                _projectedSource = _live;
                _black = false;
                _extraProjectionActive = true;

                if (SingleScreen)
                {
                    _webWindow?.HideWindow();
                    Log("目前只有一台螢幕：網頁只在控制台顯示網址；接上投影幕後才會送到螢幕 2。");
                }
                else
                {
                    ShowWebWindow(_live.Web!);
                    Log($"已推送網頁「{_live.Web?.Name}」到投影螢幕：{_live.Web?.Url}");
                }

                RenderAll();
                break;
            }

            case "zoom":
            {
                StopProjectionForSwitch(_live);
                string message;
                bool ok = _zoomService.Activate(_target, out message);
                Log(message);
                _zoomProjectionActive = ok;
                _projectedSource = ok ? _live : null;
                _extraProjectionActive = false;

                if (ok)
                {
                    _black = false;
                    Log("提醒：Zoom 的「分享螢幕 2」要在 Zoom 裡自己選；本程式只負責把會議視窗鋪到投影幕。");
                }

                RenderAll();
                break;
            }
        }
    }

    private void ShowScriptureWindow(ScripturePassage passage, List<VerseLine>? verses)
    {
        if (_target is null || verses is null)
        {
            return;
        }

        _scriptureWindow ??= new ScriptureWindow();
        _scriptureWindow.ShowPassage(_target, passage, verses);
    }

    private void ShowWebWindow(WebLink link)
    {
        if (_target is null)
        {
            return;
        }

        _webWindow ??= new WebWindow();
        _webWindow.ShowUrl(_target, link.Url);
    }

    // ------------------------------------------------------------ 翻頁

    private void FlipScripture(bool forward)
    {
        if (_live?.Scripture is null)
        {
            return;
        }

        if (!IsProjected(_live))
        {
            Log("經文尚未推送投影；請先按「推送投影」，再翻頁。");
            return;
        }

        var window = _scriptureWindow;
        if (window is null || !window.IsVisible || window.PageCount == 0)
        {
            Log("單螢幕：經文沒有投影視窗，只在控制台預覽；接上第二螢幕後才可翻頁。");
            if (!SingleScreen)
            {
                ShowScriptureWindow(_live.Scripture, _live.Verses);
            }

            return;
        }

        int next = _live.PageIndex + (forward ? 1 : -1);
        if (next < 0)
        {
            Log("經文已在第一頁。");
            return;
        }

        if (next >= window.PageCount)
        {
            Log($"經文已在最後一頁（共 {window.PageCount} 頁）。");
            return;
        }

        window.SetPage(next);
        _live.PageIndex = next;
        _black = false;
        Log($"經文 → 第 {next + 1} / {window.PageCount} 頁。");
        RenderPreview();
        RenderThumbs();
        UpdateShowStatus();
    }

    private void FlipWeb(bool forward)
    {
        if (!IsProjected(_live))
        {
            Log("網頁尚未推送投影；請先按「推送投影」，再滾動網頁。");
            return;
        }

        if (_webWindow is null || !_webWindow.IsVisible)
        {
            Log("網頁沒有投影視窗（單螢幕或尚未投出），無法滾動。");
            return;
        }

        _webWindow.Scroll(forward);
        Log(forward ? "網頁：向下滾動一屏。" : "網頁：向上滾動一屏。");
    }

    private void RestoreExtraProjection()
    {
        var source = _projectedSource;
        if (source?.Kind == "scripture" && source.Scripture is not null && source.Verses is not null && IsProjected(source))
        {
            ShowScriptureWindow(source.Scripture, source.Verses);
            _scriptureWindow?.SetPage(source.PageIndex);
            Log("已恢復經文畫面。");
            return;
        }

        if (source?.Kind == "web" && source.Web is not null && IsProjected(source))
        {
            ShowWebWindow(source.Web);
            Log("已恢復網頁畫面。");
        }
    }

    // ------------------------------------------------------------ 預覽與縮圖

    private void RenderExtraPreview()
    {
        if (_live is null)
        {
            return;
        }

        PreviewLiveImage.Source = null;
        PreviewNextImage.Source = null;
        PreviewNextHint.Text = "—";
        PreviewNextHint.Visibility = Visibility.Visible;
        PreviewNextText.Text = string.Empty;
        PreviewNextText.Visibility = Visibility.Collapsed;
        PreviewNextName.Text = "—";
        PreviewLiveHint.Visibility = Visibility.Collapsed;
        PreviewLiveText.Visibility = Visibility.Visible;

        switch (_live.Kind)
        {
            case "scripture":
            {
                var passage = _live.Scripture;
                var window = _scriptureWindow;
                bool pushed = IsProjected(_live);
                int pageCount = pushed && window?.IsVisible == true ? window.PageCount : 0;
                _live.PageIndex = pushed && window?.IsVisible == true ? window.CurrentIndex : _live.PageIndex;

                PreviewLiveName.Text = passage is null
                    ? "經文"
                    : $"經文 · {passage.Reference}"
                      + (pageCount > 0 ? $"　{_live.PageIndex + 1}/{pageCount}" : "　待推送");

                string text = pushed && window?.IsVisible == true
                    ? window.CurrentPagePlain
                    : (_live.Verses ?? new List<VerseLine>())
                        .Select(v => $"{v.Number}　{v.Text}")
                        .Aggregate(string.Empty, (a, b) => a.Length == 0 ? b : a + "\n" + b);

                if (!pushed)
                {
                    text = "（尚未推送投影）按中間「推送投影」送到螢幕 2。\n\n" + text;
                }

                PreviewLiveText.Text = Trim(text, 700);
                PreviewLiveText.FontSize = 15;

                // 多頁經文：右側顯示真正的「下一頁」內容，翻頁筆才有可預期性。
                if (pushed && window?.IsVisible == true && pageCount > 0)
                {
                    int nextIndex = _live.PageIndex + 1;
                    if (nextIndex < pageCount)
                    {
                        PreviewNextName.Text = $"下一頁　第 {nextIndex + 1} / {pageCount} 頁";
                        PreviewNextText.Text = Trim(window.GetPagePlain(nextIndex), 700);
                        PreviewNextText.FontSize = 15;
                        PreviewNextText.Visibility = Visibility.Visible;
                        PreviewNextHint.Visibility = Visibility.Collapsed;
                    }
                    else
                    {
                        PreviewNextName.Text = "下一頁";
                        PreviewNextHint.Text = pageCount <= 1
                            ? "此段經文只有 1 頁"
                            : "已是最後一頁";
                    }
                }
                else
                {
                    PreviewNextName.Text = "下一頁";
                    PreviewNextHint.Text = "尚未推送投影\n推送後可預覽下一頁";
                }
                break;
            }

            case "web":
            {
                bool pushed = IsProjected(_live);
                PreviewLiveName.Text = $"網頁 · {_live.Web?.Name}" + (pushed ? "" : " · 待推送");
                PreviewLiveText.Text = (_live.Web?.Url ?? string.Empty)
                                       + (pushed
                                           ? "\n\n（投影幕上顯示的是這個網頁；控制台可重新載入、上一頁／下一頁或上下滾動）"
                                           : "\n\n（尚未推送投影；按中間「推送投影」才會送到螢幕 2）");
                PreviewLiveText.FontSize = 14;
                break;
            }

            default:
            {
                bool pushed = IsProjected(_live);
                PreviewLiveName.Text = "Zoom · 會議投影";
                if (pushed)
                {
                    PreviewLiveText.Text = "Zoom 會議視窗已鋪在投影螢幕上（無邊框、置頂）。\n\n"
                                           + "流程：Zoom 分享螢幕 2 → 把主持權交給線上講員 → 取消螢幕 2 分享 → 按「推送投影」把 Zoom 鋪到螢幕 2。\n\n"
                                           + "按「斷開投影」會把 Zoom 還原成普通視窗。";
                }
                else
                {
                    PreviewLiveText.Text = $"尚未推送投影。\n\n"
                                           + $"{_zoomService.StatusText}\n\n"
                                           + "按中間「推送投影」把 Zoom 會議視窗鋪到投影螢幕；右欄可直接開啟 Zoom、重新抓取或還原視窗。";
                }

                PreviewLiveText.FontSize = 13;
                break;
            }
        }
    }

    private void RenderExtraThumbs()
    {
        if (_live is null)
        {
            return;
        }

        if (_live.Kind == "scripture")
        {
            var window = _scriptureWindow;
            bool pushed = IsProjected(_live);
            int count = pushed && window?.IsVisible == true ? window.PageCount : 0;
            if (count == 0)
            {
                AddThumbTile(pushed ? "單螢幕" : "尚未推送", null, false, () => { });
                return;
            }

            for (int i = 0; i < count; i++)
            {
                int index = i;
                AddScriptureThumbTile(i + 1, window!.GetPagePlain(i), index == window.CurrentIndex, () =>
                {
                    window.SetPage(index);
                    _live!.PageIndex = index;
                    RenderPreview();
                    RenderThumbs();
                    UpdateShowStatus();
                });
            }

            return;
        }

        if (_live.Kind == "web")
        {
            AddThumbTile("網頁", null, true, () => { });
            return;
        }

        AddThumbTile("Zoom", null, true, () => { });
    }

    private static string Trim(string text, int max)
    {
        if (text.Length <= max)
        {
            return text;
        }

        return text[..max] + "…";
    }

    private string DescribeLiveSource()
    {
        var source = _projectedSource;
        if (source is null)
        {
            return "螢幕 2：未投影";
        }

        return "螢幕 2：" + DescribeSourceName(source);
    }

    // ------------------------------------------------------------ 按鈕

    private void ScriptureButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            ScriptureStore.EnsureLoaded();
        }
        catch (Exception ex)
        {
            Log(ex.Message);
            MessageBox.Show(this, ex.Message, "找不到聖經資料", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dialog = new ScriptureDialog(_scriptures.FirstOrDefault()) { Owner = this };
        if (dialog.ShowDialog() != true || dialog.Result is null)
        {
            return;
        }

        var result = dialog.Result;
        var existing = _scriptures.FirstOrDefault(s => s.Key == result.Key);
        if (existing is not null)
        {
            existing.FontSize = result.FontSize;
            existing.Theme = result.Theme;
            existing.ShowNumbers = result.ShowNumbers;
            result = existing;
            Log($"已更新既有經文「{result.Reference}」的字級／樣式。");
        }
        else
        {
            _scriptures.Add(result);
            Log($"已加入經文「{result.Reference}」。");
        }

        SaveSupplemental();
        RebuildSupplementalSources();

        var source = _scriptureSources.First(s => s.Scripture == result);
        ActivateSource(source);
    }

    private void WebButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new WebDialog { Owner = this };
        if (dialog.ShowDialog() != true || dialog.Result is null)
        {
            return;
        }

        _webs.Add(dialog.Result);
        SaveSupplemental();
        RebuildSupplementalSources();
        Log($"已加入網頁「{dialog.Result.Name}」：{dialog.Result.Url}");

        var source = _webSources.First(s => s.Web == dialog.Result);
        ActivateSource(source);
    }

    private void ZoomButton_Click(object sender, RoutedEventArgs e) => ActivateSource(_zoomSource);

    private void ZoomOpen_Click(object sender, RoutedEventArgs e)
    {
        string message;
        _zoomService.Launch(out message);
        Log(message);
        _zoomService.Refresh();
        UpdateZoomPanel();
    }

    private void ZoomRefresh_Click(object sender, RoutedEventArgs e)
    {
        _zoomService.Refresh();
        Log(_zoomService.StatusText);
        UpdateZoomPanel();
    }

    private void ZoomPlace_Click(object sender, RoutedEventArgs e)
    {
        if (_target is null)
        {
            Log("尚未選擇投影螢幕，無法鋪 Zoom。");
            return;
        }

        if (_live?.Kind != "zoom")
        {
            _selected = _zoomSource;
            _live = _zoomSource;
        }

        PushExtraProjection();
    }

    private void ZoomRestore_Click(object sender, RoutedEventArgs e)
    {
        _zoomService.Restore();
        _zoomProjectionActive = false;
        Log("Zoom 已還原成普通視窗。");
        UpdateZoomPanel();
        RenderAll();
    }

    private void UpdateZoomPanel()
    {
        if (ZoomStatusText is null)
        {
            return;
        }

        _zoomService.Refresh();
        ZoomStatusText.Text = _zoomService.IsActive
            ? "投影中：會議視窗已鋪在螢幕 2（無邊框、置頂）"
            : _zoomService.IsRunning
                ? $"Zoom 執行中 · 可見視窗 {_zoomService.WindowCount} 個"
                : "Zoom 未執行";
        ZoomWindowText.Text = string.IsNullOrEmpty(_zoomService.WindowTitle)
            ? "尚未取得會議視窗標題"
            : _zoomService.WindowTitle;
        ZoomStatusText.Foreground = _zoomService.IsActive
            ? ThemeManager.Brush("TealBrush")
            : ThemeManager.Brush("TextDimBrush");
    }

    private void WebReload_Click(object sender, RoutedEventArgs e)
    {
        _webWindow?.Reload();
        Log("網頁：已重新載入。");
    }

    private void WebBack_Click(object sender, RoutedEventArgs e)
    {
        _webWindow?.Back();
        Log("網頁：上一頁。");
    }

    private void WebForward_Click(object sender, RoutedEventArgs e)
    {
        _webWindow?.Forward();
        Log("網頁：下一頁。");
    }

    // ------------------------------------------------------------ 右鍵選單

    private ContextMenu? BuildExtraMenu(SourceItem source)
    {
        if (source.Kind == "zoom")
        {
            var restore = new MenuItem { Header = "還原 Zoom 視窗" };
            restore.Click += (_, _) =>
            {
                _zoomService.Restore();
                Log("Zoom 已還原成普通視窗。");
                RenderAll();
            };

            var zoomMenu = new ContextMenu();
            zoomMenu.Items.Add(restore);
            return zoomMenu;
        }

        if (source.Kind is not ("scripture" or "web"))
        {
            return null;
        }

        var menu = new ContextMenu();

        if (source.Kind == "scripture")
        {
            var edit = new MenuItem { Header = "修改經文（書卷／章節／字級）" };
            edit.Click += (_, _) => EditScripture(source);
            menu.Items.Add(edit);
        }

        var remove = new MenuItem { Header = "從左欄移除" };
        remove.Click += (_, _) => RemoveExtraSource(source);
        menu.Items.Add(remove);

        return menu;
    }

    /// <summary>修改既有經文：改書卷、章節範圍、字級、樣式，存檔後就地更新左欄卡片。</summary>
    private void EditScripture(SourceItem source)
    {
        if (source.Scripture is null)
        {
            return;
        }

        var passage = source.Scripture;
        var dialog = new ScriptureDialog(passage) { Owner = this };
        if (dialog.ShowDialog() != true || dialog.Result is null)
        {
            return;
        }

        var result = dialog.Result;

        passage.Book = result.Book;
        passage.Chapter = result.Chapter;
        passage.StartVerse = result.StartVerse;
        passage.EndVerse = result.EndVerse;
        passage.FontSize = result.FontSize;
        passage.Theme = result.Theme;
        passage.ShowNumbers = result.ShowNumbers;

        // 改完後若和另一條已存在的經文重複，就合併成這一條
        var duplicate = _scriptures.FirstOrDefault(s => s.Id != passage.Id && s.Key == passage.Key);
        if (duplicate is not null)
        {
            _scriptures.Remove(duplicate);
        }

        SaveSupplemental();
        bool wasLive = _live == source;
        RebuildSupplementalSources();

        var updated = _scriptureSources.FirstOrDefault(s => s.Scripture == passage);
        if (wasLive && updated is not null)
        {
            ActivateScripture(updated);
        }

        Log($"已修改經文為「{passage.Reference}」，字級 {passage.FontSize} pt。");
        RenderAll();
    }

    private void RemoveExtraSource(SourceItem source)
    {
        if (_live == source)
        {
            _scriptureWindow?.HideWindow();
            _webWindow?.HideWindow();
            _live = null;
            _selected = null;
        }

        if (source.Kind == "scripture" && source.Scripture is not null)
        {
            _scriptures.RemoveAll(s => s.Key == source.Scripture!.Key);
        }
        else if (source.Kind == "web" && source.Web is not null)
        {
            _webs.RemoveAll(w => w.Id == source.Web!.Id);
        }

        SaveSupplemental();
        RebuildSupplementalSources();
        Log($"已移除「{source.Name}」。");
        RenderAll();
    }
}
