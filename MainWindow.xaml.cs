using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using ProjectionManager.Services;

namespace ProjectionManager;

public partial class MainWindow : Window
{
    /// <summary>控制權歸屬。只作標示與記錄，不做互斥鎖定。</summary>
    private enum ControlMode
    {
        Operator,
        Speaker,
    }

    /// <summary>投影源：主 PPT（由 PowerPoint 原生放映）或一份已導入的素材。</summary>
    private sealed class SourceItem
    {
        public string Kind = "main";           // main | material | scripture | web | zoom
        public string Name = string.Empty;
        public MaterialItem? Material;
        public int PageIndex;                  // 素材／經文為 0 起算
        public ScripturePassage? Scripture;
        public WebLink? Web;
        public List<VerseLine>? Verses;
    }

    // 主題切換時這些 Brush 會由 ThemeManager 就地換色，既有卡片與文字會跟著更新。
    private static Brush OperatorBrush => ThemeManager.Brush("OperatorBrush");
    private static Brush SpeakerBrush => ThemeManager.Brush("SpeakerBrush");
    private static Brush CardBrush => ThemeManager.Brush("Panel2Brush");
    private static Brush CardLiveBrush => ThemeManager.Brush("AccentSoftBrush");
    private static Brush LineBrush => ThemeManager.Brush("LineBrush");
    private static Brush DimBrush => ThemeManager.Brush("TextDimBrush");
    private static Brush FaintBrush => ThemeManager.Brush("TextFaintBrush");
    private static Brush DangerBrush => ThemeManager.Brush("DangerBrush");

    private readonly PowerPointService _ppt = new();
    private readonly List<MonitorInfo> _monitors = new();
    private readonly DispatcherTimer _statusTimer = new();
    private readonly DispatcherTimer _bindTimer = new();
    private readonly List<MaterialItem> _materials = new();
    private readonly List<MaterialGroup> _groups = new();
    private readonly List<SourceItem> _sources = new();
    private readonly List<Border> _thumbBorders = new();
    private double _thumbTileW = 160;
    private double _thumbTileH = 90;
    private System.Windows.Threading.DispatcherTimer? _thumbResizeTimer;
    private bool _thumbPrefetchStarted;
    private int _thumbCols = 2;

    private AppSettings _settings = new();
    private RawInputService? _rawInput;
    private KeyboardGuard? _guard;
    private MonitorInfo? _target;
    private MaterialWindow? _materialWindow;

    private SourceItem? _live;
    private SourceItem? _selected;          // 右欄操作的對象，可能與目前投影中的來源不同
    private SourceItem? _projectedSource;    // 真正在螢幕 2 上的來源；只有按推送／開始放映才會改變
    // 素材預設只進「待推送」預覽；按「推送投影」後才真的送到螢幕 2。
    private bool _materialProjectionActive;

    // 經文、網頁、Zoom 同樣改成「先選擇、再推送投影」，與素材語意一致。
    private bool _extraProjectionActive;
    private bool _zoomProjectionActive;
    private string _sourceSearchKeyword = string.Empty;

    private ControlMode _mode = ControlMode.Operator;
    private int _handoverSlide;
    private long _repositionUntil;
    private bool _binding;
    private int _bindReceived;
    private bool _black;
    private bool _ready;
    private int _lastMainPage;
    private int _lastMaterialPage = -1;
    private long _lastArrowNavAt = long.MinValue / 2;
    private bool _noActivateNoticeShown;
    private string _activeGroupId = MaterialGroup.DefaultId;   // 匯入素材時放進哪個分組
    private readonly HashSet<string> _collapsedGroups = new();  // 左欄收合中的分組（存到設定檔）
    private bool _importing;
    private CancellationTokenSource? _importCts;
    private ImportProgressDialog? _importDialog;

    // 添加簡報素材時如果主 PPT 正在放映，PowerPoint 不允許同一個實例再開別份簡報。
    // 匯入期間先暫停放映並用靜態畫面蓋住投影幕，完成後自動回到原本頁數。
    private bool _pptSuspendedForImport;
    private bool _presentationImportScope;
    private bool _pptShowWasRunningForImport;
    private bool _pptSuspendForImportWasProjected;
    private int _pptSuspendForImportSlide;
    private MaterialWindow? _importFreezeWindow;
    private string? _importFreezeImage;
    private bool _topmostBeforeImport;

    // 素材縮圖／預覽改成背景循序解碼：點卡片先出畫面，再補圖，不再讓 UI 執行緒卡在 BitmapImage。
    private readonly SemaphoreSlim _thumbLoadGate = new(3);
    private readonly SemaphoreSlim _previewLoadGate = new(2);
    private int _thumbLoadGeneration;
    private int _previewLoadGeneration;
    private string? _previewRequestedKey;
    private string? _nextPreviewRequestedKey;

    // 測試模式：用半個主螢幕模擬「螢幕 2」，可自動放映，供自動化測試跑完整流程
    private bool _testMode;
    private bool _autoShow;

    // 主 PPT 的靜態預覽圖（每次放映用一個新的暫存資料夾，避免同名檔被舊圖快取）
    private string? _previewDir;
    private int _previewPage = -1;
    private string? _previewFile;
    private bool _previewSlowNoted;

    // 預覽圖改成「先更新介面、再排隊匯出」：大幅跳頁時不再讓 UI 卡在 COM 匯出上
    private bool _previewExportRunning;
    private long _previewRetryAt;

    /// <summary>
    /// 本次執行「已加入」的主 PPT 路徑。預設是空的：開機不會自動掛上次那份簡報，
    /// 要由使用者按「加入主 PPT」或拖進來才會出現；移除後也回到空白。
    /// </summary>
    private string? _mainPath;

    public MainWindow()
    {
        _settings = SettingsStore.Load();
        ThemeManager.Apply(_settings.UiTheme);

        InitializeComponent();

        ThemeManager.ThemeChanged += (_, _) => Dispatcher.BeginInvoke(() =>
        {
            // ThemeManager 每次切換會建立新的 Brush；控制者名稱、chip 用的是程式碼指定色，
            // 必須重新套一次，否則會停留在舊主題的顏色。
            ApplyMode();
            RenderAll();
        });

        var startupArgs = Environment.GetCommandLineArgs();
        _testMode = startupArgs.Any(a => a.Equals("--test-mode", StringComparison.OrdinalIgnoreCase));
        _autoShow = startupArgs.Any(a => a.Equals("--auto-show", StringComparison.OrdinalIgnoreCase));

        WindowedShowCheck.IsChecked = _settings.WindowedShow;
        NoActivateCheck.IsChecked = _settings.NoActivateConsole;

        foreach (var id in _settings.CollapsedGroups ?? new List<string>())
        {
            _collapsedGroups.Add(id);
        }

        var library = MaterialStore.Load();
        _materials.AddRange(library.Items);
        _groups.AddRange(library.Groups);

        LoadMonitors();
        ApplyDebugBadge();
        LoadSupplementalSources();
        RebuildSources();
        StartTimers();
        ApplyMode();
        RenderAll();
        StartMobileControl();

        Log("控制台已啟動。目前控制權：操作員。");
        Log($"素材庫：{_materials.Count} 份（索引 {MaterialStore.ImageRoot}）。");

        if (_monitors.Count < 2)
        {
            Log("警告：僅偵測到一台螢幕。請在 Windows 顯示設定中選「延伸」，否則操作員與投影畫面無法分離。");
        }
        else if (_testMode || _settings.DebugMode)
        {
            Log("調試模式：已用右半螢幕模擬第二螢幕（僅在只有一台實體螢幕時啟用）。");
        }

        if (!string.IsNullOrEmpty(_settings.BoundDeviceName))
        {
            Log("已載入綁定的翻頁筆：" + Describe(_settings.BoundDeviceName!) + "（如需更換請重新綁定）。");
        }
        else
        {
            Log("目前未綁定翻頁筆：講員按筆不會有反應，只有控制台按鈕有效。要交給講員前請先綁定。");
        }

        // 以參數啟動：ProjectionManager.exe "D:\...\簡報.pptx"
        // 測試模式可加上 --test-mode / --auto-show，其餘參數不影響檔案判斷
        var args = Environment.GetCommandLineArgs();
        string? mainArg = args.Skip(1)
            .FirstOrDefault(a => !a.StartsWith("--", StringComparison.Ordinal) && File.Exists(a));

        if (mainArg is not null)
        {
            try
            {
                _mainPath = mainArg;
                _ppt.Open(mainArg);
                _settings.LastPresentation = mainArg;
                SettingsStore.Save(_settings);
                MainSource.Name = Path.GetFileNameWithoutExtension(mainArg);
                Log("已開啟簡報：" + mainArg);
            }
            catch (Exception ex)
            {
                Log(ex.Message);
            }
        }

        if (_testMode)
        {
            Log("測試模式：已啟用模擬螢幕 2，手機服務僅綁定 127.0.0.1。");
        }

        _ready = true;
        TryRunDiagnostics(args);
        RenderAll();

        if (_testMode && _autoShow)
        {
            Loaded += async (_, _) =>
            {
                await Task.Delay(800);
                await StartShow();
            };
        }
    }

    // ---------------------------------------------------------------- 投影源

    private SourceItem MainSource => _sources.First(s => s.Kind == "main");

    private void RebuildSources()
    {
        _sources.Clear();

        var main = new SourceItem
        {
            Kind = "main",
            Name = _mainPath is not null
                ? Path.GetFileNameWithoutExtension(_mainPath)
                : "＋ 加入主 PPT（點此選檔）",
        };
        _sources.Add(main);

        foreach (var group in _groups)
        {
            foreach (var material in _materials.Where(m => m.GroupId == group.Id))
            {
                _sources.Add(new SourceItem
                {
                    Kind = "material",
                    Name = material.Name,
                    Material = material,
                    PageIndex = 0,
                });
            }
        }

        var previous = _live;
        var projectedPrevious = _projectedSource;
        if (previous is null)
        {
            _live = _sources[0];
        }
        else
        {
            var match = FindEquivalentSource(previous);
            _live = match ?? _sources[0];
        }

        _selected = _live;
        _projectedSource = projectedPrevious is null ? null : FindEquivalentSource(projectedPrevious);
        if (_projectedSource is null)
        {
            _materialProjectionActive = false;
            _extraProjectionActive = false;
            _zoomProjectionActive = false;
        }

        SourceCountText.Text = $"{_materials.Count + 1} 項";
    }

    private SourceItem? FindEquivalentSource(SourceItem previous)
    {
        var candidates = _sources
            .Concat(_scriptureSources)
            .Concat(_webSources)
            .ToList();

        if (_zoomSource is not null)
        {
            candidates.Add(_zoomSource);
        }

        return candidates.FirstOrDefault(s => s.Kind == previous.Kind
                                              && s.Material?.Id == previous.Material?.Id
                                              && s.Scripture?.Key == previous.Scripture?.Key
                                              && s.Web?.Url == previous.Web?.Url);
    }

    private bool SourceMatchesKeyword(SourceItem source)
    {
        var keyword = _sourceSearchKeyword.Trim();
        if (keyword.Length == 0)
        {
            return true;
        }

        if (source.Material is null)
        {
            return false;
        }

        if (source.Material.Name.Contains(keyword, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var groupName = _groups.FirstOrDefault(g => g.Id == source.Material.GroupId)?.Name ?? string.Empty;
        return groupName.Contains(keyword, StringComparison.OrdinalIgnoreCase);
    }

    private void RenderSources()
    {
        SourcePanel.Children.Clear();

        var keyword = _sourceSearchKeyword.Trim();
        if (keyword.Length > 0)
        {
            var matches = _sources
                .Where(s => s.Kind == "material" && SourceMatchesKeyword(s))
                .ToList();

            foreach (var group in _groups)
            {
                var groupMatches = matches.Where(s => s.Material?.GroupId == group.Id).ToList();
                if (groupMatches.Count == 0)
                {
                    continue;
                }

                SourcePanel.Children.Add(BuildGroupHeader(group));
                foreach (var source in groupMatches)
                {
                    SourcePanel.Children.Add(BuildSourceCard(source, 0));
                }
            }

            if (matches.Count == 0)
            {
                SourcePanel.Children.Add(new TextBlock
                {
                    Text = $"找不到符合「{keyword}」的素材",
                    Foreground = FaintBrush,
                    FontSize = 13,
                    Margin = new Thickness(8, 18, 8, 8),
                    TextWrapping = TextWrapping.Wrap,
                });
            }

            SourceCountText.Text = $"找到 {matches.Count} 項";
            if (SourceSearchStatus is not null)
            {
                SourceSearchStatus.Text = matches.Count == 0
                    ? $"沒有符合「{keyword}」的素材"
                    : $"找到 {matches.Count} 項；點擊結果即可選取";
            }
            return;
        }

        var main = _sources.FirstOrDefault(s => s.Kind == "main");
        if (main is not null)
        {
            SourcePanel.Children.Add(BuildSourceCard(main, 0));
        }

        int index = 1;
        foreach (var group in _groups)
        {
            SourcePanel.Children.Add(BuildGroupHeader(group));

            if (_collapsedGroups.Contains(group.Id))
            {
                continue;   // 收合時只留組標題，卡片全部隱藏
            }

            foreach (var source in _sources.Where(s => s.Kind == "material"
                                                       && s.Material?.GroupId == group.Id))
            {
                SourcePanel.Children.Add(BuildSourceCard(source, index));
                index++;
            }
        }

        index = RenderSupplementalSources(index);
        SourceCountText.Text = _sources.Count.ToString();
    }

    /// <summary>經文、網頁、Zoom 三類來源接在素材分組後面，回傳下一個可用的卡片序號。</summary>
    private int RenderSupplementalSources(int index)
    {
        if (_scriptureSources.Count > 0)
        {
            SourcePanel.Children.Add(BuildSectionHeader("經文", SectionScriptureKey, _scriptureSources.Count));
            if (!_collapsedGroups.Contains(SectionScriptureKey))
            {
                foreach (var source in _scriptureSources)
                {
                    var card = BuildSourceCard(source, index++);
                    card.ContextMenu = BuildExtraMenu(source);
                    SourcePanel.Children.Add(card);
                }
            }
        }

        if (_webSources.Count > 0)
        {
            SourcePanel.Children.Add(BuildSectionHeader("網頁", SectionWebKey, _webSources.Count));
            if (!_collapsedGroups.Contains(SectionWebKey))
            {
                foreach (var source in _webSources)
                {
                    var card = BuildSourceCard(source, index++);
                    card.ContextMenu = BuildExtraMenu(source);
                    SourcePanel.Children.Add(card);
                }
            }
        }

        SourcePanel.Children.Add(BuildSectionLabel("Zoom"));
        var zoomCard = BuildSourceCard(_zoomSource, index++);
        zoomCard.ContextMenu = BuildExtraMenu(_zoomSource);
        SourcePanel.Children.Add(zoomCard);

        return index;
    }

    private const string SectionScriptureKey = "__section_scripture";
    private const string SectionWebKey = "__section_web";

    /// <summary>經文／網頁區塊標題：可點擊收合，狀態沿用 CollapsedGroups 一起持久化。</summary>
    private Border BuildSectionHeader(string name, string key, int count)
    {
        bool collapsed = _collapsedGroups.Contains(key);

        var chevron = new TextBlock
        {
            Text = collapsed ? "▸" : "▾",
            Foreground = DimBrush,
            FontSize = 12,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 6, 0),
        };

        var label = new TextBlock
        {
            Text = $"{name}　({count})",
            Foreground = DimBrush,
            FontSize = 12,
            FontWeight = FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
        };

        var row = new Grid();
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        row.Children.Add(chevron);
        row.Children.Add(label);
        Grid.SetColumn(label, 1);

        var header = new Border
        {
            Background = Brushes.Transparent,
            BorderBrush = LineBrush,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(8, 8, 8, 6),
            Margin = new Thickness(0, 12, 0, 6),
            Cursor = Cursors.Hand,
            Child = row,
            ToolTip = collapsed ? $"展開{name}" : $"收合{name}",
        };

        AutomationProperties.SetName(header, $"{name}（{(collapsed ? "已收合" : "已展開")}，{count} 項）");
        AutomationProperties.SetAutomationId(header, "sec_" + key.TrimStart('_'));
        header.MouseLeftButtonUp += (_, _) => ToggleSectionCollapsed(key, name);
        return header;
    }

    private void ToggleSectionCollapsed(string key, string name)
    {
        if (_collapsedGroups.Contains(key))
        {
            _collapsedGroups.Remove(key);
            Log($"已展開{name}區塊。");
        }
        else
        {
            _collapsedGroups.Add(key);
            Log($"已收合{name}區塊。");
        }

        _settings.CollapsedGroups = _collapsedGroups.ToList();
        SettingsStore.Save(_settings);
        RenderSources();
    }

    private Border BuildSectionLabel(string text)
    {
        var label = new TextBlock
        {
            Text = text,
            Foreground = DimBrush,
            FontSize = 12,
            FontWeight = FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
        };

        return new Border
        {
            Background = Brushes.Transparent,
            BorderBrush = LineBrush,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(8, 8, 8, 6),
            Margin = new Thickness(0, 12, 0, 6),
            Child = label,
        };
    }

    private Border BuildGroupHeader(MaterialGroup group)
    {
        bool active = group.Id == _activeGroupId;
        bool collapsed = _collapsedGroups.Contains(group.Id);
        int count = _materials.Count(m => m.GroupId == group.Id);

        var chevron = new Button
        {
            Content = collapsed ? "▸" : "▾",
            Style = (Style)FindResource("SectionHeaderButton"),
            Foreground = active ? OperatorBrush : DimBrush,
            FontSize = 13,
            Width = 24,
            Height = 22,
            Padding = new Thickness(0),
            Margin = new Thickness(0),
            MinWidth = 0,
            Focusable = false,
            Cursor = Cursors.Hand,
            ToolTip = collapsed ? "展開這一組" : "收合這一組",
            VerticalAlignment = VerticalAlignment.Center,
        };
        AutomationProperties.SetName(chevron, $"{(collapsed ? "展開" : "收合")}分組：{group.Name}");
        AutomationProperties.SetAutomationId(chevron, "tgl_" + group.Id);
        chevron.Click += (_, _) => ToggleGroupCollapsed(group);

        var name = new TextBlock
        {
            Text = group.Name,
            Foreground = active ? ThemeManager.Brush("AccentTextBrush") : DimBrush,
            FontSize = 13,
            FontWeight = active ? FontWeights.SemiBold : FontWeights.Normal,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(6, 0, 0, 0),
        };

        var countText = new TextBlock
        {
            Text = count.ToString(),
            Foreground = FaintBrush,
            FontSize = 12,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 0, 2, 0),
        };

        var row = new Grid();
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        Grid.SetColumn(name, 1);
        Grid.SetColumn(countText, 2);
        row.Children.Add(chevron);
        row.Children.Add(name);
        row.Children.Add(countText);

        var header = new Border
        {
            Background = active ? ThemeManager.Brush("AccentSoftBrush") : Brushes.Transparent,
            BorderBrush = active ? OperatorBrush : LineBrush,
            BorderThickness = new Thickness(active ? 3 : 0, 0, 0, 1),
            Padding = new Thickness(6, 8, 8, 8),
            Margin = new Thickness(0, 10, 0, 6),
            Cursor = Cursors.Hand,
            Child = row,
            Tag = group,
        };

        AutomationProperties.SetName(header, $"分組：{group.Name}（{(collapsed ? "已收合" : "已展開")}，{count} 項）");
        AutomationProperties.SetAutomationId(header, "grp_" + group.Id);

        header.MouseLeftButtonUp += (_, _) =>
        {
            _activeGroupId = group.Id;
            RenderSources();
            Log($"目前分組：{group.Name}（此後匯入的素材會放進這一組）");
        };

        var menu = new ContextMenu();

        var expandItem = new MenuItem { Header = collapsed ? "展開這一組" : "收合這一組" };
        expandItem.Click += (_, _) => ToggleGroupCollapsed(group);
        menu.Items.Add(expandItem);

        var rename = new MenuItem { Header = "重新命名分組", IsEnabled = !group.IsDefault };
        rename.Click += (_, _) => RenameGroup(group);
        menu.Items.Add(rename);

        var delete = new MenuItem { Header = "刪除分組（素材移到默認分組）", IsEnabled = !group.IsDefault };
        delete.Click += (_, _) => DeleteGroup(group);
        menu.Items.Add(delete);

        header.ContextMenu = menu;
        return header;
    }

    private void SourceSearchToggle_Click(object sender, RoutedEventArgs e)
    {
        bool opening = SourceSearchPanel.Visibility != Visibility.Visible;
        SourceSearchPanel.Visibility = opening ? Visibility.Visible : Visibility.Collapsed;
        SourceSearchToggle.Content = opening ? "收起" : "查找";

        if (opening)
        {
            SourceSearchBox.Focus();
            return;
        }

        SourceSearchBox.Text = string.Empty;
        _sourceSearchKeyword = string.Empty;
        SourceSearchStatus.Text = "輸入關鍵字後按「查找」";
        RenderSources();
    }

    private void SourceSearchButton_Click(object sender, RoutedEventArgs e)
    {
        ApplySourceSearch();
    }

    private void SourceSearchClearButton_Click(object sender, RoutedEventArgs e)
    {
        SourceSearchBox.Text = string.Empty;
        _sourceSearchKeyword = string.Empty;
        SourceSearchStatus.Text = "輸入關鍵字後按「查找」";
        RenderSources();
        SourceSearchBox.Focus();
    }

    private void SourceSearchBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            ApplySourceSearch();
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            SourceSearchBox.Text = string.Empty;
            _sourceSearchKeyword = string.Empty;
            SourceSearchPanel.Visibility = Visibility.Collapsed;
            SourceSearchToggle.Content = "查找";
            RenderSources();
            e.Handled = true;
        }
    }

    private void ApplySourceSearch()
    {
        _sourceSearchKeyword = SourceSearchBox.Text.Trim();
        RenderSources();

        if (_sourceSearchKeyword.Length == 0)
        {
            Log("素材查找：已清除關鍵字，恢復完整清單。");
        }
        else
        {
            int found = _sources.Count(s => s.Kind == "material" && SourceMatchesKeyword(s));
            Log(found == 0
                ? $"素材查找：「{_sourceSearchKeyword}」沒有符合項目。"
                : $"素材查找：「{_sourceSearchKeyword}」找到 {found} 項。");
        }
    }

    /// <summary>切換分組的展開／收合，並把狀態寫進設定檔（重啟後保持）。</summary>
    private void ToggleGroupCollapsed(MaterialGroup group)
    {
        if (_collapsedGroups.Contains(group.Id))
        {
            _collapsedGroups.Remove(group.Id);
            Log($"已展開分組：{group.Name}");
        }
        else
        {
            _collapsedGroups.Add(group.Id);
            Log($"已收合分組：{group.Name}");
        }

        _settings.CollapsedGroups = _collapsedGroups.ToList();
        SettingsStore.Save(_settings);
        RenderSources();
    }

    private Button BuildSourceCard(SourceItem source, int index)
    {
        // 「播出中」只給真的在螢幕上的來源；選取其他來源不會把原本的投影標記熄掉。
        bool isLive = IsProjected(source);
        bool isSelected = _selected == source && !isLive;
        var card = new Button
        {
            Style = (Style)FindResource("SourceCardStyle"),
            Background = isLive
                ? CardLiveBrush
                : isSelected ? ThemeManager.Brush("Panel3Brush") : Brushes.Transparent,
            BorderBrush = isLive
                ? OperatorBrush
                : isSelected ? ThemeManager.Brush("LineBrush") : Brushes.Transparent,
            Tag = source,
        };
        AutomationProperties.SetName(card, source.Name);
        AutomationProperties.SetAutomationId(card, "src_" + index);

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var title = new TextBlock
        {
            Text = source.Name,
            Foreground = isLive ? ThemeManager.Brush("AccentTextBrush") : ThemeManager.Brush("TextPrimaryBrush"),
            FontSize = 13,
            TextTrimming = TextTrimming.CharacterEllipsis,
        };

        var sub = new TextBlock
        {
            Text = source.Kind == "main" && _mainPath is null
                ? "尚未加入"
                : isLive
                    ? "播出中"
                    : isSelected
                        ? "已選擇 · 待推送"
                        : source.Kind switch
                        {
                            "main" => "主 PPT",
                            "scripture" => "經文",
                            "web" => "網頁",
                            "zoom" => "Zoom",
                            _ => "素材",
                        },
            Foreground = isLive
                ? OperatorBrush
                : isSelected ? ThemeManager.Brush("TextDimBrush") : FaintBrush,
            FontSize = 12,
            Margin = new Thickness(0, 3, 0, 0),
        };

        var meta = new TextBlock
        {
            Text = MetaOf(source),
            Foreground = FaintBrush,
            FontSize = 12,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Top,
        };

        var texts = new StackPanel();
        texts.Children.Add(title);
        texts.Children.Add(sub);

        Grid.SetColumn(texts, 0);
        Grid.SetColumn(meta, 1);
        grid.Children.Add(texts);
        grid.Children.Add(meta);

        card.Content = grid;
        card.Click += (_, _) => ActivateSource(source);

        if (source.Kind == "main")
        {
            card.ContextMenu = BuildMainMenu();
        }
        else if (source.Kind == "material" && source.Material is not null)
        {
            card.ContextMenu = BuildMaterialMenu(source.Material);
        }

        return card;
    }

    /// <summary>主 PPT 卡片的右鍵選單：加入／更換、加入上次那份、移除。</summary>
    private ContextMenu BuildMainMenu()
    {
        var menu = new ContextMenu();

        var pick = new MenuItem { Header = _mainPath is null ? "加入主 PPT…" : "更換主 PPT…" };
        pick.Click += (_, _) => OpenPresentation();
        menu.Items.Add(pick);

        string? recent = _settings.LastPresentation;
        if (_mainPath is null && !string.IsNullOrEmpty(recent) && File.Exists(recent))
        {
            var useRecent = new MenuItem
            {
                Header = "加入上次的簡報（" + Path.GetFileNameWithoutExtension(recent) + "）",
            };
            useRecent.Click += (_, _) => AddMainPresentation(recent!);
            menu.Items.Add(useRecent);
        }

        var remove = new MenuItem { Header = "移除主 PPT", IsEnabled = _mainPath is not null };
        remove.Click += (_, _) => RemoveMainPresentation();
        menu.Items.Add(remove);

        return menu;
    }

    /// <summary>素材卡片的右鍵選單：重新匯出、換分組、移除。</summary>
    private ContextMenu BuildMaterialMenu(MaterialItem item)
    {
        var menu = new ContextMenu();

        var reexport = new MenuItem { Header = "重新匯出頁面", IsEnabled = !item.IsImage };
        reexport.Click += (_, _) =>
        {
            _selected = _sources.FirstOrDefault(s => s.Material?.Id == item.Id);
            ReexportMaterial_Click(this, new RoutedEventArgs());
        };
        menu.Items.Add(reexport);

        foreach (var group in _groups.Where(g => g.Id != item.GroupId))
        {
            var move = new MenuItem { Header = "移到分組：" + group.Name };
            move.Click += (_, _) => MoveMaterialToGroup(item, group.Id);
            menu.Items.Add(move);
        }

        menu.Items.Add(new Separator());

        var remove = new MenuItem { Header = "移除素材" };
        remove.Click += (_, _) =>
        {
            _selected = _sources.FirstOrDefault(s => s.Material?.Id == item.Id);
            RemoveMaterial_Click(this, new RoutedEventArgs());
        };
        menu.Items.Add(remove);

        return menu;
    }

    private void MoveMaterialToGroup(MaterialItem item, string groupId)
    {
        item.GroupId = groupId;
        SaveLibrary();
        RebuildSources();
        RenderAll();

        string groupName = _groups.FirstOrDefault(g => g.Id == groupId)?.Name ?? "默認分組";
        Log($"已將素材「{item.Name}」移到分組「{groupName}」。");
    }

    /// <summary>把指定路徑登記為主 PPT（不立刻開 PowerPoint，等按「開始放映」才開）。</summary>
    private void AddMainPresentation(string path)
    {
        _mainPath = path;
        _settings.LastPresentation = path;
        SettingsStore.Save(_settings);

        MainSource.Name = Path.GetFileNameWithoutExtension(path);
        _previewFile = null;
        _previewPage = -1;

        Log("已加入主 PPT：" + path + "（按「開始放映」時才由 PowerPoint 開啟）");
        RenderAll();
    }

    /// <summary>移除主 PPT，左欄回到「＋ 加入主 PPT」的空白狀態。</summary>
    private void RemoveMainPresentation()
    {
        if (_ppt.IsShowRunning)
        {
            MessageBox.Show(this, "放映進行中，請先按「結束放映」再移除主 PPT。",
                "移除主 PPT", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string? old = _mainPath;
        _ppt.ClosePresentation();

        _mainPath = null;
        _previewFile = null;
        _previewPage = -1;
        MainSource.Name = "＋ 加入主 PPT（點此選檔）";

        Log(old is null
            ? "目前沒有主 PPT。"
            : "已移除主 PPT：" + Path.GetFileName(old) + "（原始檔案沒有刪除）");

        RenderAll();
    }

    private string MetaOf(SourceItem source)
    {
        if (source.Kind == "main")
        {
            if (_mainPath is null)
            {
                // 卡片標題已寫「＋ 加入主 PPT（點此選檔）」、副標已寫「尚未加入」，
                // 右側不再重複第三個「尚未加入」。
                return string.Empty;
            }

            return _ppt.IsShowRunning ? $"第 {_ppt.CurrentSlide} / {ShowTotal} 頁" : "已加入";
        }

        if (source.Kind == "scripture")
        {
            var passage = source.Scripture;
            return passage is null ? "—" : $"{passage.FontSize} pt";
        }

        if (source.Kind == "web")
        {
            return source.Web?.Host ?? "—";
        }

        if (source.Kind == "zoom")
        {
            return _zoomService.IsActive ? "已鋪到投影幕" : "未啟用";
        }

        int count = source.Material?.Pages.Count ?? 0;
        return $"{count} 頁";
    }

    /// <summary>只有真正在螢幕 2 上的來源才算「播出中」；單純選取不算。</summary>
    private bool IsProjected(SourceItem? source)
    {
        if (source is null || _projectedSource != source)
        {
            return false;
        }

        return source.Kind switch
        {
            "main" => _mainPath is not null && _ppt.IsShowRunning,
            "material" => _materialProjectionActive,
            "scripture" or "web" => _extraProjectionActive,
            "zoom" => _zoomProjectionActive,
            _ => false,
        };
    }

    /// <summary>切換投影源時才收掉上一份投影；單純選取其他來源不會呼叫這裡。</summary>
    private void StopProjectionForSwitch(SourceItem target)
    {
        if (_projectedSource == target)
        {
            return;
        }

        if (_extraProjectionActive)
        {
            _scriptureWindow?.HideWindow();
            _webWindow?.HideWindow();
            _extraProjectionActive = false;
        }

        if (_materialProjectionActive)
        {
            _materialWindow?.HideWindow();
            _materialProjectionActive = false;
        }

        if (_zoomProjectionActive)
        {
            _zoomService.Restore();
            _zoomProjectionActive = false;
        }

        // 切換到其他投影源時不要結束主 PPT 放映：只把它遮黑留在背景，
        // 這樣斷開素材／經文／網頁後，主 PPT 還能原頁回到螢幕 2。
        if (_projectedSource?.Kind == "main" && _ppt.IsShowRunning)
        {
            _ppt.SetScreenState(PowerPointService.StateBlackScreen);
        }

        _projectedSource = null;
    }

    /// <summary>中斷目前已選來源的投影：回到主 PPT；沒有主 PPT 放映才回到桌面。</summary>
    private void DisconnectProjection()
    {
        var source = _live;
        if (!IsProjected(source))
        {
            return;
        }

        switch (source!.Kind)
        {
            case "material":
                _materialWindow?.HideWindow();
                _materialProjectionActive = false;
                break;
            case "scripture":
            case "web":
                _scriptureWindow?.HideWindow();
                _webWindow?.HideWindow();
                _extraProjectionActive = false;
                break;
            case "zoom":
                _zoomService.Restore();
                _zoomProjectionActive = false;
                break;
            case "main":
                _ppt.EndShow();
                break;
        }

        _black = false;
        if (source.Kind != "main" && _ppt.IsShowRunning)
        {
            // 主 PPT 放映還在，只是剛才被遮黑；斷開後恢復它，而不是留下桌面。
            _ppt.SetScreenState(PowerPointService.StateRunning);
            _projectedSource = MainSource;
            _live = _projectedSource;
            _selected = _projectedSource;
            _lastMainPage = _ppt.CurrentSlide;
            Log($"已斷開投影：{DescribeSourceName(source)}；已返回主 PPT 第 {_ppt.CurrentSlide} 頁。");
        }
        else
        {
            _projectedSource = null;
            Log($"已斷開投影：{DescribeSourceName(source)}。控制台仍保留預覽。");
        }

        RenderAll();
    }

    private static string DescribeSourceName(SourceItem source)
    {
        return source.Kind switch
        {
            "main" => "主 PPT",
            "material" => $"素材 · {source.Material?.Name}",
            "scripture" => $"經文 · {source.Scripture?.Reference}",
            "web" => $"網頁 · {source.Web?.Name}",
            "zoom" => "Zoom 會議",
            _ => "投影源",
        };
    }

    private void ActivateSource(SourceItem source)
    {
        if (BeforeActivate(source))
        {
            return;   // 經文／網頁／Zoom 走另一條路徑（MainWindow.Sources.cs）
        }

        if (source.Kind == "material" && (source.Material is null || source.Material.Pages.Count == 0))
        {
            Log("這份素材沒有可用頁面，請重新匯出。");
            return;
        }

        var previous = _live;
        _live = source;
        _selected = source;

        if (source.Kind == "main")
        {
            // 主 PPT 還沒加入時，點這張卡片就直接叫他選檔，不要只丟一句「尚未開啟簡報」
            if (_mainPath is null)
            {
                if (!OpenPresentation())
                {
                    return;
                }

                _live = source;
                _selected = source;
                Log("主 PPT 已加入，按「開始放映」就會由 PowerPoint 開啟並投出。");
                RenderAll();
                return;
            }

            if (_ppt.IsShowRunning)
            {
                Log("已切換至主 PPT（PowerPoint 原生放映中，頁碼不變）。");
            }
            else
            {
                Log("已切換至主 PPT。目前尚未放映，按「開始放映」後才會出現在投影幕。");
            }
        }
        else
        {
            // 切到另一份素材時回到第 1 頁；同一份或它已在投影中則保留目前頁。
            if (previous != source && _projectedSource != source)
            {
                source.PageIndex = 0;
            }

            if (IsProjected(source))
            {
                Log($"已選擇素材「{source.Material!.Name}」第 {source.PageIndex + 1} 頁；此來源正在投影，按「斷開投影」可停止。");
            }
            else
            {
                Log($"已選擇素材「{source.Material!.Name}」第 {source.PageIndex + 1} 頁；尚未投影，請按「推送投影」。");
            }
            if (SingleScreen)
            {
                Log("目前只有一台螢幕：推送後仍只在控制台預覽，接上投影幕設「延伸」才會送到螢幕 2。");
            }
        }

        RenderAll();
    }

    private void PushMaterialButton_Click(object sender, RoutedEventArgs e)
    {
        if (IsProjected(_live))
        {
            DisconnectProjection();
            return;
        }

        if (_live?.Kind is "scripture" or "web" or "zoom")
        {
            PushExtraProjection();
            return;
        }

        if (_live?.Kind != "material" || _live.Material is null)
        {
            Log("請先在左欄選擇一份素材。");
            return;
        }

        if (_target is null)
        {
            Log("尚未選擇投影螢幕，無法推送素材。");
            return;
        }

        StopProjectionForSwitch(_live);
        _projectedSource = _live;
        _materialProjectionActive = true;
        _black = false;

        if (SingleScreen)
        {
            _materialWindow?.HideWindow();
            Log("目前只有一台螢幕：素材只在控制台預覽。接上投影幕並設為「延伸」後，推送才會送到螢幕 2。");
            RenderAll();
            return;
        }

        ShowMaterialPage();
        Log($"已推送素材「{_live.Material.Name}」第 {_live.PageIndex + 1} 頁到投影螢幕。");
        RenderAll();
    }

    private void ShowMaterialPage()
    {
        var source = _projectedSource;
        if (!_materialProjectionActive || _target is null || source?.Material is null)
        {
            return;
        }

        if (source.Material.Pages.Count == 0)
        {
            return;
        }

        if (SingleScreen)
        {
            // 單螢幕：不投出，只在控制台中央預覽
            _materialWindow?.HideWindow();
            return;
        }

        int page = Math.Clamp(source.PageIndex, 0, source.Material.Pages.Count - 1);
        source.PageIndex = page;

        _materialWindow ??= new MaterialWindow();
        _materialWindow.ShowPage(_target, source.Material.Pages[page]);
        _lastMaterialPage = page;
    }

    private void FlipMaterial(bool forward)
    {
        if (_live?.Material is null)
        {
            return;
        }

        int next = _live.PageIndex + (forward ? 1 : -1);
        int last = _live.Material.Pages.Count - 1;

        if (next < 0)
        {
            Log("素材已在第一頁。");
            return;
        }

        if (next > last)
        {
            Log($"素材已在最後一頁（共 {last + 1} 頁）。");
            return;
        }

        _live.PageIndex = next;
        _black = false;

        if (IsProjected(_live))
        {
            _materialWindow?.SetPage(_live.Material.Pages[next]);
            _lastMaterialPage = next;
            Log($"素材 → 第 {next + 1} / {last + 1} 頁（投影畫面已同步更新）。");
        }
        else
        {
            Log($"素材預覽 → 第 {next + 1} / {last + 1} 頁（尚未投影）。");
        }
    }

    // ---------------------------------------------------------------- 渲染

    private void RenderAll()
    {
        RenderSources();
        RenderThumbs();
        RenderPreview();
        UpdateShowStatus();
    }

    /// <summary>每次放映換一個暫存資料夾，舊圖不會被 ImageCache 當成新圖。</summary>
    private void PreparePreviewCache()
    {
        try
        {
            var root = Path.Combine(Path.GetTempPath(), "ProjectionManager", "preview");
            if (Directory.Exists(root))
            {
                Directory.Delete(root, true);
            }

            _previewDir = Path.Combine(root, DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            Directory.CreateDirectory(_previewDir);
        }
        catch
        {
            _previewDir = null;
        }

        _previewPage = -1;
        _previewFile = null;
        _previewSlowNoted = false;
        _thumbPrefetchStarted = false;
    }

    /// <summary>一次把主 PPT 全部頁面匯出成小縮圖，避免未造訪頁一直顯示黑底頁碼磚。</summary>
    private void PrefetchMainThumbs(int total)
    {
        if (_thumbPrefetchStarted || _previewDir is null)
        {
            return;
        }

        _thumbPrefetchStarted = true;
        int exported = 0;
        try
        {
            for (int page = 1; page <= total; page++)
            {
                if (_live?.Kind != "main")
                {
                    return;
                }

                string file = Path.Combine(_previewDir, $"thumb_p{page:D3}.png");
                if (File.Exists(file))
                {
                    continue;
                }

                _ppt.ExportSlideTo(page, file, 320, 180);
                exported++;
            }
        }
        catch (Exception ex)
        {
            Log($"縮圖預取中斷：{ex.Message}");
        }

        if (exported > 0 && _live?.Kind == "main")
        {
            RenderThumbs();
            // 右側「下一頁」預覽可能還沒圖；預取完成後補上，不必等下一次換頁。
            RenderPreview();
        }
    }

    /// <summary>放映中優先用放映視窗自己的總頁數；沒放映時才用簡報檔的頁數。</summary>
    private int ShowTotal
    {
        get
        {
            int live = _ppt.ShowTotalSlides;
            return live > 0 ? live : _ppt.TotalSlides;
        }
    }

    /// <summary>
    /// 主 PPT 是 PowerPoint 原生放映，程式不重繪它；控制台要看到畫面，
    /// 就把「當前頁」用 Slide.Export 匯成一張圖顯示。實測放映中匯出可行，
    /// 但每次換頁會佔用 UI 執行緒約 0.2–1.3 秒，所以只在頁碼真的變了才做。
    /// </summary>
    private bool RefreshMainPreview()
    {
        if (_live?.Kind != "main" || !_ppt.IsShowRunning)
        {
            return false;
        }

        int page = _ppt.CurrentSlide;
        if (page <= 0)
        {
            return false;
        }

        if (string.IsNullOrEmpty(_previewDir))
        {
            PreparePreviewCache();
        }

        if (string.IsNullOrEmpty(_previewDir))
        {
            return false;
        }

        // 頁碼徽章只看放映狀態，不依賴匯出成功與否：匯出慢或失敗時頁數仍然看得到
        int total = ShowTotal;
        PreviewLivePage.Text = $"第 {page} / {Math.Max(total, 1)} 頁";
        PreviewLivePageBadge.Visibility = Visibility.Visible;

        if (_previewPage != page)
        {
            _previewPage = page;
            string file = Path.Combine(_previewDir!, $"main_p{page:D3}.png");
            _previewFile = File.Exists(file) ? file : null;

            if (_previewFile is null)
            {
                SchedulePreviewExport(page);
            }
        }

        if (_previewFile is null)
        {
            // 新頁的圖還在匯出：保留上一張圖，不把畫面清空
            if (PreviewLiveImage.Source is null)
            {
                PreviewLiveHint.Text = $"PowerPoint 原生放映中　第 {page} / {Math.Max(total, 1)} 頁\n（正在產生靜態預覽圖…）";
                PreviewLiveHint.Visibility = Visibility.Visible;
            }

            return false;
        }

        var bitmap = ImageCache.Load(_previewFile);
        if (bitmap is null)
        {
            return false;
        }

        PreviewLiveImage.Source = bitmap;
        PreviewLiveHint.Visibility = Visibility.Collapsed;
        PreviewLiveName.Text = $"主 PPT　第 {page} / {Math.Max(total, 1)} 頁 · 靜態預覽";
        return true;
    }

    /// <summary>
    /// 主 PPT 放映中，右側顯示真正的下一頁。縮圖預取完成後可直接載入，
    /// 不必在每次換頁時再多做一次 COM 匯出。
    /// </summary>
    private void RefreshMainNextPreview(int current, int total)
    {
        PreviewNextImage.Source = null;
        PreviewNextPageBadge.Visibility = Visibility.Collapsed;

        if (current <= 0 || current >= total)
        {
            PreviewNextName.Text = "—";
            PreviewNextHint.Text = "已到最後一頁";
            PreviewNextHint.Visibility = Visibility.Visible;
            return;
        }

        int next = current + 1;
        PreviewNextName.Text = $"第 {next} / {Math.Max(total, 1)} 頁";

        if (string.IsNullOrEmpty(_previewDir))
        {
            PreviewNextHint.Text = "正在產生下一頁預覽…";
            PreviewNextHint.Visibility = Visibility.Visible;
            return;
        }

        string thumb = Path.Combine(_previewDir, $"thumb_p{next:D3}.png");
        if (!File.Exists(thumb))
        {
            PreviewNextHint.Text = "正在產生下一頁預覽…";
            PreviewNextHint.Visibility = Visibility.Visible;
            return;
        }

        var bitmap = ImageCache.Load(thumb);
        if (bitmap is null)
        {
            PreviewNextHint.Text = "下一頁預覽載入失敗";
            PreviewNextHint.Visibility = Visibility.Visible;
            return;
        }

        PreviewNextImage.Source = bitmap;
        PreviewNextHint.Visibility = Visibility.Collapsed;
        PreviewNextPage.Text = $"第 {next} / {Math.Max(total, 1)} 頁";
        PreviewNextPageBadge.Visibility = Visibility.Visible;
    }

    /// <summary>
    /// 把匯出排到 UI 佇列的最低優先權：滑鼠點擊、頁碼變更先反映，COM 匯出（0.2–1.3 秒）
    /// 等介面畫完才做，所以大幅跳頁不會出現「按了沒反應」。
    /// </summary>
    private void SchedulePreviewExport(int page)
    {
        if (_previewExportRunning)
        {
            return;
        }

        if (Environment.TickCount64 < _previewRetryAt)
        {
            return;
        }

        _previewExportRunning = true;
        Dispatcher.BeginInvoke(new Action(() =>
        {
            try
            {
                if (_live?.Kind != "main" || !_ppt.IsShowRunning || string.IsNullOrEmpty(_previewDir))
                {
                    return;
                }

                int current = _ppt.CurrentSlide;
                if (current <= 0)
                {
                    return;
                }

                // 等排到這裡時頁碼可能又變了：直接匯出「目前這一頁」
                string file = Path.Combine(_previewDir!, $"main_p{current:D3}.png");
                if (!File.Exists(file))
                {
                    var watch = System.Diagnostics.Stopwatch.StartNew();
                    _ppt.ExportSlideTo(current, file, 1024, 576);
                    watch.Stop();
                    if (!_previewSlowNoted && watch.ElapsedMilliseconds > 1200)
                    {
                        _previewSlowNoted = true;
                        Log($"預覽圖匯出花了 {watch.ElapsedMilliseconds} ms（PowerPoint 忙時會慢），投影畫面不受影響。");
                    }
                }

                if (!File.Exists(file))
                {
                    _previewRetryAt = Environment.TickCount64 + 1500;
                    return;
                }

                if (current == _ppt.CurrentSlide)
                {
                    _previewPage = current;
                    _previewFile = file;
                    var bitmap = ImageCache.Load(file);
                    if (bitmap is not null)
                    {
                        PreviewLiveImage.Source = bitmap;
                        PreviewLiveHint.Visibility = Visibility.Collapsed;
                        PreviewLiveName.Text = $"主 PPT　第 {current} / {Math.Max(ShowTotal, 1)} 頁 · 靜態預覽";
                    }
                }
            }
            catch
            {
                _previewRetryAt = Environment.TickCount64 + 1500;
            }
            finally
            {
                _previewExportRunning = false;
            }
        }), System.Windows.Threading.DispatcherPriority.Background);
    }


    private void RenderPreview()
    {
        // 主 PPT 放映中：保留上一張預覽圖，等新頁的圖匯出好再換，避免跳頁時畫面一閃空白
        if (_live?.Kind != "main" || !_ppt.IsShowRunning)
        {
            PreviewLiveImage.Source = null;
        }

        PreviewNextImage.Source = null;
        PreviewNextText.Text = string.Empty;
        PreviewNextText.Visibility = Visibility.Collapsed;
        PreviewLiveHint.Visibility = Visibility.Visible;
        PreviewNextHint.Visibility = Visibility.Visible;
        PreviewLiveText.Visibility = Visibility.Collapsed;
        PreviewLiveText.Text = string.Empty;
        PreviewLivePageBadge.Visibility = Visibility.Collapsed;
        PreviewNextPageBadge.Visibility = Visibility.Collapsed;
        PreviewLiveStateLabel.Text = "尚未播出";
        PreviewNextName.Text = "—";
        WebNavPanel.Visibility = Visibility.Collapsed;

        if (_live is null)
        {
            PreviewLiveHint.Text = "尚未選擇投影源";
            PreviewNextHint.Text = "—";
            PreviewLiveName.Text = "—";
            return;
        }

        if (_live.Kind is "scripture" or "web" or "zoom")
        {
            bool pushed = IsProjected(_live);
            PreviewLiveStateLabel.Text = _black
                ? "黑屏"
                : pushed ? "播出中" : "已選擇 · 待推送";
            PreviewNextHint.Text = "—";
            RenderExtraPreview();
            return;
        }

        if (_live.Kind == "main")
        {
            int current = _ppt.CurrentSlide;
            int total = ShowTotal;
            PreviewLiveStateLabel.Text = _ppt.IsShowRunning ? "播出中" : "已停止放映";
            PreviewLiveName.Text = _ppt.IsShowRunning
                ? $"主 PPT　第 {current} / {Math.Max(total, 1)} 頁"
                : "主 PPT";
            PreviewLiveHint.Text = _ppt.IsShowRunning
                ? $"PowerPoint 原生放映中　第 {current} / {Math.Max(total, 1)} 頁\n（動畫與特效由 PowerPoint 繪製，此處為靜態預覽）"
                : _mainPath is null
                    ? "主 PPT 尚未加入\n點左欄「＋ 加入主 PPT」或按上方「加入主 PPT」選檔"
                    : "主 PPT 已加入，尚未放映\n按上方「開始放映」即由 PowerPoint 開啟並投出";

            if (!RefreshMainPreview() && PreviewLiveImage.Source is null)
            {
                PreviewLiveHint.Visibility = Visibility.Visible;
            }

            if (_ppt.IsShowRunning)
            {
                // 先確保當前頁的暫存目錄已建立，再抓下一頁的預取縮圖。
                RefreshMainNextPreview(current, total);
            }
            else
            {
                PreviewNextName.Text = "—";
                PreviewNextHint.Text = "—";
            }

            return;
        }

        var material = _live.Material;
        if (material is null || material.Pages.Count == 0)
        {
            PreviewLiveHint.Text = "素材沒有可用頁面";
            return;
        }

        PreviewLiveStateLabel.Text = _materialProjectionActive && IsProjected(_live)
            ? "播出中"
            : "已選擇（未投影）";
        PreviewLiveName.Text = IsProjected(_live)
            ? $"素材 · {material.Name}"
            : $"素材 · {material.Name} · 待推送";

        var currentPath = material.Pages[_live.PageIndex];
        var cached = ImageCache.TryGet(currentPath);
        if (cached is not null)
        {
            PreviewLiveImage.Source = cached;
            PreviewLiveHint.Visibility = Visibility.Collapsed;
        }
        else
        {
            PreviewLiveImage.Source = null;
            PreviewLiveHint.Text = "正在載入預覽…";
            PreviewLiveHint.Visibility = Visibility.Visible;
            RequestPreviewLoad(PreviewLiveImage, currentPath, isNext: false);
        }

        PreviewLivePage.Text = $"第 {_live.PageIndex + 1} / {material.Pages.Count} 頁";
        PreviewLivePageBadge.Visibility = Visibility.Visible;

        if (_live.PageIndex + 1 < material.Pages.Count)
        {
            int next = _live.PageIndex + 1;
            PreviewNextName.Text = $"第 {next + 1} / {material.Pages.Count} 頁";
            PreviewNextPage.Text = $"第 {next + 1} / {material.Pages.Count} 頁";
            PreviewNextPageBadge.Visibility = Visibility.Visible;

            var nextPath = material.Pages[next];
            var nextBitmap = ImageCache.TryGet(nextPath);
            if (nextBitmap is not null)
            {
                PreviewNextImage.Source = nextBitmap;
                PreviewNextHint.Visibility = Visibility.Collapsed;
            }
            else
            {
                PreviewNextHint.Text = "正在載入下一頁…";
                RequestPreviewLoad(PreviewNextImage, nextPath, isNext: true);
            }
        }
        else
        {
            PreviewNextName.Text = "—";
            PreviewNextHint.Text = "已到最後一頁";
        }
    }

    /// <summary>預覽圖背景解碼：先讓卡片與頁碼出現，再補上整頁大圖。</summary>
    private void RequestPreviewLoad(Image target, string path, bool isNext)
    {
        string key = path;
        if (isNext)
        {
            if (_nextPreviewRequestedKey == key)
            {
                return;
            }

            _nextPreviewRequestedKey = key;
        }
        else
        {
            if (_previewRequestedKey == key)
            {
                return;
            }

            _previewRequestedKey = key;
        }

        int generation = ++_previewLoadGeneration;
        _ = Task.Run(async () =>
        {
            await _previewLoadGate.WaitAsync();
            try
            {
                var bitmap = ImageCache.Load(path);
                if (bitmap is null)
                {
                    return;
                }

                await Dispatcher.InvokeAsync(() =>
                {
                    string? currentKey = isNext ? _nextPreviewRequestedKey : _previewRequestedKey;
                    if (currentKey != key)
                    {
                        return;
                    }

                    target.Source = bitmap;
                    if (isNext)
                    {
                        PreviewNextHint.Visibility = Visibility.Collapsed;
                    }
                    else
                    {
                        PreviewLiveHint.Visibility = Visibility.Collapsed;
                    }
                });
            }
            catch
            {
                // 單張解碼失敗時保留「載入中」提示，不影響換頁
            }
            finally
            {
                _previewLoadGate.Release();
            }
        });
    }

    private void CenterColumn_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdatePreviewHeight(e.NewSize.Width);
    }

    /// <summary>預覽本體維持精確 4:3：扣掉卡片邊框後，高度由本體寬度換算；16:9 素材上下留黑但完整顯示。</summary>
    private void UpdatePreviewHeight(double centerWidth)
    {
        double cardW = Math.Max(140, (centerWidth - 10) / 2);
        double bodyW = Math.Max(120, cardW - 7);     // 扣掉 Border 的左右邊框與內距
        double h = Math.Round(bodyW * 3 / 4 + 25);   // 25px 給「預備／播出中」標題列
        h = Math.Max(168, Math.Min(320, h));
        if (Math.Abs(PreviewRow.Height.Value - h) >= 1)
        {
            PreviewRow.Height = new GridLength(h);
        }
    }

    private void NavBar_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        // 視窗較窄時藏起右側提示，避免和跳頁按鈕重疊。
        NavHintText.Visibility = e.NewSize.Width >= 600 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void ThumbArea_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (Math.Abs(e.NewSize.Width - e.PreviousSize.Width) < 8)
        {
            return;
        }

        if (_thumbResizeTimer is null)
        {
            _thumbResizeTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(180)
            };
            _thumbResizeTimer.Tick += (_, _) =>
            {
                _thumbResizeTimer!.Stop();
                RenderThumbs();
            };
        }

        _thumbResizeTimer.Stop();
        _thumbResizeTimer.Start();
    }

    /// <summary>依縮圖區實際寬度計算圖磚尺寸，讓每一列填滿可用寬度，避免右側留下大片黑邊。</summary>
    private void UpdateThumbTileSize()
    {
        double avail = ThumbArea.ActualWidth - 34;   // 內距與垂直滾動條
        if (avail < 240)
        {
            avail = 760;                              // 視窗尚未完成配置時採預設，之後重繪會校正
        }

        int cols = Math.Max(2, Math.Min(6, (int)((avail + 8) / 190)));
        double w = (avail - cols * 8 - 2) / cols;   // 每個圖磚都帶 8px 右邊距，最後一個也算在內
        w = Math.Max(130, Math.Min(240, w));
        _thumbTileW = Math.Floor(w);
        _thumbTileH = Math.Round(_thumbTileW * 9 / 16);
        _thumbCols = cols;
    }

    private void RenderThumbs()
    {
        _thumbLoadGeneration++;   // 讓上一輪還沒解碼完的縮圖自動失效
        ThumbPanel.Children.Clear();
        _thumbBorders.Clear();
        UpdateThumbTileSize();

        if (_live is null)
        {
            return;
        }

        if (_live.Kind == "main")
        {
            int total = Math.Min(ShowTotal, 200);
            if (!_thumbPrefetchStarted && _previewDir is not null)
            {
                Dispatcher.BeginInvoke(new Action(() => PrefetchMainThumbs(total)),
                    System.Windows.Threading.DispatcherPriority.Background);
            }

            for (int i = 1; i <= total; i++)
            {
                int index = i;   // 必須在迴圈內另存：直接閉包 i 會抓到迴圈結束後的值（總頁數+1）
                string? thumb = _previewDir is null ? null : Path.Combine(_previewDir, $"main_p{index:D3}.png");
                if (thumb is not null && !File.Exists(thumb))
                {
                    string small = Path.Combine(_previewDir!, $"thumb_p{index:D3}.png");
                    thumb = File.Exists(small) ? small : null;   // 尚未預取到的頁面先顯示頁碼方塊
                }

                AddThumbTile(index.ToString(), thumb, index == _ppt.CurrentSlide, () =>
                {
                    if (!_ppt.IsShowRunning)
                    {
                        Log("主 PPT 尚未放映，無法跳頁。");
                        return;
                    }

                    if (!TryJumpToPage(index, out string error))
                    {
                        Log(error);
                    }
                });
            }

            return;
        }

        var material = _live.Material;
        if (material is null)
        {
            RenderExtraThumbs();
            return;
        }

        for (int i = 0; i < material.Pages.Count; i++)
        {
            int index = i;
            AddThumbTile((i + 1).ToString(), material.Pages[i], i == _live.PageIndex, () =>
            {
                if (!TryJumpToPage(index + 1, out string error))
                {
                    Log(error);
                }
            });
        }
    }

    /// <summary>跳頁後只更新高亮框，不重建整條縮圖（200 個元素重建一次會明顯卡）。</summary>
    private void UpdateThumbHighlight()
    {
        if (_live is null)
        {
            return;
        }

        int active = _live.Kind switch
        {
            "main" => _ppt.CurrentSlide,
            "material" => _live.PageIndex + 1,
            _ => 0
        };

        for (int i = 0; i < _thumbBorders.Count; i++)
        {
            bool current = (i + 1) == active;
            _thumbBorders[i].BorderThickness = new Thickness(current ? 2 : 1);
            _thumbBorders[i].BorderBrush = current ? OperatorBrush : LineBrush;
        }
    }

    private void AddThumbTile(string label, string? imagePath, bool current, Action onClick)
    {
        var border = new Border
        {
            Width = _thumbTileW,
            Height = _thumbTileH,
            Margin = new Thickness(0, 0, 8, 8),
            CornerRadius = new CornerRadius(4),
            BorderThickness = new Thickness(current ? 2 : 1),
            BorderBrush = current ? OperatorBrush : LineBrush,
            Background = ThemeManager.Brush("PreviewPanelBrush"),
            Cursor = Cursors.Hand,
            ClipToBounds = true,
        };

        var grid = new Grid();

        // 先放空圖框，縮圖由背景解碼補上；整頁大圖不再在 UI 執行緒同步解碼。
        var image = new Image { Stretch = Stretch.Uniform };
        grid.Children.Add(image);

        if (!string.IsNullOrEmpty(imagePath))
        {
            var cached = ImageCache.TryGet(imagePath, 480);
            if (cached is not null)
            {
                image.Source = cached;
            }
            else
            {
                LoadThumbAsync(image, imagePath, _thumbLoadGeneration);
            }
        }

        grid.Children.Add(new TextBlock
        {
            Text = label,
            Foreground = current ? OperatorBrush : FaintBrush,
            FontSize = 10,
            Margin = new Thickness(6, 3, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Top,
        });

        border.Child = grid;
        border.MouseLeftButtonUp += (_, _) => onClick();
        ThumbPanel.Children.Add(border);
        _thumbBorders.Add(border);
    }

    /// <summary>經文頁沒有截圖檔，直接把該頁文字排進縮圖，讓下方每一頁都能預覽。</summary>
    private void AddScriptureThumbTile(int pageNumber, string text, bool current, Action onClick)
    {
        var border = new Border
        {
            Width = _thumbTileW,
            Height = _thumbTileH,
            Margin = new Thickness(0, 0, 8, 8),
            CornerRadius = new CornerRadius(4),
            BorderThickness = new Thickness(current ? 2 : 1),
            BorderBrush = current ? OperatorBrush : LineBrush,
            Background = ThemeManager.Brush("PreviewPanelBrush"),
            Cursor = Cursors.Hand,
            ClipToBounds = true,
            ToolTip = text,
        };

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var number = new TextBlock
        {
            Text = pageNumber.ToString(),
            Foreground = current ? OperatorBrush : FaintBrush,
            FontSize = 10,
            Margin = new Thickness(6, 3, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Left,
        };

        var body = new TextBlock
        {
            Text = text,
            Foreground = ThemeManager.Brush("TextDimBrush"),
            FontSize = 9,
            LineHeight = 12,
            TextWrapping = TextWrapping.Wrap,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(6, 2, 6, 5),
        };

        Grid.SetRow(body, 1);
        grid.Children.Add(number);
        grid.Children.Add(body);

        border.Child = grid;
        AutomationProperties.SetName(border, $"經文第 {pageNumber} 頁");
        border.MouseLeftButtonUp += (_, _) => onClick();
        ThumbPanel.Children.Add(border);
        _thumbBorders.Add(border);
    }

    private void LoadThumbAsync(Image target, string path, int generation)
    {
        _ = Task.Run(async () =>
        {
            await _thumbLoadGate.WaitAsync();
            try
            {
                var bitmap = ImageCache.Load(path, 480);
                if (bitmap is null)
                {
                    return;
                }

                await Dispatcher.InvokeAsync(() =>
                {
                    if (generation == _thumbLoadGeneration)
                    {
                        target.Source = bitmap;
                    }
                });
            }
            catch
            {
                // 單張縮圖解不出來不影響整批；留在深色底即可
            }
            finally
            {
                _thumbLoadGate.Release();
            }
        });
    }

    // ---------------------------------------------------------------- 放映

    private bool _startingShow;

    private async Task StartShow()
    {
        if (_startingShow)
        {
            return;
        }

        if (_target is null)
        {
            Log("尚未選擇投影螢幕。");
            return;
        }

        _startingShow = true;
        try
        {
            await StartShowCore();
        }
        finally
        {
            _startingShow = false;
        }
    }

    /// <summary>
    /// 放映前確保 PowerPoint 裡有一份「活著、可以放映」的簡報：
    /// 1) 上次用過的檔案直接沿用，不必再選檔；
    /// 2) 沒有就彈選檔；
    /// 3) 開檔後等 PowerPoint 建好文件視窗（新開實例需要幾百毫秒），最多等 3 秒。
    /// </summary>
    private async Task<bool> EnsurePresentationForShow()
    {
        if (_ppt.IsShowReady)
        {
            return true;
        }

        if (string.IsNullOrEmpty(_ppt.PresentationPath) || !_ppt.HasLivePresentation)
        {
            string? path = _mainPath is not null && File.Exists(_mainPath)
                ? _mainPath
                : _ppt.PresentationPath is not null && File.Exists(_ppt.PresentationPath)
                    ? _ppt.PresentationPath
                    : null;

            if (path is null)
            {
                Log("尚未加入主 PPT，請先選擇檔案。");
                if (!OpenPresentation())
                {
                    return false;
                }

                path = _mainPath;
                if (string.IsNullOrEmpty(path) || !File.Exists(path))
                {
                    return false;
                }
            }

            try
            {
                _ppt.Open(path!);
                MainSource.Name = Path.GetFileNameWithoutExtension(path);
                Log("已由 PowerPoint 開啟主 PPT：" + path);
                RenderAll();
            }
            catch (Exception ex)
            {
                Log(ex.Message);
                return false;
            }
        }

        for (int i = 0; i < 15 && !_ppt.IsShowReady; i++)
        {
            await Task.Delay(200);
        }

        if (!_ppt.IsShowReady)
        {
            Log("PowerPoint 還沒準備好放映。若剛關閉過 PowerPoint 視窗，請再按一次「開始放映」。");
            return false;
        }

        return true;
    }

    private async Task StartShowCore()
    {
        if (!await EnsurePresentationForShow())
        {
            return;
        }

        if (LooksLikeMirror())
        {
            var result = MessageBox.Show(
                this,
                "偵測到兩台螢幕的位置與尺寸完全相同，目前可能是「複製」模式。複製模式下操作員無法獨立作業，是否仍要啟動放映？",
                "螢幕模式提醒", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes)
            {
                return;
            }
        }

        try
        {
            bool windowed = WindowedShowCheck.IsChecked == true || _testMode;
            Log(windowed
                ? "正在啟動放映（視窗化模式，失焦仍會播放動畫）…"
                : "正在啟動放映（全螢幕模式；失去焦點時 PowerPoint 會暫停動畫）…");

            _black = false;
            _ppt.StartShow(_target, 1, windowed);
            _repositionUntil = Environment.TickCount64 + 3000;   // 放映剛起來會自己搶回原位

            PreparePreviewCache();

            ApplyConsoleActivation();

            _mode = ControlMode.Operator;
            _handoverSlide = 0;
            ApplyMode();

            // 放映啟動後自動回到主 PPT 這個投影源
            _materialWindow?.HideWindow();
            _live = _sources[0];
            _selected = _live;
            _projectedSource = _live;

            Log($"放映已啟動，放映視窗句柄 0x{_ppt.ShowWindowHandle.ToInt64():X}。控制權在操作員。");

            if (SingleScreen)
            {
                Log("目前只有一台螢幕：放映鋪滿整螢幕會蓋住控制台。按 Alt+Tab 可切回控制台操作，切回放映畫面即恢復顯示。");
            }

            RenderAll();
        }
        catch (Exception ex)
        {
            Log(ex.Message);
        }
    }

    /// <summary>兩個放映選項變更：記住設定，放映中就立即套用「不主動搶焦點」。</summary>
    private void ShowOption_Changed(object sender, RoutedEventArgs e)
    {
        if (!_ready)
        {
            return;
        }

        _settings.WindowedShow = WindowedShowCheck.IsChecked == true;
        _settings.NoActivateConsole = NoActivateCheck.IsChecked == true;
        SettingsStore.Save(_settings);

        ApplyConsoleActivation();
    }

    /// <summary>只有一台螢幕：素材不開投影視窗，且控制台保持可操作。</summary>
    private bool SingleScreen => _monitors.Count < 2;

    private void ApplyConsoleActivation()
    {
        bool returnFocusToShow = NoActivateCheck.IsChecked == true && _ppt.IsShowRunning && !SingleScreen;
        IntPtr hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero)
        {
            return;
        }

        // 不要用 WS_EX_NOACTIVATE 鎖住主控制台：該 style 會讓滑鼠點擊永遠無法激活視窗，
        // 使用者從桌面點回控制台時就收不到鍵盤事件（方向鍵因此失效）。
        // 「不搶焦點」只表示程式不主動搶，使用者點擊時仍應正常取得焦點。
        Native.SetNoActivate(hwnd, false);

        if (NoActivateCheck.IsChecked == true && _ppt.IsShowRunning && SingleScreen && !_noActivateNoticeShown)
        {
            _noActivateNoticeShown = true;
            Log("只有一台螢幕：控制台保持可操作，不自動把焦點交還放映視窗。");
        }

        if (returnFocusToShow && _ppt.ShowWindowHandle != IntPtr.Zero)
        {
            Native.SetForegroundWindow(_ppt.ShowWindowHandle);
            Log("已把前台交還放映視窗；控制台仍可用滑鼠點擊取得焦點。");
        }
    }

    private void Navigate(bool forward)
    {
        if (_live?.Kind == "material")
        {
            FlipMaterial(forward);
            RenderPreview();
            RenderThumbs();
            UpdateShowStatus();
            return;
        }

        if (_live?.Kind == "scripture")
        {
            FlipScripture(forward);
            return;
        }

        if (_live?.Kind == "web")
        {
            FlipWeb(forward);
            return;
        }

        if (_live?.Kind == "zoom")
        {
            Log("Zoom 投影源：翻頁請在 Zoom 視窗內操作（或用 Zoom 的快捷鍵）。");
            return;
        }

        if (!_ppt.IsShowRunning)
        {
            Log("目前沒有進行中的放映。");
            return;
        }

        try
        {
            if (forward)
            {
                _ppt.Next();
            }
            else
            {
                _ppt.Previous();
            }

            _black = false;
        }
        catch (PowerPointException ex)
        {
            Log(ex.Message);
        }

        _lastMainPage = _ppt.CurrentSlide;
        RefreshMainShowUi();
    }

    private void PrevButton_Click(object sender, RoutedEventArgs e) => Navigate(false);

    private void NextButton_Click(object sender, RoutedEventArgs e) => Navigate(true);

    private void JumpButton_Click(object sender, RoutedEventArgs e) => JumpFromBox();

    private void JumpLastButton_Click(object sender, RoutedEventArgs e)
    {
        if (_live?.Kind == "main")
        {
            PageJumpBox.Text = Math.Max(ShowTotal, 1).ToString();
            JumpFromBox();
            return;
        }

        if (_live?.Kind == "material" && _live.Material is not null)
        {
            PageJumpBox.Text = _live.Material.Pages.Count.ToString();
            JumpFromBox();
            return;
        }

        Log("目前的投影源不支援跳頁。");
    }

    private void PageJumpBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            e.Handled = true;
            JumpFromBox();
        }
    }

    private void JumpFromBox()
    {
        string text = PageJumpBox.Text.Trim();
        if (!int.TryParse(text, out int page))
        {
            Log("請先輸入頁碼，再按「跳至」或 Enter。");
            return;
        }

        if (!TryJumpToPage(page, out string message))
        {
            Log(message);
            return;
        }

        PageJumpBox.Text = string.Empty;
        Log($"跳到第 {page} 頁。");
    }

    /// <summary>
    /// 控制台、縮圖、手機遠端共用的跳頁入口：先跳頁，再更新畫面；不重建縮圖條。
    /// </summary>
    private bool TryJumpToPage(int page, out string message)
    {
        message = string.Empty;
        if (page <= 0)
        {
            message = "頁碼必須大於 0。";
            return false;
        }

        if (_live?.Kind == "main")
        {
            if (!_ppt.IsShowRunning)
            {
                message = "主 PPT 尚未放映，無法跳頁。";
                return false;
            }

            var watch = System.Diagnostics.Stopwatch.StartNew();
            if (!_ppt.TryGotoSlide(page, out string error))
            {
                message = error;
                return false;
            }

            watch.Stop();
            _black = false;
            _lastMainPage = _ppt.CurrentSlide;
            RefreshMainShowUi();
            if (watch.ElapsedMilliseconds > 700)
            {
                Log($"跳到第 {page} 頁花了 {watch.ElapsedMilliseconds} ms（PowerPoint 本身換頁較慢，投影畫面正常）。");
            }

            return true;
        }

        if (_live?.Kind == "material" && _live.Material is not null)
        {
            int count = _live.Material.Pages.Count;
            if (count == 0)
            {
                message = "這個素材沒有頁面。";
                return false;
            }

            if (page > count)
            {
                message = $"第 {page} 頁超出範圍（共 {count} 頁），已忽略。";
                return false;
            }

            _live.PageIndex = page - 1;
            _black = false;

            if (IsProjected(_live))
            {
                _materialWindow?.SetPage(_live.Material.Pages[_live.PageIndex]);
                _lastMaterialPage = _live.PageIndex;
            }

            RenderPreview();
            UpdateThumbHighlight();
            UpdateShowStatus();
            return true;
        }

        if (_live?.Kind == "scripture")
        {
            message = "經文請用上一頁／下一頁翻頁。";
            return false;
        }

        message = "目前的投影源不支援跳頁。";
        return false;
    }

    /// <summary>主 PPT 換頁後只做必要的介面更新：預覽、頁碼、縮圖高亮、來源列。</summary>
    private void RefreshMainShowUi()
    {
        if (_ppt.IsShowRunning && _live?.Kind == "main")
        {
            RenderPreview();
            UpdateThumbHighlight();
            RenderSources();
        }

        UpdateShowStatus();
    }

    private void BlackButton_Click(object sender, RoutedEventArgs e)
    {
        var source = _projectedSource;
        if (source?.Kind is "scripture" or "web")
        {
            if (!IsProjected(source))
            {
                Log("經文／網頁尚未推送投影，不需要黑屏。請先按「推送投影」。");
                return;
            }

            _black = true;
            _scriptureWindow?.HideWindow();
            _webWindow?.HideWindow();
            Log("投影幕已黑屏（經文／網頁視窗已收起）。");
            UpdateShowStatus();
            return;
        }

        if (source?.Kind == "zoom")
        {
            Log("Zoom 由 Zoom 自己的會議視窗顯示；請在 Zoom 內按「停止分享」或按「斷開投影」。");
            return;
        }

        if (source?.Kind == "material")
        {
            if (!IsProjected(source))
            {
                Log("素材尚未推送投影，不需要黑屏。請先按「推送投影」。");
                return;
            }

            _black = true;
            _materialWindow?.HideWindow();

            // 素材是蓋在放映之上的，收起來會露出下面的 PPT；要真的黑屏就得連 PPT 一起黑
            if (_ppt.IsShowRunning)
            {
                _ppt.SetScreenState(PowerPointService.StateBlackScreen);
            }

            Log("投影幕已黑屏（素材與放映畫面都已遮黑）。");
            UpdateShowStatus();
            return;
        }

        if (!_ppt.IsShowRunning)
        {
            Log("目前沒有進行中的放映。");
            return;
        }

        _black = true;
        _ppt.SetScreenState(PowerPointService.StateBlackScreen);
        UpdateShowStatus();
    }

    private void ResumeButton_Click(object sender, RoutedEventArgs e)
    {
        _black = false;

        var source = _projectedSource;
        if (source?.Kind is "scripture" or "web")
        {
            if (!IsProjected(source))
            {
                Log("經文／網頁尚未推送投影，沒有畫面需要恢復；請按「推送投影」。");
                UpdateShowStatus();
                return;
            }

            RestoreExtraProjection();
            UpdateShowStatus();
            return;
        }

        if (source?.Kind == "material")
        {
            if (!IsProjected(source))
            {
                Log("素材目前只是預覽，尚未推送投影；請按「推送投影」。");
                UpdateShowStatus();
                return;
            }

            if (_ppt.IsShowRunning)
            {
                _ppt.SetScreenState(PowerPointService.StateRunning);
            }

            ShowMaterialPage();
            Log("已恢復素材畫面。");
            UpdateShowStatus();
            return;
        }

        if (!_ppt.IsShowRunning)
        {
            return;
        }

        _ppt.SetScreenState(PowerPointService.StateRunning);
        UpdateShowStatus();
    }

    private void RepositionButton_Click(object sender, RoutedEventArgs e)
    {
        if (_target is null)
        {
            Log("尚未選擇投影螢幕。");
            return;
        }

        if (_projectedSource?.Kind == "material")
        {
            if (!IsProjected(_projectedSource))
            {
                Log("素材尚未推送投影，沒有需要重新定位的投影視窗。");
                return;
            }

            ShowMaterialPage();
            Log("已將素材畫面重新定位到投影螢幕。");
            return;
        }

        if (!_ppt.IsShowRunning)
        {
            Log("沒有可定位的放映視窗。");
            return;
        }

        if (_ppt.Reposition(_target))
        {
            Log("已將放映視窗重新定位到投影螢幕。");
        }
        else
        {
            Log("重新定位失敗：放映視窗可能已關閉或已結束放映。");
        }
    }

    private void EndShowButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_ppt.IsShowRunning)
        {
            Log("目前沒有進行中的放映。");
            return;
        }

        Native.SetNoActivate(new WindowInteropHelper(this).Handle, false);

        var result = MessageBox.Show(this, "結束放映後，講員將無法繼續翻頁。確定要結束嗎？",
            "結束放映", MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes)
        {
            ApplyConsoleActivation();
            return;
        }

        _ppt.EndShow();
        if (_projectedSource?.Kind == "main")
        {
            _projectedSource = null;
        }
        ApplyConsoleActivation();
        _black = false;
        _mode = ControlMode.Operator;
        _handoverSlide = 0;
        ApplyMode();
        Log("放映已結束，控制權回到操作員。素材源不受影響。");
        RenderAll();
    }

    private void OpenButton_Click(object sender, RoutedEventArgs e)
    {
        OpenPresentation();
    }

    private void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SettingsWindow(_settings)
        {
            Owner = this,
        };
        dialog.ShowDialog();

        // 主題在設定視窗內即時預覽；關閉後再套用一次，確保主視窗與所有主題畫刷同步。
        ThemeManager.Apply(_settings.UiTheme);
        SettingsStore.Save(_settings);

        // 調試模式可能在設定裡被打開／關閉：重新整理螢幕清單，立即出現或移除模擬屏幕 2。
        LoadMonitors();
        ApplyDebugBadge();
        RenderAll();
    }

    private void ApplyDebugBadge()
    {
        if (DebugBadge is null)
        {
            return;
        }

        bool simulated = (_settings.DebugMode || _testMode)
            && _monitors.Any(m => m.DeviceName == "\\\\.\\TEST2");
        DebugBadge.Visibility = simulated ? Visibility.Visible : Visibility.Collapsed;
        DebugBadgeText.Text = simulated ? "調試模式 · 螢幕 2 已模擬" : string.Empty;
    }

    private async void StartShowButton_Click(object sender, RoutedEventArgs e) => await StartShow();

    /// <summary>
    /// 把主 PPT「加入」清單。只記住檔案位置，不急著開 PowerPoint：
    /// 加入的當下不該跳出編輯視窗打斷聚會，真正開檔等按「開始放映」時才做。
    /// </summary>
    private bool OpenPresentation()
    {
        var dialog = new OpenFileDialog
        {
            Title = "加入主 PPT",
            Filter = "PowerPoint 簡報|*.pptx;*.ppt",
        };

        if (ExistingFolder(_settings.LastPresentationFolder) is { } recentFolder)
        {
            dialog.InitialDirectory = recentFolder;
        }

        if (dialog.ShowDialog() != true)
        {
            return false;
        }

        _mainPath = dialog.FileName;
        _settings.LastPresentation = dialog.FileName;
        _settings.LastPresentationFolder = Path.GetDirectoryName(dialog.FileName);
        SettingsStore.Save(_settings);
        MainSource.Name = Path.GetFileNameWithoutExtension(dialog.FileName);
        Log("已加入主 PPT：" + dialog.FileName + "（按「開始放映」時才由 PowerPoint 開啟）");
        RenderAll();
        return true;
    }

    private static string? ExistingFolder(string? path) =>
        !string.IsNullOrWhiteSpace(path) && Directory.Exists(path) ? path : null;

    private bool LooksLikeMirror()
    {
        if (_monitors.Count < 2)
        {
            return false;
        }

        var first = _monitors[0];
        var second = _monitors[1];
        return first.Left == second.Left && first.Top == second.Top
               && first.Width == second.Width && first.Height == second.Height;
    }

    // ---------------------------------------------------------------- 素材導入

    private static readonly string[] PresentationExtensions = { ".pptx", ".ppt", ".ppsx", ".pptm" };
    private static readonly string[] ImageExtensions = { ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff" };

    private static bool IsPresentation(string path) =>
        PresentationExtensions.Contains(Path.GetExtension(path).ToLowerInvariant());

    private static bool IsImageFile(string path) =>
        ImageExtensions.Contains(Path.GetExtension(path).ToLowerInvariant());

    private void ImportMaterial_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "選擇素材（簡報或圖片）",
            Filter = "支援的素材|*.pptx;*.ppt;*.ppsx;*.pptm;*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff"
                     + "|PowerPoint 簡報|*.pptx;*.ppt;*.ppsx;*.pptm"
                     + "|圖片|*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff",
        };

        if (ExistingFolder(_settings.LastMaterialFolder) is { } recentFolder)
        {
            dialog.InitialDirectory = recentFolder;
        }

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        _settings.LastMaterialFolder = Path.GetDirectoryName(dialog.FileName);
        SettingsStore.Save(_settings);
        ImportPaths(new[] { dialog.FileName }, _activeGroupId);
    }

    /// <summary>把整個資料夾（只取這一層）的簡報與圖片加進目前分組。</summary>
    private void ImportFolder_Click(object sender, RoutedEventArgs e)
    {
        // 每次從「本機」根開始，可瀏覽 C/D 槽、外接碟；不沿用設定頁或上次選過的位置。
        string? selected = PickFolder(startAtComputerRoot: true);
        if (string.IsNullOrEmpty(selected))
        {
            return;
        }

        var paths = Directory.GetFiles(selected)
            .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (paths.Count == 0)
        {
            Log("資料夾裡沒有檔案。");
            return;
        }

        ImportPaths(paths, _activeGroupId);
    }

    private string? PickFolder(string? initialFolder = null, bool startAtComputerRoot = false)
    {
        try
        {
            return FolderPicker.Pick(this,
                "選擇素材資料夾（只匯入這一層的 PPT 與圖片）",
                initialFolder, startAtComputerRoot);
        }
        catch (Exception ex)
        {
            Log("開啟資料夾選擇窗失敗：" + ex.Message);
            return null;
        }
    }

    private void SourcePanel_DragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
            SourceScroll.Background = ThemeManager.Brush("Panel2Brush");
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    private void SourcePanel_DragLeave(object sender, DragEventArgs e)
    {
        SourceScroll.Background = Brushes.Transparent;
        e.Handled = true;
    }

    private void SourcePanel_Drop(object sender, DragEventArgs e)
    {
        SourceScroll.Background = Brushes.Transparent;

        if (!e.Data.GetDataPresent(DataFormats.FileDrop)
            || e.Data.GetData(DataFormats.FileDrop) is not string[] dropped)
        {
            e.Handled = true;
            return;
        }

        var files = ExpandDroppedPaths(dropped);
        if (files.Count == 0)
        {
            Log("拖放加入素材：沒有可支援的 PowerPoint 或圖片檔案。");
            e.Handled = true;
            return;
        }

        Log($"拖放加入素材：{files.Count} 個檔案，放入「{GroupNameOf(_activeGroupId)}」。");
        ImportPaths(files, _activeGroupId);
        e.Handled = true;
    }

    /// <summary>展開拖放進來的檔案／資料夾；資料夾只取第一層支援的素材，避免誤拖整個磁碟。</summary>
    private static List<string> ExpandDroppedPaths(IEnumerable<string> paths)
    {
        var files = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void Add(string path)
        {
            if ((!IsPresentation(path) && !IsImageFile(path)) || !seen.Add(path))
            {
                return;
            }

            files.Add(path);
        }

        foreach (string path in paths)
        {
            try
            {
                if (File.Exists(path))
                {
                    Add(path);
                }
                else if (Directory.Exists(path))
                {
                    foreach (string file in Directory.EnumerateFiles(path, "*", SearchOption.TopDirectoryOnly))
                    {
                        Add(file);
                    }
                }
            }
            catch
            {
                // 個別檔案或資料夾無法讀取時略過，其餘照常加入
            }
        }

        return files;
    }

    /// <summary>
    /// 添加簡報素材前暫停主 PPT 放映：同一個 PowerPoint 實例在放映中不能再開其他簡報。
    /// 暫停前把目前頁面做成靜態圖蓋在投影幕上，觀眾不會看到桌面或編輯視窗。
    /// </summary>
    private void SuspendShowForImport()
    {
        if (_presentationImportScope)
        {
            return;
        }

        _presentationImportScope = true;
        _pptShowWasRunningForImport = _ppt.IsShowRunning;
        _ppt.KeepApplicationMinimized = true;

        // 不論是否在放映，都先記住 PowerPoint 編輯視窗，匯出期間用 Win32 隱藏。
        _ppt.CaptureEditorWindow();

        if (!_pptShowWasRunningForImport)
        {
            // 沒有放映時，匯出簡報仍會把 PowerPoint 編輯視窗叫出來；直接藏住。
            _ppt.MinimizeApplicationWindow();
            _ppt.HideEditorWindow();
            return;
        }

        _pptSuspendedForImport = true;
        try
        {
            _pptSuspendForImportSlide = Math.Max(1, _ppt.CurrentSlide);
        }
        catch
        {
            _pptSuspendForImportSlide = 1;
        }

        _pptSuspendForImportWasProjected = _projectedSource?.Kind == "main";

        if (_pptSuspendForImportWasProjected && _target is not null)
        {
            try
            {
                string folder = Path.Combine(Path.GetTempPath(), "ProjectionManager");
                Directory.CreateDirectory(folder);
                string file = Path.Combine(folder,
                    $"import_freeze_{DateTime.Now:yyyyMMdd_HHmmss_fff}.png");

                if (_ppt.ExportSlideTo(_pptSuspendForImportSlide, file, _target.Width, _target.Height))
                {
                    _importFreezeImage = file;
                    _importFreezeWindow = new MaterialWindow();
                    _importFreezeWindow.ShowPage(_target, file);
                }
            }
            catch (Exception ex)
            {
                Log("添加素材：建立暫時投影畫面失敗：" + ex.Message);
            }
        }

        // 先記住 PowerPoint 編輯視窗的 handle；結束放映後立刻用 Win32 隱藏，
        // 不依賴 COM 的 WindowState，因為 PowerPoint 會把編輯視窗重新叫回前景。
        _ppt.CaptureEditorWindow();
        _ppt.KeepApplicationMinimized = true;
        _ppt.EndShow();
        _ppt.MinimizeApplicationWindow();
        _ppt.HideEditorWindow();
        _repositionUntil = 0;

        // 結束放映後 PowerPoint 編輯視窗會跳到前景；暫時把控制台置頂並搶回焦點，
        // 避免添加素材時控制台被 PowerPoint 蓋住。
        _topmostBeforeImport = Topmost;
        Topmost = true;
        BringConsoleToFront();

        Log($"添加素材：已暫停主 PPT 放映（第 {_pptSuspendForImportSlide} 頁），完成後自動恢復。");
    }

    private void BringConsoleToFront()
    {
        try
        {
            Activate();
            Native.SetForegroundWindow(new WindowInteropHelper(this).Handle);
        }
        catch
        {
            // 搶焦點失敗不影響匯入
        }
    }

    /// <summary>添加素材結束後恢復主 PPT；若原本在背景，維持背景並遮黑。</summary>
    private void ResumeShowAfterImport()
    {
        if (!_presentationImportScope)
        {
            return;
        }

        bool resumeShow = _pptSuspendedForImport;
        _presentationImportScope = false;
        _pptSuspendedForImport = false;
        int slide = _pptSuspendForImportSlide;
        bool wasProjected = _pptSuspendForImportWasProjected;

        try
        {
            if (resumeShow && _target is not null && _ppt.HasLivePresentation)
            {
                bool windowed = WindowedShowCheck.IsChecked == true || _testMode;

                try
                {
                    // PowerPoint 視窗仍保持最小化，Run() 直接開放映，不會把編輯視窗帶到前景。
                    _ppt.StartShow(_target, slide, windowed);
                }
                catch
                {
                    // 少數版本在編輯視窗隱藏或最小化時無法 Run；先以最小化不搶焦點顯示，再試一次。
                    _ppt.ShowEditorWindowMinimized();
                    _ppt.StartShow(_target, slide, windowed);
                }

                _repositionUntil = Environment.TickCount64 + 3000;

                if (wasProjected)
                {
                    _projectedSource = MainSource;
                    _lastMainPage = slide;
                    Log($"已恢復主 PPT 放映：第 {slide} 頁。");
                }
                else
                {
                    // 原本投影的是素材／經文／網頁：主 PPT 只在背景待命，不要蓋到目前投影。
                    _ppt.SetScreenState(PowerPointService.StateBlackScreen);
                    Log($"已恢復主 PPT 背景放映：第 {slide} 頁（目前投影來源不變）。");
                }
            }
        }
        catch (Exception ex)
        {
            if (wasProjected && _projectedSource?.Kind == "main")
            {
                _projectedSource = null;
            }
            Log("添加素材後恢復主 PPT 放映失敗：" + ex.Message);
        }
        finally
        {
            _ppt.KeepApplicationMinimized = false;
            _ppt.ShowEditorWindowMinimized();

            try
            {
                _importFreezeWindow?.HideWindow();
                _importFreezeWindow?.Close();
            }
            catch
            {
                // 關閉暫時畫面失敗不影響匯入結果
            }

            _importFreezeWindow = null;

            if (_importFreezeImage is not null)
            {
                try
                {
                    File.Delete(_importFreezeImage);
                }
                catch
                {
                    // 檔案仍被快取佔用時留在暫存
                }
                _importFreezeImage = null;
            }

            Topmost = _topmostBeforeImport;
            BringConsoleToFront();
        }
    }

    private void ImportPaths(IReadOnlyList<string> paths, string groupId)
    {
        if (_importing)
        {
            Log("上一批素材還在匯入，請稍候。");
            return;
        }

        // 沒有指定分組、或分組已被刪除，一律落到預設分組，避免素材憑空消失。
        groupId = NormalizeGroupId(groupId);

        int width = _target?.Width ?? 1920;
        int height = _target?.Height ?? 1080;

        _importing = true;
        _importCts = new CancellationTokenSource();
        Cursor = Cursors.Wait;
        ImportMaterialButton.IsEnabled = false;
        ImportFolderButton.IsEnabled = false;
        SetImportBusy(true);

        var dialog = new ImportProgressDialog { Owner = this, Topmost = true };
        _importDialog = dialog;
        dialog.Closed += (_, _) =>
        {
            if (dialog.Cancelled)
            {
                _importCts?.Cancel();
            }
        };
        dialog.SetIndeterminate("正在準備匯入清單…");
        dialog.Show();

        // 匯入期間鎖住控制台內容，只留下進度視窗的「取消添加」可以操作，
        // 避免使用者在 COM 匯出途中又觸發另一個 PowerPoint 動作。
        RootGrid.IsEnabled = false;

        int ok = 0;
        int skipped = 0;
        int failed = 0;
        bool cancelled = false;
        MaterialItem? firstImported = null;
        string? lastError = null;

        try
        {
            if (paths.Any(IsPresentation))
            {
                SuspendShowForImport();
            }

            string groupName = GroupNameOf(groupId);
            Log($"開始匯入到分組「{groupName}」：共 {paths.Count} 個檔案（簡報以 {width}×{height} 逐頁轉圖）。");

            for (int i = 0; i < paths.Count; i++)
            {
                if (_importCts.IsCancellationRequested)
                {
                    cancelled = true;
                    break;
                }

                string path = paths[i];
                string fileName = Path.GetFileName(path);
                dialog.SetProgress($"正在處理第 {i + 1} / {paths.Count} 個：{fileName}", i, paths.Count);
                PumpImportUi();

                if (!IsPresentation(path) && !IsImageFile(path))
                {
                    skipped++;
                    continue;
                }

                if (_materials.Any(m => string.Equals(m.SourcePath, path, StringComparison.OrdinalIgnoreCase)))
                {
                    skipped++;
                    continue;
                }

                try
                {
                    MaterialItem item = IsImageFile(path)
                        ? ImportImageFile(path, groupId)
                        : ImportPresentationFile(path, groupId, width, height, dialog);

                    _materials.Add(item);
                    firstImported ??= item;
                    ok++;
                }
                catch (OperationCanceledException)
                {
                    cancelled = true;
                    break;
                }
                catch (Exception ex)
                {
                    failed++;
                    lastError = $"「{fileName}」導入失敗：{ex.Message}";
                    Log(lastError);
                }
            }

            SaveLibrary();
            RebuildSources();

            if (cancelled)
            {
                Log($"已取消添加素材：完成 {ok}、略過 {skipped}、失敗 {failed}。已完成的素材會保留。");
            }
            else
            {
                Log($"匯入完成：成功 {ok}、略過 {skipped}（重複或不支援）、失敗 {failed}。已放入「{groupName}」。");
            }
        }
        finally
        {
            // 這批匯入結束就回到預設分組；下次要放進其他分組，必須先在左欄點選該分組。
            _activeGroupId = MaterialGroup.DefaultId;

            RootGrid.IsEnabled = true;
            if (dialog.IsVisible)
            {
                dialog.Close();
            }

            _importDialog = null;
            _importCts?.Dispose();
            _importCts = null;

            SetImportBusy(false);
            Cursor = Cursors.Arrow;
            ImportMaterialButton.IsEnabled = true;
            ImportFolderButton.IsEnabled = true;
            _importing = false;

            ResumeShowAfterImport();

            if (ok > 0)
            {
                // 匯入後確保新素材一定看得見：離開搜尋結果、展開目標分組。
                _collapsedGroups.Remove(groupId);
                _settings.CollapsedGroups = _collapsedGroups.ToList();
                SettingsStore.Save(_settings);
                ClearSourceSearchForImport();
            }

            RenderAll();

            if (ok == 0 && failed > 0 && lastError is not null)
            {
                MessageBox.Show(this, "添加素材失敗。\n\n" + lastError, "添加素材",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (ok > 0 && firstImported is not null)
            {
                Dispatcher.BeginInvoke(new Action(() => BringMaterialIntoView(firstImported)),
                    System.Windows.Threading.DispatcherPriority.Background);
            }
        }
    }

    private void ClearSourceSearchForImport()
    {
        _sourceSearchKeyword = string.Empty;
        if (SourceSearchBox is not null)
        {
            SourceSearchBox.Text = string.Empty;
        }
        if (SourceSearchPanel is not null)
        {
            SourceSearchPanel.Visibility = Visibility.Collapsed;
            SourceSearchToggle.Content = "查找";
            SourceSearchStatus.Text = "輸入關鍵字後按「查找」";
        }
    }

    private void BringMaterialIntoView(MaterialItem material)
    {
        var source = _sources.FirstOrDefault(s => s.Kind == "material"
                                                   && s.Material is not null
                                                   && s.Material.Id == material.Id);
        if (source is null)
        {
            return;
        }

        var card = SourcePanel.Children
            .OfType<Button>()
            .FirstOrDefault(b => ReferenceEquals(b.Tag, source));
        card?.BringIntoView();
    }

    private string NormalizeGroupId(string? groupId)
    {
        if (string.IsNullOrWhiteSpace(groupId))
        {
            return MaterialGroup.DefaultId;
        }

        return _groups.Any(g => g.Id == groupId) ? groupId : MaterialGroup.DefaultId;
    }

    /// <summary>
    /// 匯入在 UI 執行緒進行（PowerPoint COM 的要求），每處理一頁就抽一次訊息迴圈：
    /// 進度條會動、取消按鈕可以按，但不會同時執行第二個匯入。
    /// </summary>
    private void PumpImportUi()
    {
        var frame = new DispatcherFrame();
        Dispatcher.CurrentDispatcher.BeginInvoke(
            DispatcherPriority.Background,
            new Action(() => frame.Continue = false));
        Dispatcher.PushFrame(frame);
    }

    private MaterialItem ImportPresentationFile(
        string path,
        string groupId,
        int width,
        int height,
        ImportProgressDialog? progressDialog = null)
    {
        string name = Path.GetFileNameWithoutExtension(path);
        string folder = MaterialStore.NewImageFolder(name);

        try
        {
            int count = _ppt.ExportSlides(path, folder, width, height,
                (page, total) =>
                {
                    if (progressDialog is not null)
                    {
                        progressDialog.SetProgress(
                            $"添加素材中：{name}　第 {page} / {total} 頁（請等待進度完成）",
                            page, total);
                    }

                    PumpImportUi();

                    if (_importCts?.IsCancellationRequested == true)
                    {
                        throw new OperationCanceledException("使用者取消添加素材。");
                    }
                },
                () => _importCts?.IsCancellationRequested == true);

            var pages = Directory.GetFiles(folder, "p*.png")
                .OrderBy(p => p, StringComparer.Ordinal)
                .ToList();

            if (pages.Count == 0)
            {
                throw new InvalidOperationException($"「{name}」沒有匯出任何頁面。");
            }

            Log($"已導入「{name}」：{count} 頁。");

            return new MaterialItem
            {
                Name = name,
                SourcePath = path,
                Folder = folder,
                Pages = pages,
                GroupId = groupId,
            };
        }
        catch (OperationCanceledException)
        {
            // 取消時刪掉只匯出一半的快取資料夾，不留下半份素材。
            try
            {
                Directory.Delete(folder, true);
            }
            catch
            {
                // 檔案仍被 PowerPoint 佔用時留在暫存，不影響下一次匯入
            }

            throw;
        }
    }

    private MaterialItem ImportImageFile(string path, string groupId)
    {
        string name = Path.GetFileNameWithoutExtension(path);
        string folder = MaterialStore.NewImageFolder(name);
        string extension = Path.GetExtension(path).ToLowerInvariant();
        string destination = Path.Combine(folder, "p001" + extension);
        File.Copy(path, destination, true);

        Log($"已導入圖片「{name}」。");

        return new MaterialItem
        {
            Name = name,
            SourcePath = path,
            Folder = folder,
            Pages = new List<string> { destination },
            GroupId = groupId,
        };
    }

    private void SaveLibrary()
    {
        MaterialStore.Save(new MaterialLibrary { Items = _materials, Groups = _groups });
    }

    private string GroupNameOf(string groupId) =>
        _groups.FirstOrDefault(g => g.Id == groupId)?.Name ?? "默認分組";

    private void AddGroup_Click(object sender, RoutedEventArgs e)
    {
        string? name = PromptText("新增素材庫分組", "分組名稱（例：詩歌、經文、報告）", "新分組");
        if (string.IsNullOrWhiteSpace(name))
        {
            return;
        }

        name = name.Trim();
        if (_groups.Any(g => string.Equals(g.Name, name, StringComparison.OrdinalIgnoreCase)))
        {
            Log($"已有同名分組「{name}」。");
            return;
        }

        var group = new MaterialGroup { Name = name };
        _groups.Add(group);
        _activeGroupId = group.Id;
        SaveLibrary();
        RenderSources();
        Log($"已新增分組「{name}」，此後匯入的素材會放進這一組。");
    }

    private void RenameGroup(MaterialGroup group)
    {
        if (group.IsDefault)
        {
            Log("「默認分組」是固定名稱，不能重新命名。");
            return;
        }

        string? name = PromptText("重新命名分組", "新名稱", group.Name);
        if (string.IsNullOrWhiteSpace(name))
        {
            return;
        }

        name = name.Trim();
        if (_groups.Any(g => g.Id != group.Id && string.Equals(g.Name, name, StringComparison.OrdinalIgnoreCase)))
        {
            Log($"已有同名分組「{name}」。");
            return;
        }

        string oldName = group.Name;
        group.Name = name;
        SaveLibrary();
        RenderAll();
        Log($"分組已改名：{oldName} → {name}。");
    }

    private void DeleteGroup(MaterialGroup group)
    {
        if (group.IsDefault)
        {
            Log("「默認分組」是預設分組，不能刪除。");
            return;
        }

        var result = MessageBox.Show(
            this,
            $"刪除分組「{group.Name}」？\n組裡的 {_materials.Count(m => m.GroupId == group.Id)} 份素材會移到「默認分組」，不會被刪掉。",
            "刪除分組", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        foreach (var material in _materials.Where(m => m.GroupId == group.Id))
        {
            material.GroupId = MaterialGroup.DefaultId;
        }

        _groups.Remove(group);
        if (_activeGroupId == group.Id)
        {
            _activeGroupId = MaterialGroup.DefaultId;
        }

        SaveLibrary();
        RebuildSources();
        RenderAll();
        Log($"已刪除分組「{group.Name}」，組內素材已移到「默認分組」。");
    }

    /// <summary>簡單的文字輸入窗（WPF 沒有內建 InputBox）。</summary>
    private string? PromptText(string title, string label, string initial)
    {
        var window = new Window
        {
            Title = title,
            Width = 380,
            Height = 180,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            ResizeMode = ResizeMode.NoResize,
            Background = ThemeManager.Brush("Panel2Brush"),
            ShowInTaskbar = false,
        };

        var panel = new StackPanel { Margin = new Thickness(18) };
        panel.Children.Add(new TextBlock
        {
            Text = label,
            Foreground = DimBrush,
            FontSize = 12,
            Margin = new Thickness(0, 0, 0, 8),
        });

        var box = new TextBox
        {
            Text = initial,
            Height = 30,
            FontSize = 13,
            Background = ThemeManager.Brush("InputBrush"),
            Foreground = ThemeManager.Brush("TextPrimaryBrush"),
            BorderBrush = LineBrush,
            Padding = new Thickness(8, 4, 8, 4),
        };
        panel.Children.Add(box);

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0),
        };

        string? result = null;
        var ok = new Button { Content = "確定", Width = 84, Height = 30, Margin = new Thickness(0, 0, 8, 0) };
        var cancel = new Button { Content = "取消", Width = 84, Height = 30 };
        ok.Click += (_, _) => { result = box.Text; window.DialogResult = true; };
        cancel.Click += (_, _) => window.DialogResult = false;
        buttons.Children.Add(ok);
        buttons.Children.Add(cancel);
        panel.Children.Add(buttons);

        window.Content = panel;
        window.Loaded += (_, _) => { box.Focus(); box.SelectAll(); };

        return window.ShowDialog() == true ? result : null;
    }

    private void SetImportBusy(bool busy)
    {
        ImportMaterialButton.Content = busy ? "匯入中…" : "＋ 素材";
        AddGroupButton.IsEnabled = !busy;
    }

    private void ReexportMaterial_Click(object sender, RoutedEventArgs e)
    {
        var item = _selected?.Material;
        if (item is null)
        {
            Log("請先在左欄選擇一份素材。");
            return;
        }

        if (item.IsImage)
        {
            Log("這是圖片素材，沒有需要重新匯出的頁面。要換圖請「移除」後重新匯入。");
            return;
        }

        if (!File.Exists(item.SourcePath))
        {
            Log($"找不到原始檔：{item.SourcePath}");
            return;
        }

        int width = _target?.Width ?? 1920;
        int height = _target?.Height ?? 1080;
        Cursor = Cursors.Wait;

        try
        {
            string oldFolder = item.Folder;
            string folder = MaterialStore.NewImageFolder(item.Name);
            int count = _ppt.ExportSlides(item.SourcePath, folder, width, height);

            item.Pages = Directory.GetFiles(folder, "p*.png")
                .OrderBy(p => p, StringComparer.Ordinal)
                .ToList();
            item.Folder = folder;
            item.Added = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

            if (!string.IsNullOrEmpty(oldFolder) && oldFolder != folder)
            {
                try
                {
                    Directory.Delete(oldFolder, true);
                }
                catch
                {
                    // 舊資料夾被佔用就留著，不影響新頁面
                }
            }

            SaveLibrary();
            if (_live?.Material is not null)
            {
                _live.PageIndex = Math.Clamp(_live.PageIndex, 0, Math.Max(0, _live.Material.Pages.Count - 1));
                if (_materialProjectionActive)
                {
                    ShowMaterialPage();
                }
            }

            Log($"已重新匯出「{item.Name}」，共 {count} 頁。");
            RenderAll();
        }
        catch (Exception ex)
        {
            Log(ex.Message);
        }
        finally
        {
            Cursor = Cursors.Arrow;
        }
    }

    private void RemoveMaterial_Click(object sender, RoutedEventArgs e)
    {
        var item = _selected?.Material;
        if (item is null)
        {
            Log("請先在左欄選擇一份素材。");
            return;
        }

        var result = MessageBox.Show(this, $"移除素材「{item.Name}」？原始簡報檔不會被刪除。",
            "移除素材", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        bool wasLive = _live?.Material?.Id == item.Id;
        if (wasLive)
        {
            _live = _sources[0];
            _materialWindow?.HideWindow();
        }

        _materials.Remove(item);
        SaveLibrary();
        MaterialStore.DeleteFolder(item);
        _selected = null;

        RebuildSources();
        Log($"已移除素材「{item.Name}」。");
        RenderAll();
    }

    // ---------------------------------------------------------------- 控制權

    private void HandoverButton_Click(object sender, RoutedEventArgs e)
    {
        if (_mode == ControlMode.Operator)
        {
            _mode = ControlMode.Speaker;
            _handoverSlide = _ppt.IsShowRunning ? _ppt.CurrentSlide : 0;

            if (_handoverSlide > 0)
            {
                Log($"已將控制權交給講員（交出頁：第 {_handoverSlide} 頁）。操作員的按鈕仍可用，鼠標與鍵盤不影響放映。");
            }
            else
            {
                Log("已將控制權交給講員（目前尚未開始放映，交出頁未記錄）。");
            }
        }
        else
        {
            _mode = ControlMode.Operator;
            int current = _ppt.IsShowRunning ? _ppt.CurrentSlide : 0;

            if (_handoverSlide > 0 && current > 0)
            {
                Log($"已收回控制權。交出時第 {_handoverSlide} 頁，現在第 {current} 頁（差 {current - _handoverSlide} 頁）。");
            }
            else
            {
                Log("已收回控制權，交還操作員。");
            }
        }

        ApplyMode();
        UpdateShowStatus();
    }

    private void ReturnSlideButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_ppt.IsShowRunning || _handoverSlide <= 0)
        {
            Log("沒有可回跳的交出頁記錄。");
            return;
        }

        try
        {
            if (!_ppt.TryGotoSlide(_handoverSlide, out string error))
            {
                Log(error);
                return;
            }

            _black = false;
            Log($"已跳回交出頁：第 {_handoverSlide} 頁。");
        }
        catch (PowerPointException ex)
        {
            Log(ex.Message);
        }

        UpdateShowStatus();
    }

    private void ApplyMode()
    {
        bool operatorInControl = _mode == ControlMode.Operator;

        ControllerText.Text = operatorInControl ? "操作員" : "講員";
        ControllerText.Foreground = operatorInControl ? OperatorBrush : SpeakerBrush;
        HandoverButton.Content = operatorInControl ? "交給講員" : "收回控制";

        ControllerChipText.Text = operatorInControl ? "操作員" : "講員";
        ControllerChipText.Foreground = operatorInControl ? OperatorBrush : SpeakerBrush;
        ControllerChip.BorderBrush = operatorInControl ? OperatorBrush : SpeakerBrush;
        ControllerChip.Background = ThemeManager.Brush("AccentSoftBrush");

        ControllerHint.Text = operatorInControl
            ? "由操作員推進。講員上台時按「交給講員」，僅切換標示與記錄，不鎖定任何一方。"
            : "講員以翻頁筆推進中。操作員的按鈕仍可用，鼠標與鍵盤不會影響放映頁面。";
    }

    // ---------------------------------------------------------------- 翻頁筆

    private void OnRawKey(KeyInput input)
    {
        bool pageKey = IsPageKey(input.VirtualKey);

        if (_rawInput is not null && _rawInput.ReceivedCount <= 3 && !pageKey)
        {
            Log($"診斷：收到原始按鍵 VK=0x{input.VirtualKey:X2}，裝置={Describe(input.DeviceName)}");
        }

        if (_binding)
        {
            _bindReceived++;
            Log($"綁定中收到 VK=0x{input.VirtualKey:X2}（裝置：{Describe(input.DeviceName)}）。");

            if (IsBindingKey(input.VirtualKey))
            {
                FinishBinding(input.DeviceName);
            }

            return;
        }

        bool isPen = IsBoundDevice(input.DeviceName);

        if (!isPen)
        {
            if (pageKey)
            {
                string dev = Describe(input.DeviceName);
                Log($"收到翻頁鍵 VK=0x{input.VirtualKey:X2}，但它不屬於已綁定的翻頁筆" +
                    $"（本次裝置：{(string.IsNullOrEmpty(dev) ? "無法識別" : dev)}），已忽略。");
            }

            return;
        }

        if (!pageKey)
        {
            return;
        }

        bool forward = IsForwardKey(input.VirtualKey);
        if (!TryBeginNavigation())
        {
            return;
        }

        Log($"翻頁筆 → {(forward ? "下一頁" : "上一頁")}（VK=0x{input.VirtualKey:X2}）。");
        Navigate(forward);
    }

    /// <summary>
    /// 放映防護層吞掉 PageUp/PageDown 後，由這裡執行唯一一次導航。
    /// 若不接手，按鍵被吞後 Raw Input 也收不到，會出現「綁定成功但 PgUp/PgDn 沒反應」。
    /// </summary>
    private void OnGuardedPageKey(int vk)
    {
        Dispatcher.BeginInvoke(new Action(() =>
        {
            if (_binding || string.IsNullOrWhiteSpace(_settings.BoundDeviceName))
            {
                return;
            }

            bool forward = vk == Native.VK_NEXT;
            if (!TryBeginNavigation())
            {
                return;
            }

            Log($"翻頁筆（放映防護） → {(forward ? "下一頁" : "上一頁")}（VK=0x{vk:X2}）。");
            try
            {
                Navigate(forward);
            }
            catch (Exception ex)
            {
                Log("放映防護導航失敗：" + ex.Message);
            }
        }));
    }

    /// <summary>同一次按鍵可能同時由 Raw Input 與 WPF KeyDown 送達；只放行第一個事件。</summary>
    private bool TryBeginNavigation()
    {
        long now = Environment.TickCount64;
        if (now - _lastArrowNavAt < 100)
        {
            return false;
        }

        _lastArrowNavAt = now;
        return true;
    }

    /// <summary>綁定期間按任何非修飾鍵都可完成綁定；自帶鍵盤、簡報筆都能用。</summary>
    private static bool IsBindingKey(int vk) =>
        vk is not (0x10 or 0x11 or 0x12 or 0x14 or 0x5B or 0x5C);  // Shift/Ctrl/Alt/Caps/Win

    /// <summary>翻頁筆常見按鍵：PageUp/PageDown、方向鍵、空白鍵。綁定時全部接受，避免實際筆只送方向鍵而綁不上。</summary>
    private static bool IsPageKey(int vk) =>
        vk is Native.VK_PRIOR or Native.VK_NEXT
            or Native.VK_LEFT or Native.VK_RIGHT
            or Native.VK_UP or Native.VK_DOWN
            or Native.VK_SPACE
            or Native.VK_MEDIA_NEXT_TRACK or Native.VK_MEDIA_PREV_TRACK;

    private static bool IsForwardKey(int vk) =>
        vk is Native.VK_NEXT or Native.VK_RIGHT or Native.VK_DOWN or Native.VK_SPACE
            or Native.VK_MEDIA_NEXT_TRACK;

    private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e)
    {
        // 控制台在放映期間可能被 PowerPoint 壓在前景；使用者主動點擊時一定要能取回焦點。
        if (!IsActive)
        {
            Activate();
            IntPtr hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd != IntPtr.Zero)
            {
                Native.SetForegroundWindow(hwnd);
            }
        }
    }

    /// <summary>
    /// 主視窗方向鍵：← 上一頁、→ 下一頁。焦點在搜尋框／跳頁框／下拉選單時不攔截，
    /// 避免打字與選項操作被翻頁搶走。
    /// </summary>
    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (_binding || e.Key is not (Key.Left or Key.Right))
        {
            return;
        }

        if (Keyboard.Modifiers != ModifierKeys.None)
        {
            return;
        }

        if (Keyboard.FocusedElement is TextBoxBase or PasswordBox or ComboBox or ComboBoxItem)
        {
            return;
        }

        // 翻頁筆若同時送出 Raw Input 與 WPF KeyDown，只讓第一個事件推進一次。
        if (!TryBeginNavigation())
        {
            e.Handled = true;
            return;
        }

        e.Handled = true;
        Log($"鍵盤方向鍵 → {(e.Key == Key.Right ? "下一頁" : "上一頁")}。");
        Navigate(e.Key == Key.Right);
    }

    private bool IsBoundDevice(string deviceName)
    {
        string bound = _settings.BoundDeviceName ?? string.Empty;
        if (string.IsNullOrWhiteSpace(bound) || string.IsNullOrWhiteSpace(deviceName))
        {
            return false;
        }

        if (string.Equals(bound, deviceName, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        // 次要比對：USB 埠不同會導致裝置路徑改變，改以 VID/PID 比對
        string key = DeviceIdKey(deviceName);
        return key.Length > 0 && string.Equals(key, DeviceIdKey(bound), StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 從裝置路徑取出「裝置身分」（VID + PID），用來跨 USB 埠、跨重新列舉比對同一支筆。
    /// USB：\\?\HID#VID_046D&amp;PID_C52B#…#{guid}；BLE：…Dev_VID&amp;02045e&amp;PID&amp;0237…；
    /// ACPI：沒有 VID/PID，只能靠完整路徑比對。
    /// </summary>
    private static string DeviceIdKey(string deviceName)
    {
        string vid = ReadHex(deviceName, "vid_", "vid&");
        string pid = ReadHex(deviceName, "pid_", "pid&");

        if (vid.Length == 0 && pid.Length == 0)
        {
            return string.Empty;
        }

        return ("vid_" + vid + "&pid_" + pid).ToLowerInvariant();
    }

    private static string ReadHex(string source, string markerA, string markerB)
    {
        int start = source.IndexOf(markerA, StringComparison.OrdinalIgnoreCase);
        if (start < 0)
        {
            start = source.IndexOf(markerB, StringComparison.OrdinalIgnoreCase);
        }

        if (start < 0)
        {
            return string.Empty;
        }

        start += markerA.Length;
        int end = start;
        while (end < source.Length && end - start < 8 && Uri.IsHexDigit(source[end]))
        {
            end++;
        }

        return source.Substring(start, end - start);
    }

    private static string Describe(string deviceName)
    {
        string key = DeviceIdKey(deviceName);
        return key.Length > 0 ? key.ToUpperInvariant() : deviceName;
    }

    private void BindButton_Click(object sender, RoutedEventArgs e)
    {
        _binding = true;
        _bindReceived = 0;
        if (_guard is not null)
        {
            _guard.Suspended = true;  // 綁定期間不可吞 PageUp/PageDown
        }
        BindButton.Content = "請按一下翻頁筆…";
        BoundDeviceText.Text = "等待輸入（30 秒內，按任意鍵）";
        BindButton.IsEnabled = false;

        _bindTimer.Stop();
        _bindTimer.Interval = TimeSpan.FromSeconds(30);
        _bindTimer.Start();

        Log("綁定模式：請按一下翻頁筆的下一頁鍵（或任何按鍵）。");
    }

    private void BindTimer_Tick(object? sender, EventArgs e)
    {
        _bindTimer.Stop();
        if (!_binding)
        {
            return;
        }

        _binding = false;
        if (_guard is not null)
        {
            _guard.Suspended = false;
        }
        BindButton.Content = "重新綁定翻頁筆";
        BindButton.IsEnabled = true;
        BoundDeviceText.Text = string.IsNullOrEmpty(_settings.BoundDeviceName)
            ? "尚未綁定"
            : ShortDevice(Describe(_settings.BoundDeviceName!), 26);
        BoundDeviceText.ToolTip = string.IsNullOrEmpty(_settings.BoundDeviceName)
            ? null
            : Describe(_settings.BoundDeviceName!);
        Log($"綁定逾時：這 30 秒內收到 {_bindReceived} 個按鍵（累計 {_rawInput?.ReceivedCount ?? 0} 個）。" +
            "請確認翻頁筆已開機、接收器已插好，再按一次「綁定翻頁筆」並按任意鍵。");
        UpdateNavHint();
    }

    private void FinishBinding(string deviceName)
    {
        _binding = false;
        if (_guard is not null)
        {
            _guard.Suspended = false;
        }
        _bindTimer.Stop();
        BindButton.IsEnabled = true;

        if (string.IsNullOrEmpty(deviceName))
        {
            BindButton.Content = "重新綁定翻頁筆";
            BoundDeviceText.Text = "無法識別裝置，請重試";
            Log("無法取得裝置路徑，綁定失敗。");
            UpdateNavHint();
            return;
        }

        _settings.BoundDeviceName = deviceName;
        try
        {
            SettingsStore.Save(_settings);
        }
        catch (Exception ex)
        {
            Log("綁定已生效，但儲存設定失敗：" + ex.Message);
        }

        BindButton.Content = "重新綁定翻頁筆";
        BoundDeviceText.Text = ShortDevice(Describe(deviceName), 26);
        BoundDeviceText.ToolTip = Describe(deviceName);
        Log("已綁定翻頁筆：" + Describe(deviceName));
        UpdateNavHint();
    }

    private void UnbindButton_Click(object sender, RoutedEventArgs e)
    {
        _settings.BoundDeviceName = null;
        SettingsStore.Save(_settings);

        BindButton.Content = "綁定翻頁筆";
        BoundDeviceText.Text = "尚未綁定";
        Log("已解除綁定：此後任何裝置的按鍵都不再推進，只有控制台按鈕有效。");
        UpdateNavHint();
    }

    private void UpdateNavHint()
    {
        if (_binding)
        {
            return;   // 綁定中由綁定流程顯示「請按一下翻頁筆…」，不要被 300ms 狀態刷新蓋掉
        }

        if (string.IsNullOrEmpty(_settings.BoundDeviceName))
        {
            NavHintText.Text = "翻頁筆未綁定";
            NavHintText.ToolTip = null;
            BoundDeviceText.Text = "尚未綁定";
            BoundDeviceText.ToolTip = null;
            return;
        }

        string full = Describe(_settings.BoundDeviceName!);
        string shortName = ShortDevice(full, 26);
        NavHintText.Text = "翻頁筆：已綁定 · " + ShortDevice(full, 18);
        NavHintText.ToolTip = "已綁定翻頁筆：" + full;
        BoundDeviceText.Text = shortName;
        BoundDeviceText.ToolTip = full;
    }

    /// <summary>裝置路徑可能是一長串沒有空格的 ACPI 字串；介面只顯示前段，完整值放 ToolTip。</summary>
    private static string ShortDevice(string value, int max)
    {
        if (string.IsNullOrEmpty(value) || value.Length <= max)
        {
            return value;
        }

        return value[..Math.Max(1, max - 1)] + "…";
    }

    // ---------------------------------------------------------------- 狀態

    private void StartTimers()
    {
        _statusTimer.Interval = TimeSpan.FromMilliseconds(300);
        _statusTimer.Tick += StatusTimer_Tick;
        _statusTimer.Start();

        _bindTimer.Tick += BindTimer_Tick;
    }

    private void StatusTimer_Tick(object? sender, EventArgs e)
    {
        // 放映可能不是我們結束的（講員在投影幕按 ESC）。不及時校正，
        // 控制台會停在「以為還在放映」的假狀態。
        if (_ppt.IsShowRunning && !_ppt.CheckShowAlive())
        {
            _black = false;
            _mode = ControlMode.Operator;
            _handoverSlide = 0;
            ApplyMode();
            ApplyConsoleActivation();
            Log("偵測到放映已結束（可能由講員按 ESC 或在 PowerPoint 端結束）。控制台已恢復可操作。");
        }

        if (_ppt.IsShowRunning && _live?.Kind == "main" && _target is not null
            && Environment.TickCount64 < _repositionUntil)
        {
            _ppt.Reposition(_target);
        }

        // 素材蓋在放映之上：PowerPoint 播放動畫時會重繪，定期把素材視窗拉回最上層
        if (_live?.Kind == "material" && !_black)
        {
            _materialWindow?.KeepOnTop();
        }

        UpdateShowStatus();

        if (_ppt.IsShowRunning && _ppt.CurrentSlide != _lastMainPage)
        {
            _lastMainPage = _ppt.CurrentSlide;
            if (_live?.Kind == "main")
            {
                // 只換預覽圖與高亮框：整條縮圖重建（最多 200 個元素）會讓每次翻頁都頓一下
                RenderPreview();
                UpdateThumbHighlight();
            }

            RenderSources();
        }
    }

    private void UpdateShowStatus()
    {
        var projected = _projectedSource;
        bool projectedActive = IsProjected(projected);

        if (projectedActive && projected != _live)
        {
            // 已選取其他來源，但螢幕 2 仍是上一份投影：狀態列必須以真正投出的來源為準。
            ShowStateText.Text = _black ? "黑屏" : $"{DescribeShortSource(projected)}投影中";
            ShowDot.Fill = _black ? SpeakerBrush : OperatorBrush;
            ShowProjectedPage(projected!);
        }
        else if (_ppt.IsShowRunning && _live?.Kind == "main")
        {
            int current = _ppt.CurrentSlide;
            int total = ShowTotal;

            if (_live?.Kind == "main")
            {
                PageText.Text = current > 0 ? current.ToString() : "—";
                PageTotalText.Text = total > 0 ? $"/ {total}" : "/ —";
            }

            ShowStateText.Text = _black ? "黑屏" : "放映中";
            ShowDot.Fill = _black ? SpeakerBrush : OperatorBrush;
        }
        else if (_live?.Kind == "scripture")
        {
            bool scripturePushed = IsProjected(_live);
            int pageCount = scripturePushed && _scriptureWindow?.IsVisible == true ? _scriptureWindow.PageCount : 0;
            PageText.Text = pageCount > 0 ? (_scriptureWindow!.CurrentIndex + 1).ToString() : "—";
            PageTotalText.Text = pageCount > 0 ? $"/ {pageCount}" : "/ —";
            ShowStateText.Text = _black
                ? "黑屏"
                : SingleScreen
                    ? "經文預覽（未投影）"
                    : scripturePushed ? "經文投影中" : "經文已選擇（待推送）";
            ShowDot.Fill = _black ? SpeakerBrush : OperatorBrush;
        }
        else if (_live?.Kind == "web")
        {
            PageText.Text = "—";
            PageTotalText.Text = "/ —";
            ShowStateText.Text = _black
                ? "黑屏"
                : SingleScreen
                    ? "網頁預覽（未投影）"
                    : IsProjected(_live) ? "網頁投影中" : "網頁已選擇（待推送）";
            ShowDot.Fill = _black ? SpeakerBrush : OperatorBrush;
        }
        else if (_live?.Kind == "zoom")
        {
            PageText.Text = "—";
            PageTotalText.Text = "/ —";
            ShowStateText.Text = _black
                ? "黑屏"
                : IsProjected(_live) ? "Zoom 投影中" : "Zoom 已選擇（待推送）";
            ShowDot.Fill = _black ? SpeakerBrush : OperatorBrush;
        }
        else if (_live?.Kind == "material")
        {
            var material = _live.Material;
            PageText.Text = material is null ? "—" : (_live.PageIndex + 1).ToString();
            PageTotalText.Text = material is null ? "/ —" : $"/ {material.Pages.Count}";
            ShowStateText.Text = _black
                ? "黑屏"
                : SingleScreen
                    ? "素材預覽（未投出）"
                    : IsProjected(_live) ? "素材投影中" : "素材已選擇（未投影）";
            ShowDot.Fill = _black ? SpeakerBrush : OperatorBrush;
        }
        else
        {
            PageText.Text = "—";
            PageTotalText.Text = "/ —";
            ShowStateText.Text = "未開始放映";
            ShowDot.Fill = FaintBrush;
        }

        SourceStatusText.Text = projectedActive
            ? $"螢幕 2：{DescribeSourceName(projected!)}"
            : "螢幕 2：未投影";

        ReturnSlideButton.IsEnabled = _handoverSlide > 0;
        UpdateNavHint();

        MaterialPanel.Visibility = _selected?.Material is not null ? Visibility.Visible : Visibility.Collapsed;
        MaterialDivider.Visibility = MaterialPanel.Visibility;
        PushMaterialButton.Visibility = _live?.Kind is "material" or "scripture" or "web" or "zoom"
            ? Visibility.Visible
            : Visibility.Collapsed;
        bool pushed = IsProjected(_live);
        PushMaterialButton.Content = pushed ? "斷開投影" : "推送投影";

        WebNavPanel.Visibility = _selected?.Kind == "web" ? Visibility.Visible : Visibility.Collapsed;
        ZoomPanel.Visibility = _selected?.Kind == "zoom" ? Visibility.Visible : Visibility.Collapsed;
        if (ExtraPushHint is not null)
        {
            ExtraPushHint.Text = pushed
                ? "已在螢幕 2 投影；翻頁會即時同步，按「斷開投影」可停止。"
                : projectedActive
                    ? $"螢幕 2 目前仍是「{DescribeShortSource(projected!)}」；按「推送投影」才會切換。"
                    : "尚未推送投影，按中間「推送投影」才會上螢幕 2。";
            ExtraPushHint.Foreground = pushed
                ? ThemeManager.Brush("TealBrush")
                : projectedActive
                    ? ThemeManager.Brush("AmberBrush")
                    : ThemeManager.Brush("AmberBrush");
        }

        if (_selected?.Kind == "zoom")
        {
            UpdateZoomPanel();
        }
        if (_selected?.Material is not null)
        {
            SelectedMaterialText.Text = $"{_selected.Material.Name}　（{_selected.Material.Pages.Count} 頁）";
        }
    }

    private void ShowProjectedPage(SourceItem source)
    {
        switch (source.Kind)
        {
            case "main":
            {
                int current = _ppt.CurrentSlide;
                PageText.Text = current > 0 ? current.ToString() : "—";
                PageTotalText.Text = ShowTotal > 0 ? $"/ {ShowTotal}" : "/ —";
                break;
            }
            case "material":
            {
                int total = source.Material?.Pages.Count ?? 0;
                PageText.Text = total > 0 ? (source.PageIndex + 1).ToString() : "—";
                PageTotalText.Text = total > 0 ? $"/ {total}" : "/ —";
                break;
            }
            case "scripture":
            {
                int total = _scriptureWindow?.IsVisible == true ? _scriptureWindow.PageCount : 0;
                PageText.Text = total > 0 ? (_scriptureWindow!.CurrentIndex + 1).ToString() : "—";
                PageTotalText.Text = total > 0 ? $"/ {total}" : "/ —";
                break;
            }
            default:
                PageText.Text = "—";
                PageTotalText.Text = "/ —";
                break;
        }
    }

    private static string DescribeShortSource(SourceItem source)
    {
        return source.Kind switch
        {
            "main" => "主 PPT",
            "material" => "素材",
            "scripture" => "經文",
            "web" => "網頁",
            "zoom" => "Zoom",
            _ => "投影源",
        };
    }

    // ---------------------------------------------------------------- 螢幕

    private void LoadMonitors()
    {
        _monitors.Clear();
        _monitors.AddRange(DisplayService.GetMonitors());

        // 測試模式或設定頁的「調試模式」：本機只有一台螢幕時，用右半螢幕模擬「螢幕 2」，
        // 讓投影目標走延伸螢幕流程，控制台留在左半，不必 Alt+Tab。
        if ((_testMode || _settings.DebugMode) && _monitors.Count < 2)
        {
            var primary = _monitors.FirstOrDefault();
            int left = primary?.Left ?? 0;
            int top = primary?.Top ?? 0;
            int width = (primary?.Width ?? 1920) / 2;
            int height = primary?.Height ?? 1080;

            _monitors.Add(new MonitorInfo
            {
                Handle = IntPtr.Zero,
                DeviceName = "\\\\.\\TEST2",
                Left = left + width,
                Top = top,
                Right = left + width * 2,
                Bottom = top + height,
                IsPrimary = false,
            });
        }

        MonitorCombo.Items.Clear();
        foreach (var monitor in _monitors)
        {
            MonitorCombo.Items.Add(monitor.ToString());
        }

        int index = PickDefaultMonitorIndex();

        // 測試模式與調試模式必須真的落在模擬的 TEST2，不能沿用上次儲存的實體螢幕設定，
        // 否則「調試模式」開了卻仍在主螢幕放映，等於沒有模擬到螢幕 2。
        int simulated = _monitors.FindIndex(m => m.DeviceName == "\\\\.\\TEST2");
        if (simulated >= 0)
        {
            index = simulated;
        }
        else if (!string.IsNullOrEmpty(_settings.TargetMonitorDevice))
        {
            int saved = _monitors.FindIndex(m => m.DeviceName == _settings.TargetMonitorDevice);
            if (saved >= 0)
            {
                index = saved;
            }
        }

        if (_monitors.Count > 0)
        {
            MonitorCombo.SelectedIndex = index;
        }

        UpdateTarget();
    }

    private int PickDefaultMonitorIndex()
    {
        int index = _monitors.FindIndex(m => !m.IsPrimary);
        return index >= 0 ? index : 0;
    }

    private void UpdateTarget()
    {
        int index = MonitorCombo.SelectedIndex;
        _target = index >= 0 && index < _monitors.Count ? _monitors[index] : _monitors.FirstOrDefault();

        if (_target is null)
        {
            TargetInfoText.Text = "投影目標：—";
            return;
        }

        TargetInfoText.Text = $"投影目標 {_target.DeviceName}\n{_target.Width}×{_target.Height}　位於 {_target.Left},{_target.Top}";

        _settings.TargetMonitorDevice = _target.DeviceName;
        SettingsStore.Save(_settings);
    }

    private void MonitorCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateTarget();

        if (_ready && _live?.Kind == "material" && _materialProjectionActive)
        {
            ShowMaterialPage();
        }
    }

    private void RefreshMonitors_Click(object sender, RoutedEventArgs e)
    {
        LoadMonitors();
        Log($"已重新整理螢幕，共 {_monitors.Count} 台。");
    }

    private void LockCursorCheck_Changed(object sender, RoutedEventArgs e)
    {
        if (LockCursorCheck.IsChecked == true)
        {
            var operatorMonitor = _monitors.FirstOrDefault(m => _target is not null && m.Handle != _target.Handle)
                                   ?? _monitors.FirstOrDefault();

            if (operatorMonitor is null)
            {
                Log("找不到作業螢幕，無法鎖定鼠標。");
                LockCursorCheck.IsChecked = false;
                return;
            }

            if (CursorLockService.Lock(operatorMonitor))
            {
                Log("鼠標已鎖定在作業螢幕：" + operatorMonitor);
            }
            else
            {
                Log("鼠標鎖定失敗（系統可能拒絕了裁剪範圍）。");
                LockCursorCheck.IsChecked = false;
            }
        }
        else
        {
            CursorLockService.Unlock();
            Log("鼠標已解除鎖定。");
        }
    }

    // ---------------------------------------------------------------- 視窗訊息

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);

        IntPtr hwnd = new WindowInteropHelper(this).Handle;

        // 插拔投影線、切換解析度時重定位，避免投影幕變空白或位置錯誤
        var source = HwndSource.FromHwnd(hwnd);
        source?.AddHook(WndProc);

        StartInputServices();
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == Native.WM_DISPLAYCHANGE)
        {
            Dispatcher.BeginInvoke(new Action(OnDisplayChanged));
        }

        return IntPtr.Zero;
    }

    private void OnDisplayChanged()
    {
        int before = _monitors.Count;
        LoadMonitors();
        Log($"偵測到顯示器設定變更，目前 {_monitors.Count} 台螢幕（原 {before} 台）。");

        if (_target is null)
        {
            return;
        }

        if (_ppt.IsShowRunning && _ppt.Reposition(_target))
        {
            Log("已依新的螢幕配置重新定位放映視窗。");
        }

        if (_live?.Kind == "material" && _materialProjectionActive)
        {
            ShowMaterialPage();
            Log("已依新的螢幕配置重新定位素材畫面。");
        }
    }

    private void StartInputServices()
    {
        try
        {
            IntPtr hwnd = new WindowInteropHelper(this).Handle;
            _rawInput = new RawInputService();
            _rawInput.KeyDown += OnRawKey;
            _rawInput.Start(hwnd);
            Log($"原始輸入已就緒（目標視窗 0x{hwnd.ToInt64():X}）。");
        }
        catch (Exception ex)
        {
            Log("原始輸入初始化失敗：" + ex.Message);
        }

        try
        {
            _guard = new KeyboardGuard(() => _ppt.ShowWindowHandle, OnGuardedPageKey);
            _guard.Install();
        }
        catch (Exception ex)
        {
            Log("鍵盤防護初始化失敗：" + ex.Message);
        }
    }

    // ---------------------------------------------------------------- 其他

    private void LogToggle_Click(object sender, RoutedEventArgs e)
    {
        LogPanel.Visibility = LogPanel.Visibility == Visibility.Visible
            ? Visibility.Collapsed
            : Visibility.Visible;
    }

    /// <summary>右欄分區標題：點一下展開／收合，減少一次看到太多設定。</summary>
    private void SectionHeader_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.Tag is not string name)
        {
            return;
        }

        StackPanel? body = name switch
        {
            "放映設定" => ShowSectionBody,
            "翻頁筆" => PenSectionBody,
            "控制者" => OwnerSectionBody,
            _ => null,
        };

        if (body is null)
        {
            return;
        }

        bool expand = body.Visibility != Visibility.Visible;
        body.Visibility = expand ? Visibility.Visible : Visibility.Collapsed;
        button.Content = $"{name}　{(expand ? "▾" : "▸")}";
    }

    private void Log(string message)
    {
        LogBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
        LogBox.ScrollToEnd();
        AppLog.Info(message, "UI");
    }

    protected override void OnClosed(EventArgs e)
    {
        CursorLockService.Unlock();
        _zoomService.Restore();
        _scriptureWindow?.Close();
        _webWindow?.Close();
        _guard?.Dispose();
        _rawInput?.Dispose();
        _materialWindow?.Close();
        _mobile?.Dispose();
        _settings.MobileDevices = _mobileDevices;
        SettingsStore.Save(_settings);
        base.OnClosed(e);
    }
}
