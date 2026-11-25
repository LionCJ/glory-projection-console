using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Windows;
using ProjectionManager.Services;

namespace ProjectionManager;

/// <summary>
/// V2.0 手機遙控器：桌面端開一個只綁區網的小型 HTTP 服務，
/// 手機用瀏覽器配對後即可上下頁、跳頁、黑屏、切換投影源。
/// 所有指令都回到 UI 執行緒，走控制台原本的按鈕邏輯，效果與翻頁筆相同。
/// </summary>
public partial class MainWindow
{
    private MobileControlServer? _mobile;
    private List<MobileDevice> _mobileDevices = new();
    private string _mobilePin = string.Empty;

    private void StartMobileControl()
    {
        _mobileDevices = _settings.MobileDevices ?? new List<MobileDevice>();
        _mobilePin = NewMobilePin();

        if (!_settings.MobileControlEnabled)
        {
            Log("手機控制：設定為關閉，未啟動服務。");
            return;
        }

        _mobile = new MobileControlServer(
            _settings.MobileControlPort,
            () => _mobilePin,
            () => _mobileDevices,
            () =>
            {
                _settings.MobileDevices = _mobileDevices;
                SettingsStore.Save(_settings);
            },
            BuildMobileState,
            ExecuteMobileCommand,
            _testMode ? IPAddress.Loopback : null,
            _testMode ? (id, role) => Dispatcher.Invoke(() => SetMobileDeviceRole(id, role)) : null,
            _testMode ? () => Dispatcher.Invoke(() => BuildTestLayout()) : null,
            _testMode ? () => Dispatcher.Invoke(() => BuildTestShot()) : null,
            _testMode ? action => Dispatcher.Invoke(() => BuildTestAction(action)) : null);

        if (_mobile.Start())
        {
            Log($"手機控制已啟動：{MobileUrl()}　配對碼 {_mobilePin}（手機需在同一個 Wi-Fi）");
            if (_testMode)
            {
                Title = $"Glory 榮耀 · 投影控制台（測試模式 PIN {_mobilePin}）";
            }
        }
        else
        {
            Log($"手機控制無法啟動（連接埠 {_settings.MobileControlPort}）：{_mobile.LastError}");
        }
    }

    private static string NewMobilePin()
    {
        // 6 位數字，避開開頭為 0 造成手機輸入被吃掉
        int value = Random.Shared.Next(100000, 1000000);
        return value.ToString();
    }

    private string MobileUrl()
    {
        return $"http://{LocalIPv4()}:{_settings.MobileControlPort}";
    }

    private static string LocalIPv4()
    {
        try
        {
            var candidates = new List<(string Address, int Score)>();
            foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (nic.OperationalStatus != OperationalStatus.Up
                    || nic.NetworkInterfaceType == NetworkInterfaceType.Loopback)
                {
                    continue;
                }

                foreach (var address in nic.GetIPProperties().UnicastAddresses)
                {
                    if (address.Address.AddressFamily != AddressFamily.InterNetwork)
                    {
                        continue;
                    }

                    string ip = address.Address.ToString();
                    if (ip.StartsWith("169.254.", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    // 常見家用／教會區網優先
                    int score = ip.StartsWith("192.168.", StringComparison.Ordinal) ? 3
                        : ip.StartsWith("10.", StringComparison.Ordinal) ? 2
                        : ip.StartsWith("172.", StringComparison.Ordinal) ? 1
                        : 0;
                    candidates.Add((ip, score));
                }
            }

            if (candidates.Count > 0)
            {
                return candidates.OrderByDescending(c => c.Score).First().Address;
            }
        }
        catch
        {
            // 取不到區網 IP 時退回本機位址
        }

        return "127.0.0.1";
    }

    /// <summary>測試模式專用：回報介面實際版面尺寸，供自動化驗證「預覽區不再佔滿畫面」。</summary>
    private object BuildTestLayout()
    {
        double center = CenterColumn.ActualHeight;
        double preview = PreviewArea.ActualHeight;
        double thumbs = ThumbArea.ActualHeight;
        double previewW = PreviewArea.ActualWidth;
        double cardW = Math.Max(0, (previewW - 10) / 2);
        double liveW = PreviewLiveImage.ActualWidth;
        double liveH = PreviewLiveImage.ActualHeight;
        double bodyW = 0;
        double bodyH = 0;
        if (PreviewLiveImage.Parent is FrameworkElement body)
        {
            bodyW = body.ActualWidth;
            bodyH = body.ActualHeight;
        }

        double imagePixelW = 0;
        double imagePixelH = 0;
        if (PreviewLiveImage.Source is System.Windows.Media.Imaging.BitmapSource bitmap)
        {
            imagePixelW = bitmap.PixelWidth;
            imagePixelH = bitmap.PixelHeight;
        }

        return new
        {
            ok = true,
            centerH = Math.Round(center, 1),
            previewH = Math.Round(preview, 1),
            thumbsH = Math.Round(thumbs, 1),
            previewRatio = center > 0 ? Math.Round(preview / center, 3) : 0,
            thumbsRatio = center > 0 ? Math.Round(thumbs / center, 3) : 0,
            previewW = Math.Round(previewW, 1),
            cardW = Math.Round(cardW, 1),
            previewBodyW = Math.Round(bodyW, 1),
            previewBodyH = Math.Round(bodyH, 1),
            previewBodyAspect = bodyH > 0 ? Math.Round(bodyW / bodyH, 3) : 0,
            mainAspect = Math.Round(_ppt.GetPresentationAspect(), 3),
            imagePixelW = Math.Round(imagePixelW, 0),
            imagePixelH = Math.Round(imagePixelH, 0),
            imagePixelAspect = imagePixelH > 0 ? Math.Round(imagePixelW / imagePixelH, 3) : 0,
            liveW = Math.Round(liveW, 1),
            liveH = Math.Round(liveH, 1),
            liveAspect = liveH > 0 ? Math.Round(liveW / liveH, 3) : 0,
            thumbCount = ThumbPanel.Children.Count,
            thumbTileW = Math.Round(_thumbTileW, 1),
            thumbTileH = Math.Round(_thumbTileH, 1),
            thumbCols = _thumbCols,

            // 九項改造的可自動驗證狀態
            targetDevice = _target?.DeviceName ?? string.Empty,
            targetWidth = _target?.Width ?? 0,
            targetHeight = _target?.Height ?? 0,
            showState = ShowStateText.Text,
            pushVisible = PushMaterialButton.Visibility == Visibility.Visible,
            pushText = PushMaterialButton.Content?.ToString() ?? string.Empty,
            debugBadgeVisible = DebugBadge.Visibility == Visibility.Visible,
            defaultGroupName = _groups.FirstOrDefault(g => g.IsDefault)?.Name ?? string.Empty,
            defaultGroupCount = _groups.Count(g => g.IsDefault),
            materialProjected = _materialProjectionActive,
            extraProjected = _extraProjectionActive,
            zoomProjected = _zoomProjectionActive,
            zoomRunning = _zoomService.IsRunning,
            zoomWindows = _zoomService.WindowCount,
            zoomPanelVisible = ZoomPanel.Visibility == Visibility.Visible,
            webPanelVisible = WebNavPanel.Visibility == Visibility.Visible,
            extraHintText = ExtraPushHint.Text,
            extraHintVisible = ExtraPushHint.Visibility == Visibility.Visible,
            mobileRunning = _mobile?.IsRunning == true,
            mobileError = _mobile?.LastError ?? string.Empty,
            previewLiveW = Math.Round(PreviewLiveImage.ActualWidth, 1),
            previewLiveH = Math.Round(PreviewLiveImage.ActualHeight, 1),
            navPrevW = Math.Round(PrevButton.ActualWidth, 1),
            navNextW = Math.Round(NextButton.ActualWidth, 1),
            navJumpW = Math.Round(JumpButton.ActualWidth, 1),
            navLastW = Math.Round(JumpLastButton.ActualWidth, 1),
            navPushW = Math.Round(PushMaterialButton.ActualWidth, 1),
            navBottomAligned = Math.Abs(PrevButton.TranslatePoint(new Point(0, 0), NavBar).Y
                    - NextButton.TranslatePoint(new Point(0, 0), NavBar).Y) < 2
                && Math.Abs(PrevButton.TranslatePoint(new Point(0, 0), NavBar).Y
                    - PushMaterialButton.TranslatePoint(new Point(0, 0), NavBar).Y) < 2
                && Math.Abs(JumpButton.TranslatePoint(new Point(0, 0), NavBar).Y
                    - JumpLastButton.TranslatePoint(new Point(0, 0), NavBar).Y) < 2,
            topButtonsSized = OpenButton.ActualHeight >= 32 && StartShowButton.ActualHeight >= 32
                && BlackButton.ActualHeight >= 32 && ResumeButton.ActualHeight >= 32
                && RepositionButton.ActualHeight >= 32 && EndShowButton.ActualHeight >= 32
                && SettingsButton.ActualHeight >= 32,
            topOpenW = Math.Round(OpenButton.ActualWidth, 1),
            topStartW = Math.Round(StartShowButton.ActualWidth, 1),
            topSettingsW = Math.Round(SettingsButton.ActualWidth, 1),
            refreshMonitorsW = Math.Round(RefreshMonitorsButton.ActualWidth, 1),
            refreshMonitorsStretched = RefreshMonitorsButton.ActualWidth >= 300
        };
    }

    /// <summary>測試模式專用：用同一條 UI 路徑驗證「先加入、再推送」與服務啟停。</summary>
    private object BuildTestAction(string action)
    {
        switch (action)
        {
            case "seed-scripture":
            {
                if (_scriptures.Count == 0)
                {
                    _scriptures.Add(new ScripturePassage
                    {
                        Book = "約翰福音",
                        Chapter = 3,
                        StartVerse = 16,
                        EndVerse = 16,
                        FontSize = 48,
                        Theme = 0,
                    });
                    RebuildSupplementalSources();
                    RenderSources();
                }

                var source = _scriptureSources.FirstOrDefault();
                if (source is null)
                {
                    return new { ok = false, error = "沒有可用經文" };
                }

                ActivateSource(source);
                return new { ok = true, kind = "scripture", projected = _extraProjectionActive, status = ShowStateText.Text };
            }

            case "seed-web":
            {
                if (_webs.Count == 0)
                {
                    _webs.Add(new WebLink { Name = "測試網頁", Url = "https://example.com/" });
                    RebuildSupplementalSources();
                    RenderSources();
                }

                var source = _webSources.FirstOrDefault();
                if (source is null)
                {
                    return new { ok = false, error = "沒有可用網頁" };
                }

                ActivateSource(source);
                return new { ok = true, kind = "web", projected = _extraProjectionActive, status = ShowStateText.Text };
            }

            case "select-zoom":
                _zoomService.Refresh();
                ActivateSource(_zoomSource);
                return new { ok = true, kind = "zoom", running = _zoomService.IsRunning, windows = _zoomService.WindowCount, projected = _zoomProjectionActive, status = ShowStateText.Text };

            case "push":
                if (_live is null)
                {
                    return new { ok = false, error = "沒有選中的投影源" };
                }

                PushMaterialButton_Click(this, new RoutedEventArgs());
                return new
                {
                    ok = true,
                    kind = _live.Kind,
                    projected = _live.Kind == "zoom" ? _zoomProjectionActive : _live.Kind == "material" ? _materialProjectionActive : _extraProjectionActive,
                    status = ShowStateText.Text,
                };

            case "mobile-stop":
                StopMobileService();
                return new { ok = true, running = _mobile?.IsRunning == true, error = _mobile?.LastError };

            case "mobile-start":
                StartMobileService();
                return new { ok = _mobile?.IsRunning == true, running = _mobile?.IsRunning == true, error = _mobile?.LastError };

            case "mobile-selftest":
                _ = TestMobileLoopbackAsync().ContinueWith(t =>
                {
                    if (t.Result.Ok)
                    {
                        Log("本機自測：" + t.Result.Message);
                    }
                });
                return new { ok = _mobile?.IsRunning == true, running = _mobile?.IsRunning == true };

            default:
                return new { ok = false, error = "未知測試動作：" + action };
        }
    }

    /// <summary>測試模式專用：用 WPF 自身渲染視窗，避免鎖屏／遠端桌面擷取失效。</summary>
    private object BuildTestShot()
    {
        int w = (int)Math.Max(1, RootGrid.ActualWidth);
        int h = (int)Math.Max(1, RootGrid.ActualHeight);
        var bmp = new System.Windows.Media.Imaging.RenderTargetBitmap(w, h, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
        bmp.Render(RootGrid);
        var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
        encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bmp));
        using var ms = new System.IO.MemoryStream();
        encoder.Save(ms);
        return new { ok = true, png = Convert.ToBase64String(ms.ToArray()) };
    }

    private MobileState BuildMobileState()
    {
        if (!Dispatcher.CheckAccess())
        {
            return Dispatcher.Invoke(BuildMobileStateCore);
        }

        return BuildMobileStateCore();
    }

    private MobileState BuildMobileStateCore()
    {
        var state = new MobileState
        {
            ShowRunning = _ppt.IsShowRunning,
            HasMain = (_mainPath is not null && File.Exists(_mainPath)) ||
                      (_ppt.PresentationPath is not null && File.Exists(_ppt.PresentationPath)),
            Black = _black,
            Mode = _mode == ControlMode.Operator ? "operator" : "speaker",
            SourceName = _live is null
                ? string.Empty
                : _live.Kind == "main" && _mainPath is null
                    ? "尚未加入主 PPT"
                    : _live.Name,
            SourceKind = _live?.Kind ?? string.Empty,
            Status = ShowStateText.Text,
            Projected = IsProjected(_live),
        };

        switch (_live?.Kind)
        {
            case "main":
                state.Page = _ppt.IsShowRunning ? _ppt.CurrentSlide : 0;
                state.Total = ShowTotal;
                break;
            case "material":
                state.Page = _live.PageIndex + 1;
                state.Total = _live.Material?.Pages.Count ?? 0;
                break;
            case "scripture":
                state.Page = _scriptureWindow is { IsVisible: true } ? _scriptureWindow.CurrentIndex + 1 : 0;
                state.Total = _scriptureWindow is { IsVisible: true } ? _scriptureWindow.PageCount : 0;
                break;
            default:
                state.Page = 0;
                state.Total = 0;
                break;
        }

        for (int i = 0; i < _sources.Count; i++)
        {
            var source = _sources[i];
            state.Sources.Add(BuildMobileSourceInfo(source, i, state));
        }

        // 手機端投影源清單要包含左欄全部來源：主 PPT → 素材 → 經文 → 網頁 → Zoom。
        // _sources 只存主 PPT 與素材，額外投影源接在後面，索引與切換指令共用同一份順序。
        int offset = _sources.Count;
        foreach (var source in _scriptureSources.Concat(_webSources).Append(_zoomSource))
        {
            state.Sources.Add(BuildMobileSourceInfo(source, offset++, state));
        }

        return state;
    }

    private MobileSourceInfo BuildMobileSourceInfo(SourceItem source, int index, MobileState state)
    {
        int page = 0;
        int total = 0;
        bool projected = false;

        string group = source.Kind switch
        {
            "main" => "主 PPT",
            "material" => _groups.FirstOrDefault(g => g.Id == source.Material?.GroupId)?.Name ?? "默認分組",
            "scripture" => "經文",
            "web" => "網頁",
            "zoom" => "Zoom",
            _ => "其他",
        };
        string groupKey = source.Kind switch
        {
            "main" => "main",
            "material" => "material:" + (source.Material?.GroupId ?? MaterialGroup.DefaultId),
            _ => source.Kind,
        };

        switch (source.Kind)
        {
            case "main":
                total = ShowTotal;
                page = _ppt.IsShowRunning ? _ppt.CurrentSlide : 0;
                projected = _ppt.IsShowRunning && _live == source;
                break;
            case "material":
                total = source.Material?.Pages.Count ?? 0;
                page = total > 0 ? source.PageIndex + 1 : 0;
                projected = _materialProjectionActive && _live == source;
                break;
            case "scripture":
                total = _scriptureWindow is { IsVisible: true } ? _scriptureWindow.PageCount : 0;
                page = total > 0 ? source.PageIndex + 1 : 0;
                projected = _extraProjectionActive && _live == source;
                break;
            case "web":
                projected = _extraProjectionActive && _live == source;
                break;
            case "zoom":
                projected = _zoomProjectionActive && _live == source;
                break;
        }

        return new MobileSourceInfo
        {
            Index = index,
            Name = source.Name,
            Kind = source.Kind,
            Group = group,
            GroupKey = groupKey,
            Live = source == _live && !(source.Kind == "main" && _mainPath is null),
            Projected = projected,
            Page = page,
            Total = total,
        };
    }

    private List<SourceItem> MobileSourceList()
    {
        var list = new List<SourceItem>(_sources);
        list.AddRange(_scriptureSources);
        list.AddRange(_webSources);
        list.Add(_zoomSource);
        return list;
    }

    private MobileCommandResult ExecuteMobileCommand(MobileDevice device, string cmd, string? arg)
    {
        if (!Dispatcher.CheckAccess())
        {
            return Dispatcher.Invoke(() => ExecuteMobileCommand(device, cmd, arg));
        }

        switch (cmd)
        {
            case "next":
                return MobileNavigate(true, device);

            case "prev":
                return MobileNavigate(false, device);

            case "goto":
                return MobileGoto(arg);

            case "black":
                if (_live is null)
                {
                    return MobileCommandResult.Fail("目前沒有投影源");
                }
                BlackButton_Click(this, new RoutedEventArgs());
                Log($"手機控制（{device.Name}）：黑屏");
                return MobileCommandResult.Success("已黑屏");

            case "resume":
                ResumeButton_Click(this, new RoutedEventArgs());
                Log($"手機控制（{device.Name}）：恢復畫面");
                return MobileCommandResult.Success("已恢復畫面");

            case "stop":
                return MobileStopShow(device);

            case "start":
                return MobileStartShow(device);

            case "push":
                if (_live is null)
                {
                    return MobileCommandResult.Fail("尚未選擇投影源");
                }

                if (_live.Kind == "main")
                {
                    return MobileCommandResult.Fail("主 PPT 請用「開始主 PPT 放映」與「結束放映」控制");
                }

                bool wasProjected = IsProjected(_live);
                var sourceName = _live.Name;
                PushMaterialButton_Click(this, new RoutedEventArgs());
                Log($"手機控制（{device.Name}）：{(wasProjected ? "斷開投影" : "推送投影")} → {sourceName}");
                return MobileCommandResult.Success(wasProjected ? "已斷開投影" : "已推送投影");

            case "source":
                return MobileActivateSource(arg, device);

            default:
                return MobileCommandResult.Fail("不支援的指令");
        }
    }

    private MobileCommandResult MobileNavigate(bool forward, MobileDevice device)
    {
        string action = forward ? "下一頁" : "上一頁";

        if (_live is null)
        {
            return MobileCommandResult.Fail("尚未選擇投影源");
        }

        if (_live.Kind == "zoom")
        {
            return MobileCommandResult.Fail("Zoom 投影源請在 Zoom 視窗內翻頁");
        }

        if (_live.Kind == "main" && !_ppt.IsShowRunning)
        {
            return MobileCommandResult.Fail("主 PPT 尚未放映");
        }

        Navigate(forward);
        Log($"手機控制（{device.Name}）：{action}");
        return MobileCommandResult.Success("已" + action);
    }

    private MobileCommandResult MobileGoto(string? arg)
    {
        if (!int.TryParse(arg, out int page) || page <= 0)
        {
            return MobileCommandResult.Fail("頁碼格式不正確");
        }

        if (_live?.Kind is "main" or "material")
        {
            if (TryJumpToPage(page, out string message))
            {
                return MobileCommandResult.Success($"已跳到第 {page} 頁");
            }

            return MobileCommandResult.Fail(message);
        }

        if (_live?.Kind == "scripture")
        {
            return MobileCommandResult.Fail("經文請用上一頁／下一頁翻頁");
        }

        return MobileCommandResult.Fail("目前的投影源不支援跳頁");
    }

    private MobileCommandResult MobileStopShow(MobileDevice device)
    {
        if (!_ppt.IsShowRunning)
        {
            return MobileCommandResult.Fail("目前沒有進行中的放映");
        }

        try
        {
            Native.SetNoActivate(new System.Windows.Interop.WindowInteropHelper(this).Handle, false);
            _ppt.EndShow();
        }
        catch (Exception ex)
        {
            return MobileCommandResult.Fail("結束放映失敗：" + ex.Message);
        }

        ApplyConsoleActivation();
        _black = false;
        _mode = ControlMode.Operator;
        _handoverSlide = 0;
        ApplyMode();
        Log($"手機控制（{device.Name}）：結束放映，控制權回到操作員。");
        RenderAll();
        return MobileCommandResult.Success("已結束放映");
    }

    private MobileCommandResult MobileStartShow(MobileDevice device)
    {
        if (_ppt.IsShowRunning)
        {
            return MobileCommandResult.Fail("主 PPT 已在放映");
        }

        if ((_mainPath is null || !File.Exists(_mainPath)) &&
            (_ppt.PresentationPath is null || !File.Exists(_ppt.PresentationPath)))
        {
            return MobileCommandResult.Fail("尚未加入主 PPT，請先在電腦端加入");
        }

        if (_target is null)
        {
            return MobileCommandResult.Fail("尚未選擇投影螢幕");
        }

        // 不在 HTTP 執行緒等待 PowerPoint；開始放映可能會有螢幕模式提醒或檔案解鎖延遲。
        Dispatcher.BeginInvoke(new Action(async () =>
        {
            try
            {
                await StartShow();
            }
            catch (Exception ex)
            {
                Log("手機控制開始放映失敗：" + ex.Message);
            }
        }));

        Log($"手機控制（{device.Name}）：開始主 PPT 放映");
        return MobileCommandResult.Success("正在開始主 PPT 放映，請稍候");
    }

    private MobileCommandResult MobileActivateSource(string? arg, MobileDevice device)
    {
        if (string.IsNullOrWhiteSpace(arg))
        {
            return MobileCommandResult.Fail("找不到這個投影源");
        }

        // 新格式：index|page；舊手機只送 index 時仍相容。
        string[] parts = arg.Split('|');
        if (!int.TryParse(parts[0], out int index))
        {
            return MobileCommandResult.Fail("找不到這個投影源");
        }

        int page = 0;
        if (parts.Length > 1 && int.TryParse(parts[1], out int parsedPage) && parsedPage > 0)
        {
            page = parsedPage;
        }

        var sources = MobileSourceList();
        if (index < 0 || index >= sources.Count)
        {
            return MobileCommandResult.Fail("找不到這個投影源");
        }

        var source = sources[index];
        ActivateSource(source);

        string suffix = string.Empty;
        if (page > 0)
        {
            switch (source.Kind)
            {
                case "main":
                    if (!_ppt.IsShowRunning)
                    {
                        suffix = "；主 PPT 尚未放映，頁碼會在開始放映後生效";
                    }
                    else if (TryJumpToPage(page, out string message))
                    {
                        suffix = $"　第 {page} 頁";
                    }
                    else
                    {
                        return MobileCommandResult.Fail("已切換來源，但跳頁失敗：" + message);
                    }
                    break;

                case "material":
                {
                    int last = Math.Max(0, (source.Material?.Pages.Count ?? 1) - 1);
                    source.PageIndex = Math.Min(Math.Max(0, page - 1), last);
                    _live = source;
                    RenderAll();
                    suffix = $"　第 {source.PageIndex + 1} / {last + 1} 頁（尚未推送）";
                    break;
                }

                case "scripture":
                {
                    source.PageIndex = Math.Max(0, page - 1);
                    _live = source;
                    RenderAll();
                    suffix = $"　第 {source.PageIndex + 1} 頁（尚未推送）";
                    break;
                }

                default:
                    suffix = "；此來源不需要頁碼";
                    break;
            }
        }

        Log($"手機控制（{device.Name}）：切換投影源 → {source.Name}{suffix}");
        return MobileCommandResult.Success($"已切換到 {source.Name}{suffix}");
    }

    // ---------------------------------------------------------------- 對「手機控制」視窗的介面

    private MobileControlDialog? _mobileDialog;

    private void MobileButton_Click(object sender, RoutedEventArgs e)
    {
        if (_mobileDialog is { IsVisible: true })
        {
            _mobileDialog.Activate();
            return;
        }

        _mobileDialog = new MobileControlDialog(this) { Owner = this };
        _mobileDialog.Closed += (_, _) => _mobileDialog = null;
        _mobileDialog.Show();
    }

    public bool IsMobileRunning => _mobile?.IsRunning == true;

    public int MobileDeviceCount => _mobileDevices.Count;

    public sealed record MobileDeviceView(string Id, string Name, string Role, string LastSeen);

    public sealed record MobileInfo(string Url, string Pin, bool Running, int Port, string? Error, bool Enabled, List<MobileDeviceView> Devices);

    public MobileInfo GetMobileInfo()
    {
        var devices = _mobileDevices
            .Select(d => new MobileDeviceView(d.Id, d.Name, d.Role, d.LastSeen))
            .ToList();

        return new MobileInfo(
            MobileUrl(),
            _mobilePin,
            _mobile?.IsRunning == true,
            _settings.MobileControlPort,
            _mobile?.LastError,
            _settings.MobileControlEnabled,
            devices);
    }

    public void SetMobileControlEnabled(bool enabled)
    {
        _settings.MobileControlEnabled = enabled;
        SettingsStore.Save(_settings);

        if (enabled)
        {
            if (_mobile is null)
            {
                StartMobileControl();
            }
            else if (!_mobile.IsRunning)
            {
                _mobile.Start();
            }

            Log($"手機控制已開啟：{MobileUrl()}　配對碼 {_mobilePin}");
        }
        else
        {
            _mobile?.Stop();
            Log("手機控制已關閉（已配對的裝置保留，下次開啟仍可自動連線）。");
        }
    }

    /// <summary>由手機控制視窗直接啟動服務；不論原本勾選狀態，啟動後都保持開啟。</summary>
    public void StartMobileService()
    {
        SetMobileControlEnabled(true);
    }

    /// <summary>由手機控制視窗直接停止服務，並同步關閉設定，避免下次啟動又自動開。</summary>
    public void StopMobileService()
    {
        SetMobileControlEnabled(false);
    }

    /// <summary>本機自測：從 127.0.0.1 抓首頁，確認 HTTP 服務真的能回應（排除防火牆與手機端因素）。</summary>
    public async Task<(bool Ok, string Message)> TestMobileLoopbackAsync()
    {
        if (_mobile?.IsRunning != true)
        {
            return (false, "服務未運行，請先按「啟動服務」。");
        }

        string url = $"http://127.0.0.1:{_settings.MobileControlPort}/";
        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
            using var response = await client.GetAsync(url);
            string body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                return (false, $"HTTP {response.StatusCode}，服務有回應但狀態異常。");
            }

            return body.Contains("投影控制", StringComparison.Ordinal) || body.Contains("配對", StringComparison.Ordinal)
                ? (true, $"本機自測成功：{url} 回傳 {body.Length} bytes。若手機仍打不開，請檢查同一 Wi-Fi 與防火牆。")
                : (false, $"本機有回應（{body.Length} bytes），但內容不像手機控制頁。請重新啟動服務。");
        }
        catch (Exception ex)
        {
            return (false, $"本機自測失敗：{ex.Message}。請確認連接埠 {_settings.MobileControlPort} 沒有被其他程式佔用。");
        }
    }

    /// <summary>列出所有可用的 IPv4 與介面，讓使用者知道手機該連哪一個網段。</summary>
    public string NetworkSummary()
    {
        try
        {
            var lines = new List<string>();
            foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (nic.OperationalStatus != OperationalStatus.Up
                    || nic.NetworkInterfaceType == NetworkInterfaceType.Loopback)
                {
                    continue;
                }

                foreach (var address in nic.GetIPProperties().UnicastAddresses)
                {
                    if (address.Address.AddressFamily != AddressFamily.InterNetwork)
                    {
                        continue;
                    }

                    string ip = address.Address.ToString();
                    if (ip.StartsWith("169.254.", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    lines.Add($"{nic.Name}：{ip}");
                }
            }

            return lines.Count == 0
                ? "找不到區網 IP；手機與電腦可能不在同一個 Wi-Fi。"
                : string.Join("　·　", lines);
        }
        catch (Exception ex)
        {
            return "讀取網路介面失敗：" + ex.Message;
        }
    }

    public void SetMobileDeviceRole(string id, string role)
    {
        var device = _mobileDevices.FirstOrDefault(d => d.Id == id);
        if (device is null)
        {
            return;
        }

        device.Role = role == "speaker" ? "speaker" : "operator";
        _settings.MobileDevices = _mobileDevices;
        SettingsStore.Save(_settings);
        Log($"手機裝置「{device.Name}」角色改為：{(device.Role == "speaker" ? "講員（只上下頁）" : "操作員（全部控制）")}");
    }

    public void RevokeMobileDevice(string id)
    {
        var device = _mobileDevices.FirstOrDefault(d => d.Id == id);
        if (device is null)
        {
            return;
        }

        _mobileDevices.Remove(device);
        _settings.MobileDevices = _mobileDevices;
        SettingsStore.Save(_settings);
        Log($"已撤銷手機裝置「{device.Name}」，該手機需要重新配對。");
    }

    public string RegenerateMobilePin()
    {
        _mobilePin = NewMobilePin();
        Log($"手機配對碼已更新：{_mobilePin}");
        return _mobilePin;
    }
}
