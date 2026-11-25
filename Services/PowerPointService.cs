using System.IO;

namespace ProjectionManager.Services;

public sealed class PowerPointException : Exception
{
    public PowerPointException(string message) : base(message)
    {
    }
}

/// <summary>
/// PowerPoint 放映控制。採晚綁定（dynamic + ProgID），不依賴特定 Office 版本的 Interop 組件。
/// 全部呼叫都在建立 PowerPoint 的同一條 STA 執行緒（WPF UI 執行緒）上完成。
/// </summary>
public sealed class PowerPointService
{
    // ppSlideShowState
    public const int StateRunning = 1;
    public const int StateBlackScreen = 3;
    public const int StateWhiteScreen = 4;

    private dynamic _app;
    private dynamic _presentation;
    private dynamic _show;
    private IntPtr _showHwnd = IntPtr.Zero;
    private IntPtr _editorHwnd = IntPtr.Zero;
    private bool _running;
    private bool _windowed;

    // 開啟主 PPT 當下的儲存狀態。匯出預覽、啟動放映都可能讓 PowerPoint 標成「未儲存」；
    // 關閉時只在「開啟時本來就是已儲存」的情況下還原，避免關閉簡報時跳出「是否儲存」。
    private bool _savedOnOpen = true;

    private bool ApplicationAlive(dynamic app)
    {
        try
        {
            _ = app.Version;
            return true;
        }
        catch
        {
            return false;
        }
    }

    public bool IsApplicationOpen => _app is not null && ApplicationAlive(_app);
    public bool IsShowRunning => _running;
    public IntPtr ShowWindowHandle => _showHwnd;
    public string? PresentationPath { get; private set; }

    /// <summary>
    /// 添加素材期間若要暫時結束放映，設為 true，匯出時就不會把 PowerPoint 視窗還原到前景，
    /// 而是保持最小化，避免編輯視窗閃現。
    /// </summary>
    public bool KeepApplicationMinimized { get; set; }

    /// <summary>
    /// 在結束放映前先記住 PowerPoint 編輯主視窗（放映視窗是 screenClass，編輯視窗是 PPTFrameClass）。
    /// 放映一結束就立刻用 Win32 隱藏它，避免編輯視窗閃到前景。
    /// </summary>
    public void CaptureEditorWindow()
    {
        if (_editorHwnd != IntPtr.Zero && Native.IsWindow(_editorHwnd))
        {
            return;
        }

        _editorHwnd = IntPtr.Zero;
        uint pid = 0;

        if (_showHwnd != IntPtr.Zero)
        {
            Native.GetWindowThreadProcessId(_showHwnd, out pid);
        }

        if (pid != 0)
        {
            _editorHwnd = Native.FindPowerPointFrameWindow((int)pid);
        }

        if (_editorHwnd == IntPtr.Zero)
        {
            // 找不到放映視窗時，退而找目前唯一的 PowerPoint 進程。
            foreach (var process in System.Diagnostics.Process.GetProcessesByName("POWERPNT"))
            {
                _editorHwnd = Native.FindPowerPointFrameWindow(process.Id);
                if (_editorHwnd != IntPtr.Zero)
                {
                    break;
                }
            }
        }
    }

    public void HideEditorWindow()
    {
        if (_editorHwnd == IntPtr.Zero || !Native.IsWindow(_editorHwnd))
        {
            CaptureEditorWindow();
        }

        if (_editorHwnd != IntPtr.Zero)
        {
            Native.ShowWindow(_editorHwnd, Native.SW_HIDE);
        }
    }

    public void ShowEditorWindowMinimized()
    {
        if (_editorHwnd != IntPtr.Zero && Native.IsWindow(_editorHwnd))
        {
            Native.ShowWindow(_editorHwnd, Native.SW_SHOWMINNOACTIVE);
        }
    }

    public void MinimizeApplicationWindow()
    {
        if (_app is null)
        {
            return;
        }

        try
        {
            _app.WindowState = 2;   // ppWindowMinimized
        }
        catch
        {
            // 取不到視窗也不影響匯入
        }
    }

    public void RestoreApplicationWindow()
    {
        if (_app is null)
        {
            return;
        }

        try
        {
            _app.WindowState = 1;   // ppWindowNormal
        }
        catch
        {
            // 還原失敗不影響放映
        }
    }

    /// <summary>PowerPoint 被關掉後，所有 COM 參考都失效；把它們全部丟掉，下次用新實例。</summary>
    public void ResetComObjects()
    {
        _presentation = null;
        _show = null;
        _running = false;
        _showHwnd = IntPtr.Zero;
        if (_app is not null && !ApplicationAlive(_app))
        {
            _app = null;
        }
    }

    /// <summary>簡報物件還活著嗎。使用者若把 PowerPoint 視窗關掉，_presentation 會變成失效的 COM 參考。</summary>
    public bool HasLivePresentation
    {
        get
        {
            if (_presentation is null)
            {
                return false;
            }

            try
            {
                _ = _presentation.Slides.Count;
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    /// <summary>
    /// 簡報是否已經可以放映（SlideShowSettings 拿得到）。
    /// 剛 Open 完的幾百毫秒內 PowerPoint 可能還沒建好文件視窗，這裡用 try/catch 當探測。
    /// </summary>
    public bool IsShowReady
    {
        get
        {
            if (!HasLivePresentation)
            {
                return false;
            }

            try
            {
                _ = _presentation.SlideShowSettings;
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    /// <summary>丟掉失效的簡報參考，下次 Open 會真的重新開檔。</summary>
    public void ForgetPresentation()
    {
        _presentation = null;
    }

    public void EnsureApplication()
    {
        if (_app is not null)
        {
            if (ApplicationAlive(_app))
            {
                return;
            }

            // 舊實例已被關閉（RPC 伺服器不可用），重新建立
            _app = null;
            _presentation = null;
            _show = null;
            _running = false;
            _showHwnd = IntPtr.Zero;
        }

        var type = Type.GetTypeFromProgID("PowerPoint.Application");
        if (type is null)
        {
            throw new PowerPointException("未偵測到 PowerPoint。請先安裝 Microsoft PowerPoint 桌面版（2019、2021 或 365）。");
        }

        _app = Activator.CreateInstance(type);
        if (_app is null)
        {
            throw new PowerPointException("無法啟動 PowerPoint 物件。");
        }

        // PowerPoint 必須可見才能正常放映
        _app.Visible = -1;
    }

    /// <summary>
    /// PowerPoint 裡已經手動開著同一份檔案時，直接接管它，避免重開失敗或被鎖檔。
    /// </summary>
    private bool TryAttachExisting(string path)
    {
        try
        {
            int count = Convert.ToInt32(_app.Presentations.Count);
            for (int i = 1; i <= count; i++)
            {
                dynamic pres = _app.Presentations.Item(i);
                string full = Convert.ToString(pres.FullName) ?? string.Empty;
                if (string.Equals(full, path, StringComparison.OrdinalIgnoreCase))
                {
                    _presentation = pres;
                    PresentationPath = path;
                    try
                    {
                        _savedOnOpen = Convert.ToBoolean(pres.Saved);
                    }
                    catch
                    {
                        _savedOnOpen = false;
                    }

                    return true;
                }
            }
        }
        catch
        {
            // 取不到就照正常流程開檔
        }

        return false;
    }

    public void Open(string path)
    {
        EnsureApplication();

        // 若這份簡報已經被使用者在 PowerPoint 裡手動開著，直接接管那個視窗，不要重開
        if (TryAttachExisting(path))
        {
            return;
        }

        // 已經開著同一份就不要重開：重開會在 PowerPoint 裡多一個視窗，
        // 使用者看到的是「按了加入主 PPT 但好像沒反應」。
        // 但若簡報已被關掉（COM 參考失效），必須真的重開。
        if (_presentation is not null && PresentationPath is not null &&
            string.Equals(PresentationPath, path, StringComparison.OrdinalIgnoreCase) &&
            HasLivePresentation)
        {
            return;
        }

        Exception? lastError = null;
        for (int attempt = 0; attempt < 2; attempt++)
        {
            try
            {
                _presentation = _app.Presentations.Open(path, 0, 0, -1);
                PresentationPath = path;
                try
                {
                    _savedOnOpen = Convert.ToBoolean(_presentation.Saved);
                }
                catch
                {
                    _savedOnOpen = true;
                }

                return;
            }
            catch (Exception ex)
            {
                lastError = ex;

                // 若簡報在 PowerPoint 中被關掉、或 PowerPoint 本身被關掉，
                // 這裡會拿到「RPC 伺服器不可用」之類的失效參考，重建實例再試一次。
                EnsureApplication();
            }
        }

        throw new PowerPointException(
            "開啟簡報失敗：" + lastError!.Message + Environment.NewLine +
            "若該檔已在 PowerPoint 中開啟，請先關閉該視窗再試；若檔案受保護或有密碼，請先解除。");
    }

    /// <summary>
    /// 只匯出一頁，供控制台做「播出中」的靜態預覽。
    /// 實測放映進行中呼叫 Slide.Export 可行（約 0.2–1.3 秒），不會中斷放映、不影響動畫。
    /// width/height 是上限，實際輸出會依簡報原始比例縮放，避免 4:3 被拉成 16:9。
    /// </summary>
    public bool ExportSlideTo(int index, string file, int width, int height)
    {
        if (!HasLivePresentation)
        {
            return false;
        }

        try
        {
            int total = Convert.ToInt32(_presentation.Slides.Count);
            if (index < 1 || index > total)
            {
                return false;
            }

            if (File.Exists(file))
            {
                File.Delete(file);
            }

            bool wasSaved = true;
            try
            {
                wasSaved = Convert.ToBoolean(_presentation.Saved);
            }
            catch
            {
                // 取不到就照常匯出
            }

            var (exportWidth, exportHeight) = FitExportSize(SlideAspect((object)_presentation), width, height);
            try
            {
                _presentation.Slides.Item(index).Export(file, "PNG", exportWidth, exportHeight);
            }
            finally
            {
                // Slide.Export 不會把 Saved 設回 True；若匯出讓簡報變成未儲存，
                // 在原本已儲存的情況下還原，避免關閉時出現「是否儲存」。
                if (wasSaved)
                {
                    try
                    {
                        _presentation.Saved = true;
                    }
                    catch
                    {
                        // 還原失敗不影響預覽
                    }
                }
            }

            return File.Exists(file);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>目前放映簡報的原始寬高比；沒有簡報或讀不到時回傳 0。</summary>
    public double GetPresentationAspect()
    {
        if (!HasLivePresentation)
        {
            return 0;
        }

        return SlideAspect((object)_presentation);
    }

    private static double SlideAspect(object presentation)
    {
        try
        {
            dynamic pres = presentation;
            double width = Convert.ToDouble(pres.PageSetup.SlideWidth);
            double height = Convert.ToDouble(pres.PageSetup.SlideHeight);
            return height > 0 ? width / height : 0;
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>把目標尺寸放進 width×height 的上限內，並保持來源比例。</summary>
    private static (int Width, int Height) FitExportSize(double aspect, int maxWidth, int maxHeight)
    {
        int width = Math.Max(1, maxWidth);
        int height = Math.Max(1, maxHeight);
        if (aspect <= 0)
        {
            return (width, height);
        }

        if (width / (double)height > aspect)
        {
            width = Math.Max(1, (int)Math.Round(height * aspect));
        }
        else
        {
            height = Math.Max(1, (int)Math.Round(width / aspect));
        }

        return (width, height);
    }

    /// <summary>
    /// 把一份簡報逐頁匯出成 PNG，供素材放映使用。
    /// 逐頁呼叫 Slide.Export 而不是整份 Export：檔名固定為 p001、p002，
    /// 不必處理中文語系下的「投影片1.PNG」命名，也不必猜排序。
    /// 匯出期間不能有放映在進行，否則 PowerPoint 會拒絕或匯出空白頁。
    /// </summary>
    public int ExportSlides(
        string path,
        string folder,
        int width,
        int height,
        Action<int, int>? progress = null,
        Func<bool>? cancelled = null)
    {
        EnsureApplication();

        if (_running)
        {
            throw new PowerPointException("放映進行中無法匯出素材，請先結束放映。");
        }

        Directory.CreateDirectory(folder);

        // 匯出素材只需要 PowerPoint 的 COM 引擎，不需要讓使用者看到 PowerPoint。
        // 先嘗試隱藏應用程式；若因為已有簡報開啟而不能隱藏，至少把視窗最小化，
        // 匯出完再還原，避免「加入素材」時突然跳出 PPT。
        // 添加素材期間（KeepApplicationMinimized）保持最小化，不要在匯出結束時還原到前景。
        bool keepMinimized = KeepApplicationMinimized;
        bool wasVisible = true;
        int windowStateBefore = 1;
        if (keepMinimized)
        {
            try { _app.WindowState = 2; } catch { }
        }
        else
        {
            try { wasVisible = Convert.ToBoolean(_app.Visible); } catch { }
            try { windowStateBefore = Convert.ToInt32(_app.WindowState); } catch { }
            try
            {
                _app.Visible = 0;
            }
            catch
            {
                try { _app.WindowState = 2; } catch { }   // ppWindowMinimized
            }
        }

        dynamic? pres = null;
        try
        {
            pres = _app.Presentations.Open(path, 0, 0, 0);   // WithWindow = false，匯出不需要視窗
            if (keepMinimized)
            {
                // 部分 PowerPoint 版本在 Open 時會把編輯視窗叫出來，這裡再藏一次。
                HideEditorWindow();
            }
        }
        catch (Exception ex)
        {
            if (!keepMinimized)
            {
                try
                {
                    if (wasVisible) { _app.Visible = -1; }
                    _app.WindowState = windowStateBefore;
                }
                catch { }
            }
            throw new PowerPointException("開啟素材簡報失敗：" + ex.Message);
        }

        bool wasSaved = true;
        try
        {
            wasSaved = Convert.ToBoolean(pres.Saved);
        }
        catch
        {
            // 取不到就照常匯出
        }

        try
        {
            int count = Convert.ToInt32(pres.Slides.Count);
            var (exportWidth, exportHeight) = FitExportSize(SlideAspect((object)pres), width, height);
            for (int i = 1; i <= count; i++)
            {
                if (cancelled?.Invoke() == true)
                {
                    throw new OperationCanceledException("使用者取消匯入。");
                }

                progress?.Invoke(i, count);
                string file = Path.Combine(folder, $"p{i:D3}.png");
                pres.Slides.Item(i).Export(file, "PNG", exportWidth, exportHeight);
            }

            return count;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new PowerPointException("匯出素材失敗：" + ex.Message);
        }
        finally
        {
            try
            {
                // Slide.Export 會把簡報標成未儲存；關閉前還原，避免留下「是否儲存」提示。
                if (wasSaved)
                {
                    pres.Saved = true;
                }
            }
            catch
            {
                // 還原失敗時仍要關閉
            }

            try
            {
                pres.Close();
            }
            catch
            {
                // 簡報已被關閉就不用處理
            }

            // 還原 PowerPoint 顯示狀態：匯出過程不讓它跳出來，結束後恢復原本狀態。
            // 添加素材期間保持最小化，避免編輯視窗在匯入完成後閃到前景。
            if (!keepMinimized)
            {
                try
                {
                    if (wasVisible) { _app.Visible = -1; }
                    _app.WindowState = windowStateBefore;
                }
                catch
                {
                    // 還原失敗不影響素材匯出結果
                }
            }
        }
    }

    /// <summary>
    /// 啟動放映並把放映視窗定位到投影螢幕。
    /// windowed=true 用 ppShowTypeWsize（視窗中瀏覽）：全螢幕放映一失去焦點就會暫停動畫，
    /// 這個模式是 PowerPoint 官方承認「失焦仍繼續播放動畫」的唯一模式；再由程式中去標題列鋪滿螢幕。
    /// </summary>
    public void StartShow(MonitorInfo target, int startSlide = 1, bool windowed = false)
    {
        if (_presentation is null || !HasLivePresentation)
        {
            ForgetPresentation();
            throw new PowerPointException("簡報已不在 PowerPoint 中開啟（視窗可能被關掉了）。請重新按「開始放映」，程式會自動重新開啟。");
        }

        if (!IsShowReady)
        {
            throw new PowerPointException("PowerPoint 還沒準備好放映（簡報剛開啟時需要幾百毫秒）。請再按一次「開始放映」。");
        }

        _windowed = windowed;

        try
        {
            var settings = _presentation.SlideShowSettings;

            // ShowType／ShowWithNarration／ShowWithAnimation 是「存在簡報檔裡」的放映設定。
            // 直接改會讓 PowerPoint 把簡報標成未儲存，關閉時就問「是否儲存」。
            // 這裡先記住原值，Run() 之後還原；若原本是已儲存狀態，也把 Saved 還原。
            bool captured = false;
            bool wasSaved = true;
            int oldShowType = 1;
            int oldNarration = -1;
            int oldAnimation = -1;
            try
            {
                oldShowType = Convert.ToInt32(settings.ShowType);
                oldNarration = Convert.ToInt32(settings.ShowWithNarration);
                oldAnimation = Convert.ToInt32(settings.ShowWithAnimation);
                wasSaved = Convert.ToBoolean(_presentation.Saved);
                captured = true;
            }
            catch
            {
                // 取不到原值時仍可放映，只是無法還原設定
            }

            try
            {
                settings.ShowType = windowed ? 2 : 1;                   // 2 = ppShowTypeWsize（視窗中瀏覽）, 1 = 全螢幕
                settings.ShowWithNarration = 0;                         // msoFalse
                settings.ShowWithAnimation = -1;                        // msoTrue

                // 絕對不要在 Run() 之前設 StartingSlide：實測在這個 PowerPoint 版本上，
                // 賦值後放映會立刻自我結束（視窗存活不到 0.5 秒）。要指定起始頁請在 Run() 之後 GotoSlide。
                // Run() 直接回傳放映視窗物件，也不需要輪詢 SlideShowWindows，更不能在 UI 執行緒 Sleep。
                _show = settings.Run();
            }
            finally
            {
                if (captured)
                {
                    try
                    {
                        var restore = _presentation.SlideShowSettings;
                        restore.ShowType = oldShowType;
                        restore.ShowWithNarration = oldNarration;
                        restore.ShowWithAnimation = oldAnimation;
                    }
                    catch
                    {
                        // 還原放映設定失敗不影響放映
                    }

                    try
                    {
                        // 只在「放映前本來就沒有未儲存變更」時還原，避免蓋掉使用者真正的編輯。
                        if (wasSaved)
                        {
                            _presentation.Saved = true;
                        }
                    }
                    catch
                    {
                        // 還原失敗不影響放映，僅代表關閉時可能仍會提示儲存
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new PowerPointException("啟動放映失敗：" + ex.Message);
        }

        _running = true;

        // SlideShowWindow.HWND 在部分版本取不到值，改用視窗類別 screenClass 定位
        _showHwnd = Native.FindSlideShowWindow();
        if (_showHwnd == IntPtr.Zero)
        {
            // 找不到視窗時要把已經起來的放映關掉，否則會留下一個沒人控制的放映窗
            try
            {
                _show!.View.Exit();
            }
            catch
            {
                // 已經結束就不用處理
            }

            _running = false;
            _show = null;
            throw new PowerPointException(
                "放映已啟動，但找不到放映視窗。請在 PowerPoint 的「投影片放映」設定中" +
                "取消勾選「使用簡報者檢視畫面」後重試。");
        }

        Reposition(target);

        // 起始頁只能在放映啟動後跳，不能在 Run() 之前設
        if (startSlide > 1)
        {
            try
            {
                _show!.View.GotoSlide(startSlide, 0);
            }
            catch
            {
                // 跳頁失敗不影響放映本身
            }
        }
    }

    /// <summary>把放映視窗移到指定螢幕並鋪滿。不啟用視窗，避免搶走焦點。</summary>
    public bool Reposition(MonitorInfo target)
    {
        if (!_running || _showHwnd == IntPtr.Zero)
        {
            return false;
        }

        if (!Native.IsWindow(_showHwnd))
        {
            _running = false;
            _showHwnd = IntPtr.Zero;
            return false;
        }

        // 視窗化放映會帶著標題列與可拖曳邊框；先拆掉再定位，才像全螢幕也不會被誤拖動
        if (_windowed)
        {
            Native.StripWindowFrame(_showHwnd);
        }

        Native.SetWindowPos(
            _showHwnd,
            IntPtr.Zero,                                   // HWND_TOP
            target.Left, target.Top, target.Width, target.Height,
            Native.SWP_NOACTIVATE | Native.SWP_SHOWWINDOW);

        return true;
    }

    public void Next() => Invoke(() => _show!.View.Next());

    public void Previous() => Invoke(() => _show!.View.Previous());

    public void GotoSlide(int index) => Invoke(() => _show!.View.GotoSlide(index, 0));

    /// <summary>
    /// 跳頁前先確認頁數在放映的有效範圍內，並把 COM 錯誤收成訊息而不是例外。
    /// 頁數對不上時（例如放映中的簡報已被換掉），只回報、不中斷放映、不讓 UI 掛掉。
    /// </summary>
    public bool TryGotoSlide(int index, out string message)
    {
        message = string.Empty;
        if (!_running || _show is null)
        {
            message = "目前沒有放映，無法跳頁。";
            return false;
        }

        int total = ShowTotalSlides;
        if (total > 0 && (index < 1 || index > total))
        {
            message = $"第 {index} 頁超出放映範圍（共 {total} 頁），已忽略。";
            return false;
        }

        try
        {
            _show.View.GotoSlide(index, 0);
            return true;
        }
        catch (Exception ex)
        {
            message = "跳頁失敗：" + ex.Message;
            return false;
        }
    }

    /// <summary>放映本身的總頁數（以放映中的那份簡報為準，避免和已換掉的檔案對不上）。</summary>
    public int ShowTotalSlides
    {
        get
        {
            try
            {
                return _running && _show is not null
                    ? Convert.ToInt32(_show.Presentation.Slides.Count)
                    : 0;
            }
            catch
            {
                return 0;
            }
        }
    }

    public void SetScreenState(int state) => Invoke(() => _show!.View.State = state);

    public int CurrentSlide
    {
        get
        {
            try
            {
                return _running && _show is not null ? Convert.ToInt32(_show.View.CurrentShowPosition) : 0;
            }
            catch
            {
                return 0;
            }
        }
    }

    public int TotalSlides
    {
        get
        {
            try
            {
                return _presentation is null ? 0 : Convert.ToInt32(_presentation.Slides.Count);
            }
            catch
            {
                return 0;
            }
        }
    }

    /// <summary>
    /// 檢查放映是否還活著。講員在投影幕按 ESC，或在 PowerPoint 端結束放映時，我們的旗標不會知道，
    /// 會留下「控制台以為還在放映」的假狀態：按鈕看似有反應但其實沒作用、控制台還掛著不搶焦點
    /// （之後連「開啟簡報」的檔案對話框都叫不出來）。定期用視窗實際狀態校正。
    /// </summary>
    public bool CheckShowAlive()
    {
        if (!_running)
        {
            return false;
        }

        if (_showHwnd == IntPtr.Zero || !Native.IsSlideShowWindow(_showHwnd))
        {
            ResetShowState();
            return false;
        }

        return true;
    }

    private void ResetShowState()
    {
        _running = false;
        _show = null;
        _showHwnd = IntPtr.Zero;
    }

    public void EndShow()
    {
        if (_show is not null)
        {
            try
            {
                _show.View.Exit();
            }
            catch
            {
                // PowerPoint 可能已被關閉
            }
        }

        ResetShowState();
    }

    /// <summary>關掉目前這份簡報（移除主 PPT 時用），不動 PowerPoint 應用程式本身。</summary>
    public void ClosePresentation()
    {
        try
        {
            // 匯出預覽或放映可能把簡報標成未儲存；若開啟時本來已儲存，關閉前先還原，
            // 才不會在關閉 PowerPoint 時跳出「是否儲存變更」。
            if (_presentation is not null && _savedOnOpen)
            {
                try
                {
                    _presentation.Saved = true;
                }
                catch
                {
                    // 還原失敗仍照常關閉
                }
            }

            _presentation?.Close();
        }
        catch
        {
            // 簡報可能已關閉，或 PowerPoint 已結束
        }

        ResetShowState();
        _presentation = null;
        PresentationPath = null;
        _savedOnOpen = true;
    }

    private void Invoke(Action action)
    {
        if (!_running || _show is null)
        {
            return;
        }

        try
        {
            action();
        }
        catch (Exception ex)
        {
            ResetShowState();
            throw new PowerPointException("放映已中斷（可能被關閉或當掉）：" + ex.Message);
        }
    }
}
