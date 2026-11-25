#requires -Version 5.1
<#
  ProjectionManager 測試模式自動化：模擬螢幕 2，跑 V2.0 控制台／手機遙控器／本輪三項修正的測試用例。
  用法：
    powershell.exe -ExecutionPolicy Bypass -File .\tests\Run-TestMode.ps1
    powershell.exe -ExecutionPolicy Bypass -File .\tests\Run-TestMode.ps1 -Presentation "D:\...\Presentation 17.pptx"
  產出：
    tests\測試結果_自動化.md、tests\artifacts\*.png
#>
param(
    [string]$Presentation = 'D:\LionCJ\OneDrive\YHHC\02 信息分享\Presentation 3.pptx',
    [int]$Port = 5180
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Add-Type @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class TestWin {
    public delegate bool EnumProc(IntPtr h, IntPtr l);

    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] public static extern int GetSystemMetrics(int index);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    public static List<IntPtr> All() {
        var list = new List<IntPtr>();
        EnumWindows((h, l) => { list.Add(h); return true; }, IntPtr.Zero);
        return list;
    }

    public static string Title(IntPtr h) {
        var sb = new StringBuilder(512);
        GetWindowText(h, sb, sb.Capacity);
        return sb.ToString();
    }

    public static string Class(IntPtr h) {
        var sb = new StringBuilder(256);
        GetClassName(h, sb, sb.Capacity);
        return sb.ToString();
    }

    public static IntPtr FindByTitle(string part) {
        foreach (var h in All()) {
            if (IsWindowVisible(h) && Title(h).Contains(part)) return h;
        }
        return IntPtr.Zero;
    }
}
'@

$root = 'D:\project\ProjectionManager'
$outDir = Join-Path $root 'tests'
$artifactDir = Join-Path $outDir 'artifacts'
New-Item -ItemType Directory -Force -Path $artifactDir | Out-Null
$testDir = Join-Path $env:TEMP 'pm_test'
New-Item -ItemType Directory -Force -Path $testDir | Out-Null
$settingsPath = Join-Path $testDir 'settings.json'
$exe = Join-Path $root 'bin\Debug\net8.0-windows\ProjectionManager.exe'

$results = New-Object System.Collections.Generic.List[object]
$progressPath = Join-Path $outDir "run_progress.log"
[IO.File]::WriteAllText($progressPath, "")

function Add-Result {
    param([string]$Id, [string]$Name, [string]$Expect, [string]$Actual, [bool]$Pass, [double]$Ms = 0)
    $results.Add([pscustomobject]@{ Id = $Id; Name = $Name; Expect = $Expect; Actual = $Actual; Pass = $Pass; Ms = [Math]::Round($Ms, 0) })
    $flag = if ($Pass) { 'PASS' } else { 'FAIL' }
    [IO.File]::AppendAllText($progressPath, ("{0} {1} {2}`r`n" -f (Get-Date -Format 'HH:mm:ss'), $Id, $flag))
    Write-Host ("[{0}] {1}  {2} / 預期 {3}" -f $flag, $Id, $Actual, $Expect)
}

function Api-State { Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/state?token={1}" -f $Port, $script:token) -TimeoutSec 5 }
function Api-Layout {
    Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/test/layout" -f $Port) -TimeoutSec 5
}
function Api-Cmd {
    param([string]$Cmd, $Arg)
    $payload = @{ token = $script:token; cmd = $Cmd }
    if ($null -ne $Arg) { $payload.arg = $Arg }
    try {
        return Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/cmd" -f $Port) -Method Post -Body ($payload | ConvertTo-Json -Depth 4) -ContentType 'application/json' -TimeoutSec 10
    } catch {
        $code = 0
        try { $code = $_.Exception.Response.StatusCode.value__ } catch { }
        return [pscustomobject]@{ ok = $false; message = $_.ErrorDetails.Message; http = $code }
    }
}
function Wait-Until {
    param([scriptblock]$Check, [int]$TimeoutMs = 2500)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
        try { if (& $Check) { return $true } } catch { }
        Start-Sleep -Milliseconds 50
    }
    return $false
}
function Ui-El {
    param([string]$Id)
    $rootEl = [Windows.Automation.AutomationElement]::FromHandle($script:hwnd)
    $cond = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, $Id)
    return $rootEl.FindFirst([Windows.Automation.TreeScope]::Descendants, $cond)
}
function Ui-Text {
    param([string]$Id)
    $el = Ui-El $Id
    if ($null -eq $el) { return '<missing>' }
    return $el.Current.Name
}
function Ui-ElWin {
    param([IntPtr]$Win, [string]$Id)
    $rootEl = [Windows.Automation.AutomationElement]::FromHandle($Win)
    $cond = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, $Id)
    return $rootEl.FindFirst([Windows.Automation.TreeScope]::Descendants, $cond)
}
function Ui-TextWin {
    param([IntPtr]$Win, [string]$Id)
    $el = Ui-ElWin $Win $Id
    if ($null -eq $el) { return '<missing>' }
    return $el.Current.Name
}
function Invoke-UiWin {
    param([IntPtr]$Win, [string]$Id)
    $el = Ui-ElWin $Win $Id
    if ($null -eq $el) { throw "找不到對話框控制項 $Id" }
    $el.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
}
function Api-Action {
    param([string]$Action)
    Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/test/action" -f $Port) -Method Post -Body (@{ action = $Action } | ConvertTo-Json) -ContentType 'application/json' -TimeoutSec 20
}
function Invoke-Ui {
    param([string]$Id)
    $el = Ui-El $Id
    if ($null -eq $el) { throw "找不到控制項 $Id" }
    $el.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
}
function Set-UiValue {
    param([string]$Id, [string]$Value)
    $el = Ui-El $Id
    if ($null -eq $el) { throw "找不到控制項 $Id" }
    $el.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue($Value)
}
function Save-Shot {
    param([string]$Name)
    Add-Type -AssemblyName System.Drawing
    if (-not ('Dpi' -as [type])) {
        Add-Type @'
using System.Runtime.InteropServices;
public class Dpi { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); }
'@
    }
[Dpi]::SetProcessDPIAware() | Out-Null
$w = [TestWin]::GetSystemMetrics(0)
$h = [TestWin]::GetSystemMetrics(1)
$bmp = New-Object System.Drawing.Bitmap($w, $h)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$path = Join-Path $artifactDir $Name
try {
    $g.CopyFromScreen(0, 0, 0, 0, $bmp.Size)
} catch {
    # 遠端／鎖屏時桌面擷取會失效：改由 WPF 自行渲染控制台視窗
    try {
        $shot = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/test/shot" -f $Port) -TimeoutSec 10
        [IO.File]::WriteAllBytes($path, [Convert]::FromBase64String($shot.png))
        $g.Dispose(); $bmp.Dispose()
        return $path
    } catch {
        $g.Clear([System.Drawing.Color]::FromArgb(20, 24, 30))
    }
}
$bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
return $path
}

# ---------------------------------------------------------------- 準備
Write-Host '=== ProjectionManager 測試模式 ==='
Get-Process ProjectionManager, POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

Add-Type @'
using System.Runtime.InteropServices;
public class DpiBoot { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); }
'@
[DpiBoot]::SetProcessDPIAware() | Out-Null

if (-not (Test-Path -LiteralPath $Presentation)) { throw "找不到測試簡報：$Presentation" }

$settings = [ordered]@{
    BoundDeviceName = $null
    TargetMonitorDevice = $null
    LastPresentation = $null
    WindowedShow = $true
    NoActivateConsole = $true
    CollapsedGroups = @()
    MobileControlEnabled = $true
    MobileControlPort = $Port
    MobileDevices = @()
}
[IO.File]::WriteAllText($settingsPath, ($settings | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))

$env:PROJECTIONMANAGER_SETTINGS = $settingsPath

$startArgs = @('--test-mode', '--auto-show', ('"{0}"' -f $Presentation))
$proc = Start-Process -FilePath $exe -ArgumentList $startArgs -PassThru
Write-Host ("啟動：{0}  PID={1}" -f $exe, $proc.Id)

$hwnd = [IntPtr]::Zero
if (-not (Wait-Until {
            $p = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue
            if ($null -eq $p) { return $false }
            $script:hwnd = $p.MainWindowHandle
            return ($script:hwnd -ne [IntPtr]::Zero -and [TestWin]::Title($script:hwnd).Contains('投影控制台'))
        } 20000)) {
    throw '主視窗未出現'
}

$swW = [TestWin]::GetSystemMetrics(0)
$swH = [TestWin]::GetSystemMetrics(1)
Write-Host ("螢幕：{0}×{1}；控制台移至左半，模擬螢幕 2 = 右半" -f $swW, $swH)
[TestWin]::SetWindowPos($hwnd, [IntPtr]::Zero, 0, 0, [int]($swW / 2), $swH, 0x0040) | Out-Null
Start-Sleep -Milliseconds 600

$title = [TestWin]::Title($hwnd)
Write-Host ("視窗標題：{0}" -f $title)
if ($title -notmatch 'PIN (\d{6})') { throw "測試模式標題沒有 PIN：$title" }
$pin = $Matches[1]

# 等 HTTP 服務
$ready = Wait-Until { try { (New-Object Net.Sockets.TcpClient('127.0.0.1', $Port)).Close(); return $true } catch { return $false } } 15000
if (-not $ready) { throw "手機服務未在 127.0.0.1:$Port 監聽" }

$pair = $null
$pairSw = [Diagnostics.Stopwatch]::StartNew()
if (-not (Wait-Until {
            try {
                $script:pair = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/pair" -f $Port) -Method Post -Body (@{ pin = $pin; name = '自動測試' } | ConvertTo-Json) -ContentType 'application/json' -TimeoutSec 5
                return $script:pair.ok
            } catch { return $false }
        } 8000)) { throw '配對失敗（PIN 不正確或服務異常）' }
$script:token = $pair.token
$deviceId = $pair.id
Write-Host ("配對成功：{0}（{1} ms）" -f $pair.name, [Math]::Round($pairSw.Elapsed.TotalMilliseconds, 0))

# J-02：手機頁面（本機以 HTTP 直取，驗證同一份 HTML）
try {
    $html = (Invoke-WebRequest -Uri ("http://127.0.0.1:{0}/" -f $Port) -UseBasicParsing -TimeoutSec 5).Content
    $j2 = ($html -like '*投影控制*' -and $html -like '*配對*')
    Add-Result 'J-02' '手機頁面可開啟' '含「投影控制」與配對表單' ("長度={0}，含標題={1}" -f $html.Length, $j2) $j2
} catch {
    Add-Result 'J-02' '手機頁面可開啟' 'HTTP 200' ('失敗：' + $_.Exception.Message) $false
}

# 等自動放映
$showReady = Wait-Until { (Api-State).showRunning } 20000
if (-not $showReady) {
    Add-Result 'K-00' '測試模式自動放映' 'showRunning=True' '未開始放映' $false
    Save-Shot 'show_failed.png' | Out-Null
    throw '自動放映未啟動'
}

# ---------------------------------------------------------------- K 組：本輪修正
$state = Api-State

$targetText = Api-Layout
$k1 = ($targetText.targetDevice -like '*TEST2*' -and $targetText.targetWidth -eq 1728 -and $targetText.targetHeight -eq 2160)
Add-Result 'K-01' '模擬螢幕 2 已成為投影目標' 'TEST2 1728×2160' ("{0} {1}×{2}" -f $targetText.targetDevice, $targetText.targetWidth, $targetText.targetHeight) $k1

$k2 = ($state.page -eq 1 -and $state.total -eq 13 -and (Ui-Text 'ShowStateText') -like '*放映中*')
Add-Result 'K-02' '放映中頁碼狀態' 'page=1 total=13 放映中' ("page={0} total={1} {2}" -f $state.page, $state.total, (Ui-Text 'ShowStateText')) $k2

$badge = Ui-Text 'PreviewLivePage'
$pageEl = Ui-El 'PreviewLivePage'
$badgeRect = if ($null -ne $pageEl) { $pageEl.Current.BoundingRectangle } else { $null }
$badgeShown = ($null -ne $badgeRect -and $badgeRect.Width -gt 10 -and $badgeRect.Height -gt 10)
$k3 = ($badge -eq '第 1 / 13 頁' -and $badgeShown)
Add-Result 'K-03' '預覽頁碼徽章' '第 1 / 13 頁（有實際版面）' ("{0}，版面={1}" -f $badge, $badgeShown) $k3

Start-Sleep -Milliseconds 900
$previewRoot = Join-Path $env:TEMP 'ProjectionManager\preview'
$shotFile = $null
if (Test-Path -LiteralPath $previewRoot) {
    $newest = Get-ChildItem -LiteralPath $previewRoot -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($newest) { $shotFile = Get-ChildItem -LiteralPath $newest.FullName -Filter 'main_p001.png' -ErrorAction SilentlyContinue | Select-Object -First 1 }
}
$previewName = Ui-Text 'PreviewLiveName'
$k4 = ($null -ne $shotFile -and $shotFile.Length -gt 1000 -and $previewName -like '*第 1 / 13 頁*')
Add-Result 'K-04' '主 PPT 靜態預覽已匯出' 'main_p001.png >1KB 且標題含頁數' ("{0}，{1} B" -f $previewName, $(if ($shotFile) { $shotFile.Length } else { 0 })) $k4

$showHwnd = [IntPtr]::Zero
foreach ($h in [TestWin]::All()) {
    if (-not [TestWin]::IsWindowVisible($h)) { continue }
    $t = [TestWin]::Title($h)
    if ($t -like '*投影片放映*' -or $t -like '*Slide Show*') {
        $r = New-Object 'TestWin+RECT'
        [TestWin]::GetWindowRect($h, [ref]$r) | Out-Null
        if (($r.Right - $r.Left) -gt 200) { $showHwnd = $h; $showRect = $r }
    }
}
if ($showHwnd -ne [IntPtr]::Zero) {
    $k5 = ([Math]::Abs($showRect.Left - ($swW / 2)) -lt 40 -and [Math]::Abs($showRect.Bottom - $swH) -lt 60)
    Add-Result 'K-05' '放映視窗落在模擬螢幕 2' ("left≈{0}" -f ($swW / 2)) ("{0},{1} {2}×{3}" -f $showRect.Left, $showRect.Top, ($showRect.Right - $showRect.Left), ($showRect.Bottom - $showRect.Top)) $k5
} else {
    Add-Result 'K-05' '放映視窗落在模擬螢幕 2' '找到放映視窗' '找不到' $false
}

$sw = [Diagnostics.Stopwatch]::StartNew()
$r = Api-Cmd 'goto' 10
$ok = Wait-Until { (Ui-Text 'PreviewLivePage') -eq '第 10 / 13 頁' -and (Ui-Text 'PageText') -eq '10' } 3000
$ms = $sw.Elapsed.TotalMilliseconds
Add-Result 'K-06' '輸入頁碼跳頁（10）' 'page=10，徽章 10/13，<1.5s' ("ok={0}，頁碼={1}，{2} ms" -f $r.ok, (Ui-Text 'PageText'), [Math]::Round($ms)) ($ok -and $ms -lt 1500) $ms
Save-Shot 'k06_page10.png' | Out-Null

Set-UiValue 'PageJumpBox' '13'
$sw = [Diagnostics.Stopwatch]::StartNew()
Invoke-Ui 'JumpButton'
$ok = Wait-Until { (Api-State).page -eq 13 } 3000
$ms = $sw.Elapsed.TotalMilliseconds
Add-Result 'K-07' '頁碼輸入框＋跳至按鈕' 'page=13（UIA 點擊，<3s）' ("頁碼={0}，{1} ms" -f (Api-State).page, [Math]::Round($ms)) ($ok -and $ms -lt 3000) $ms

$r = Api-Cmd 'goto' 14
Start-Sleep -Milliseconds 500
$pageAfter = (Api-State).page
Add-Result 'K-08' '越界頁碼不崩潰' 'ok=False，page 不變 13' ("ok={0}，page={1}，msg={2}" -f $r.ok, $pageAfter, $r.message) ((-not $r.ok) -and $pageAfter -eq 13)

$maxMs = 0.0
$allOk = $true
foreach ($p in @(1, 13, 1, 13)) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Api-Cmd 'goto' $p | Out-Null
    $hit = Wait-Until { (Api-State).page -eq $p } 3000
    $ms = $sw.Elapsed.TotalMilliseconds
    if ($ms -gt $maxMs) { $maxMs = $ms }
    if (-not $hit) { $allOk = $false }
}
Add-Result 'K-09' '大幅跳轉來回四次' '全部命中，最慢 <1.5s' ("max={0} ms，全部={1}" -f [Math]::Round($maxMs), $allOk) ($allOk -and $maxMs -lt 1500) $maxMs

# K-10：預覽區收窄（版面比例）；K-11：預覽圖不被拉扁且縮圖條夠高
$lay = Api-Layout
$k10 = ($lay.previewBodyAspect -ge 1.28 -and $lay.previewBodyAspect -le 1.40 -and $lay.previewRatio -le 0.43 -and $lay.previewRatio -ge 0.12 -and $lay.thumbsH -ge ($lay.previewH * 1.1))
Add-Result 'K-10' '預覽區維持 4:3 且不佔滿畫面' '預覽本體 4:3（1.28–1.40）、≤43%，縮圖≥預覽 1.1 倍' ("預覽={0}px（{1:P0}），本體 {2}×{3}（{4}），縮圖={5}px（{6:P0}），合計={7}px" -f $lay.previewH, $lay.previewRatio, $lay.previewBodyW, $lay.previewBodyH, $lay.previewBodyAspect, $lay.thumbsH, $lay.thumbsRatio, $lay.centerH) $k10

$srcAsp = if ($lay.mainAspect -gt 0) { $lay.mainAspect } else { 1.333 }
$imgAsp = if ($lay.imagePixelAspect -gt 0) { $lay.imagePixelAspect } else { $lay.liveAspect }
$k11 = ([Math]::Abs($imgAsp - $srcAsp) -le 0.03 -and $lay.liveW -ge ($lay.cardW * 0.5) -and $lay.liveW -le ($lay.cardW + 4) -and $lay.thumbCount -ge 8)
Add-Result 'K-11' '預覽圖比例與簡報來源一致' '預覽圖像素比例≈來源（±0.03）、不超出卡片、縮圖≥8 張' ("來源比例 {0}，預覽圖 {1}×{2}（像素比例 {3}，顯示比例 {4}，卡片寬 {5}），縮圖 {6} 張" -f $srcAsp, $lay.imagePixelW, $lay.imagePixelH, $imgAsp, $lay.liveAspect, $lay.cardW, $lay.thumbCount) $k11

# K-12：跳頁後版面尺寸不變（預覽不會因換頁又撐大）
Api-Cmd 'goto' 10 | Out-Null
Wait-Until { (Api-State).page -eq 10 } 2000 | Out-Null
Start-Sleep -Milliseconds 300
$lay2 = Api-Layout
$k12 = ([Math]::Abs($lay2.previewH - $lay.previewH) -le 3 -and [Math]::Abs($lay2.thumbsH - $lay.thumbsH) -le 3)
Add-Result 'K-12' '換頁後版面尺寸穩定' '預覽與縮圖高度不變（±3px）' ("預覽 {0}→{1}，縮圖 {2}→{3}" -f $lay.previewH, $lay2.previewH, $lay.thumbsH, $lay2.thumbsH) $k12
Save-Shot 'k12_layout.png' | Out-Null

# ---------------------------------------------------------------- J 組：手機遙控器
$r = Api-Cmd 'goto' 1
Wait-Until { (Api-State).page -eq 1 } 2000 | Out-Null
$r = Api-Cmd 'next' $null
Start-Sleep -Milliseconds 400
$p = (Api-State).page
Add-Result 'J-06a' '主 PPT 下一頁' '1→2' ("page={0}" -f $p) ($p -eq 2)

$r = Api-Cmd 'prev' $null
Start-Sleep -Milliseconds 400
$p = (Api-State).page
Add-Result 'J-06b' '主 PPT 上一頁' '2→1' ("page={0}" -f $p) ($p -eq 1)

$r = Api-Cmd 'goto' 5
Start-Sleep -Milliseconds 400
$p = (Api-State).page
Add-Result 'J-07a' '手機跳頁到 5' 'page=5' ("ok={0} page={1}" -f $r.ok, $p) ($r.ok -and $p -eq 5)

$r = Api-Cmd 'black' $null
Start-Sleep -Milliseconds 300
$black = (Api-State).black
Add-Result 'J-08a' '手機黑屏' 'black=True' ("black={0}" -f $black) $black

$r = Api-Cmd 'resume' $null
Start-Sleep -Milliseconds 300
$black = (Api-State).black
Add-Result 'J-08b' '手機恢復畫面' 'black=False' ("black={0}" -f $black) (-not $black)

$state = Api-State
$main = $state.sources | Where-Object { $_.kind -eq 'main' } | Select-Object -First 1
$j6 = ($state.sources.Count -ge 1 -and $main -ne $null -and $main.live)
Add-Result 'J-14a' '投影源清單與即時狀態' '至少 1 個源，main=live' ("{0} 個，main.live={1}" -f $state.sources.Count, $(if ($main) { $main.live } else { 'n/a' })) $j6

# J-09：切換投影源（操作員）
$idx17 = ($state.sources | Where-Object { $_.name -like 'Presentation 17*' } | Select-Object -First 1).index
if ($null -ne $idx17) {
    $r = Api-Cmd 'source' $idx17
    Start-Sleep -Milliseconds 600
    $st = Api-State
    $j9 = ($r.ok -and $st.sourceKind -eq 'material' -and $st.total -eq 14 -and $st.page -eq 1)
    Add-Result 'J-09' '切換投影源' 'Presentation 17，1/14' ("{0}，{1}/{2}" -f $st.sourceName, $st.page, $st.total) $j9
} else {
    Add-Result 'J-09' '切換投影源' '找到 Presentation 17' '清單沒有此素材' $false
}

# J-05：素材上下頁
$idx3 = ($st.sources | Where-Object { $_.kind -eq 'material' -and $_.name -eq 'Presentation 3' } | Select-Object -First 1).index
if ($null -ne $idx3) {
    $rSrc = Api-Cmd 'source' $idx3
    Start-Sleep -Milliseconds 500
    $stMat = Api-State
    Api-Cmd 'next' $null | Out-Null
    Start-Sleep -Milliseconds 400
    $p2 = (Api-State).page
    Api-Cmd 'prev' $null | Out-Null
    Start-Sleep -Milliseconds 400
    $p1 = (Api-State).page
    $j5 = ($p2 -eq 2 -and $p1 -eq 1 -and $stMat.sourceKind -eq 'material' -and $stMat.total -eq 13)
    Add-Result 'J-05' '素材上下頁' '1→2→1' ("源={0}/{1}，next={2}，prev={3}" -f $stMat.sourceKind, $stMat.total, $p2, $p1) $j5
    Save-Shot 'k13_material_thumbs.png' | Out-Null
} else {
    Add-Result 'J-05' '素材上下頁' '找到 Presentation 3 素材' '清單沒有此素材' $false
}

# ---------------------------------------------------------------- N 組：九項改造
$lay = Api-Layout
$n1 = ($lay.defaultGroupName -eq '默認分組' -and $lay.defaultGroupCount -eq 1)
Add-Result 'N-01' '預設分組名稱與唯一性' '名稱「默認分組」，只有一個預設分組' ("名稱={0}，數量={1}" -f $lay.defaultGroupName, $lay.defaultGroupCount) $n1

$stMat = Api-State
$n2 = (-not $stMat.projected -and (Ui-Text 'ShowStateText') -like '*已選擇*' -and $lay.pushVisible -and $lay.pushText -eq '推送投影')
Add-Result 'N-02' '素材切換後預設不投影' 'projected=False，狀態「已選擇」，按鈕「推送投影」' ("projected={0}，狀態={1}，可見={2}，文案={3}" -f $stMat.projected, (Ui-Text 'ShowStateText'), $lay.pushVisible, $lay.pushText) $n2

$rPush = Api-Cmd 'push' $null
$pushed = Wait-Until { (Api-State).projected -and (Ui-Text 'ShowStateText') -like '*投影中*' } 3000
$layPushed = Api-Layout
$n3 = ($rPush.ok -and $pushed -and $layPushed.pushText -eq '更新投影')
Add-Result 'N-03' '素材按推送投影才上螢幕 2' 'projected=True，狀態「投影中」，按鈕「更新投影」' ("ok={0}，狀態={1}，文案={2}" -f $rPush.ok, (Ui-Text 'ShowStateText'), $layPushed.pushText) $n3

$pBefore = (Api-State).page
Api-Cmd 'next' $null | Out-Null
Start-Sleep -Milliseconds 300
$pAfter = (Api-State).page
$n4 = ($pBefore -eq 1 -and $pAfter -eq 2)
Add-Result 'N-04' '推送後素材可翻頁' '1→2' ("{0}→{1}" -f $pBefore, $pAfter) $n4

$actScripture = Api-Action 'seed-scripture'
Start-Sleep -Milliseconds 500
$stScr = Api-State
$n5 = ($stScr.sourceKind -eq 'scripture' -and -not $stScr.projected -and (Ui-Text 'ShowStateText') -like '*待推送*')
Add-Result 'N-05' '經文先加入、待推送' 'sourceKind=scripture，projected=False，狀態含「待推送」' ("kind={0}，projected={1}，狀態={2}" -f $stScr.sourceKind, $stScr.projected, (Ui-Text 'ShowStateText')) $n5

$rPushScr = Api-Cmd 'push' $null
$scrPushed = Wait-Until { (Api-State).projected } 5000
if (-not $scrPushed) {
    Start-Sleep -Milliseconds 600
    $scrPushed = (Api-State).projected
}
$n6 = ($rPushScr.ok -and $scrPushed -and (Ui-Text 'ShowStateText') -like '*投影中*')
Add-Result 'N-06' '經文按推送投影才上螢幕 2' 'projected=True，狀態「投影中」' ("ok={0}，projected={1}，狀態={2}" -f $rPushScr.ok, (Api-State).projected, (Ui-Text 'ShowStateText')) $n6

$actWeb = Api-Action 'seed-web'
Start-Sleep -Milliseconds 600
$stWeb = Api-State
$n7 = ($stWeb.sourceKind -eq 'web' -and -not $stWeb.projected -and (Ui-Text 'ShowStateText') -like '*待推送*')
Add-Result 'N-07' '網頁先加入、待推送' 'sourceKind=web，projected=False' ("kind={0}，projected={1}，狀態={2}" -f $stWeb.sourceKind, $stWeb.projected, (Ui-Text 'ShowStateText')) $n7

$rPushWeb = Api-Cmd 'push' $null
$webPushed = Wait-Until { (Api-State).projected } 3000
$n8 = ($rPushWeb.ok -and $webPushed)
Add-Result 'N-08' '網頁按推送投影才上螢幕 2' 'projected=True' ("ok={0}，projected={1}" -f $rPushWeb.ok, (Api-State).projected) $n8

Start-Sleep -Milliseconds 500
$actZoom = Api-Action 'select-zoom'
Start-Sleep -Milliseconds 500
$layZoom = Api-Layout
$n9 = ($actZoom.ok -and $layZoom.zoomPanelVisible -and -not $actZoom.projected)
Add-Result 'N-09' 'Zoom 右欄面板與先選後推' '面板可見；未按推送前 projected=False' ("ok={0}，panel={1}，projected={2}，running={3}，windows={4}" -f $actZoom.ok, $layZoom.zoomPanelVisible, $actZoom.projected, $actZoom.zoomRunning, $actZoom.zoomWindows) $n9

$n10 = ($lay.topButtonsSized -and $lay.navBottomAligned -and $lay.navPrevW -ge 100 -and $lay.navNextW -ge 100 -and $lay.navPushW -ge 100 -and $lay.refreshMonitorsStretched)
Add-Result 'N-10' '按鈕高度、底線對齊與右欄全寬' '頂部按鈕高度≥32、導航同底線、Push≥100、重新整理螢幕全寬' ("top={0}，底線對齊={1}，prev/next/push={2}/{3}/{4}，refresh={5}" -f $lay.topButtonsSized, $lay.navBottomAligned, $lay.navPrevW, $lay.navNextW, $lay.navPushW, $lay.refreshMonitorsW) $n10

$previewAspect = if ($lay.previewLiveH -gt 0) { $lay.previewLiveW / $lay.previewLiveH } else { 0 }
$n11 = ($previewAspect -ge 1.28 -and $previewAspect -le 1.42 -and $lay.previewLiveW -le ($lay.cardW + 4))
Add-Result 'N-11' '預覽維持 4:3 並完整顯示' '預覽顯示比例 1.28–1.42，不超出卡片' ("{0}×{1}（{2:N2}），卡片 {3}" -f $lay.previewLiveW, $lay.previewLiveH, $previewAspect, $lay.cardW) $n11

Add-Result 'N-12' '調試模式徽章' '可見且文字含「模擬」' ("visible={0}，text={1}" -f (Api-Layout).debugBadgeVisible, (Ui-Text 'DebugBadgeText')) ((Api-Layout).debugBadgeVisible -and (Ui-Text 'DebugBadgeText') -like '*模擬*')

Invoke-Ui 'MobileButton'
$mobileHwnd = [IntPtr]::Zero
if (Wait-Until { $script:mobileHwnd = [TestWin]::FindByTitle('手機控制'); return ($script:mobileHwnd -ne [IntPtr]::Zero) } 5000) {
    Start-Sleep -Milliseconds 500
    $svcText = Ui-TextWin $mobileHwnd 'ServiceStateText'
    $netText = Ui-TextWin $mobileHwnd 'NetworkText'
    $n13 = ($svcText -like '*運行中*' -and $netText -notlike '*<missing>*' -and $netText -like '*本機網段*')
    Add-Result 'N-13' '手機控制服務狀態與網段顯示' '服務運行中、顯示本機網段' ("{0}；{1}" -f $svcText, $netText) $n13

    Invoke-UiWin $mobileHwnd 'SelfTestButton'
    $selfOk = Wait-Until { (Ui-TextWin $mobileHwnd 'ServiceDetailText') -like '*本機自測*' } 8000
    $selfText = Ui-TextWin $mobileHwnd 'ServiceDetailText'
    Add-Result 'N-14' '手機服務本機自測' '顯示本機自測結果' $selfText $selfOk
    try {
        $wp = [Windows.Automation.AutomationElement]::FromHandle($mobileHwnd).GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern)
        $wp.Close()
    } catch { }
} else {
    Add-Result 'N-13' '手機控制服務狀態與網段顯示' '開啟手機控制對話框' '找不到對話框' $false
    Add-Result 'N-14' '手機服務本機自測' '顯示本機自測結果' '找不到對話框' $false
}

# 手機投影源清單與頁碼切換
$mstate = Api-State
$kinds = @($mstate.sources | ForEach-Object { $_.kind })
$tail = if ($kinds.Count -ge 3) { $kinds[($kinds.Count - 3)..($kinds.Count - 1)] } else { @() }
$n15 = ($kinds.Count -ge 4 -and $kinds[0] -eq 'main' -and $kinds -contains 'material' -and $kinds -contains 'scripture' -and $kinds -contains 'web' -and $kinds -contains 'zoom' -and $tail.Count -eq 3 -and $tail[0] -eq 'scripture' -and $tail[1] -eq 'web' -and $tail[2] -eq 'zoom')
Add-Result 'N-15' '手機投影源涵蓋五類且順序正確' '主PPT→素材→經文→網頁→Zoom' ("kinds={0}" -f ($kinds -join ',')) $n15

$matIndex = ($mstate.sources | Where-Object { $_.kind -eq 'material' } | Select-Object -First 1).index
if ($null -ne $matIndex) {
    $rPageSwitch = Api-Cmd 'source' ("{0}|3" -f $matIndex)
    Start-Sleep -Milliseconds 500
    $stPageSwitch = Api-State
    $n16 = ($rPageSwitch.ok -and $stPageSwitch.sourceKind -eq 'material' -and $stPageSwitch.page -eq 3 -and -not $stPageSwitch.projected)
    Add-Result 'N-16' '手機切換素材可指定頁數且預設不投影' 'sourceKind=material，page=3，projected=False' ("ok={0}，kind={1}，page={2}，projected={3}" -f $rPageSwitch.ok, $stPageSwitch.sourceKind, $stPageSwitch.page, $stPageSwitch.projected) $n16
} else {
    Add-Result 'N-16' '手機切換素材可指定頁數且預設不投影' '找到素材來源' '清單沒有素材' $false
}

$mobileHtml = ''
try { $mobileHtml = (Invoke-WebRequest -Uri ("http://127.0.0.1:{0}/" -f $Port) -UseBasicParsing -TimeoutSec 5).Content } catch { }
$n17 = ($mobileHtml -like '*switchPage*' -and $mobileHtml -like '*再次確認*' -and $mobileHtml -like '*String(item.index)*')
Add-Result 'N-17' '手機切換投影源有頁碼與二次確認' '含頁碼輸入、再次確認、index|page 指令' ("html={0} bytes，pageInput={1}，confirm={2}" -f $mobileHtml.Length, ($mobileHtml -like '*switchPage*'), ($mobileHtml -like '*再次確認*')) $n17

Save-Shot 'n_features.png' | Out-Null

# 回到主 PPT，後續角色與結束放映測在同一投影源上
Api-Cmd 'source' 0 | Out-Null
Start-Sleep -Milliseconds 400

$r = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/test/role" -f $Port) -Method Post -Body (@{ id = $deviceId; role = 'speaker' } | ConvertTo-Json) -ContentType 'application/json' -TimeoutSec 5
Start-Sleep -Milliseconds 400
$state = Api-State
$j7 = ($state.role -eq 'speaker' -and $state.sources.Count -eq 0)
$rGoto = Api-Cmd 'goto' 10
$j7b = ($rGoto.http -eq 403 -or -not $rGoto.ok)
Add-Result 'J-11a' '講員角色閘門' 'role=speaker，sources=0，goto 被拒' ("role={0}，sources={1}，goto HTTP={2}" -f $state.role, $state.sources.Count, $rGoto.http) ($j7 -and $j7b)

Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/test/role" -f $Port) -Method Post -Body (@{ id = $deviceId; role = 'operator' } | ConvertTo-Json) -ContentType 'application/json' -TimeoutSec 5 | Out-Null
Start-Sleep -Milliseconds 300

$r = Api-Cmd 'next' $null
Start-Sleep -Milliseconds 150
$r2 = Api-Cmd 'next' $null
Start-Sleep -Milliseconds 500
$p = (Api-State).page
$j8 = ($p -ge 6 -and $r2.message -like '*略過*')
Add-Result 'J-13' '連續兩次下一頁（防連點）' '只前進 1 頁，第二次回「已略過」' ("page={0}，第二次={1}" -f $p, $r2.message) $j8

try {
    Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/state" -f $Port) -TimeoutSec 5 | Out-Null
    Add-Result 'J-03b' '未配對讀取狀態' 'HTTP 401' '竟然成功' $false
} catch {
    $code = $_.Exception.Response.StatusCode.value__
    Add-Result 'J-03b' '未配對讀取狀態' 'HTTP 401' ("HTTP {0}" -f $code) ($code -eq 401)
}

try {
    Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/pair" -f $Port) -Method Post -Body (@{ pin = '000000'; name = '壞人' } | ConvertTo-Json) -ContentType 'application/json' -TimeoutSec 5 | Out-Null
    Add-Result 'J-03c' '錯誤 PIN 不可配對' 'HTTP 403' '竟然成功' $false
} catch {
    $code = $_.Exception.Response.StatusCode.value__
    Add-Result 'J-03c' '錯誤 PIN 不可配對' 'HTTP 403' ("HTTP {0}" -f $code) ($code -eq 403)
}

Save-Shot 'k09_final.png' | Out-Null
$r = Api-Cmd 'stop' $null
Start-Sleep -Milliseconds 700
$state = Api-State
Add-Result 'J-10' '手機結束放映' 'showRunning=False' ("ok={0}，showRunning={1}" -f $r.ok, $state.showRunning) ($r.ok -and -not $state.showRunning)

# ---------------------------------------------------------------- L 組：全功能邊界與異常
$pageBefore = (Api-State).page

$r = Api-Cmd 'warp' $null
$l1 = (-not $r.ok -and $r.message -like '*不支援*')
Add-Result 'L-01' '未知指令被拒絕' 'ok=False，訊息含「不支援」' ("ok={0}，msg={1}" -f $r.ok, $r.message) $l1

$r = Api-Cmd 'goto' 'abc'
$pageAfter = (Api-State).page
$l2 = (-not $r.ok -and $pageAfter -eq $pageBefore -and $r.message -like '*格式*')
Add-Result 'L-02' '非數字頁碼被拒絕' 'ok=False，頁碼不變，訊息含「格式」' ("ok={0}，page={1}，msg={2}" -f $r.ok, $pageAfter, $r.message) $l2

$r = Api-Cmd 'goto' 0
$pageAfter = (Api-State).page
$l3 = (-not $r.ok -and $pageAfter -eq $pageBefore -and $r.message -like '*格式*')
Add-Result 'L-03' '頁碼 0 被拒絕' 'ok=False，頁碼不變' ("ok={0}，page={1}，msg={2}" -f $r.ok, $pageAfter, $r.message) $l3

try {
    Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/cmd" -f $Port) -Method Post -Body (@{ cmd = 'next' } | ConvertTo-Json) -ContentType 'application/json' -TimeoutSec 5 | Out-Null
    Add-Result 'L-04' '無 token 執行指令' 'HTTP 401' '竟然成功' $false
} catch {
    $code = $_.Exception.Response.StatusCode.value__
    Add-Result 'L-04' '無 token 執行指令' 'HTTP 401' ("HTTP {0}" -f $code) ($code -eq 401)
}

$r = Api-Cmd 'stop' $null
$l5 = (-not $r.ok -and $r.message -like '*沒有進行中*')
Add-Result 'L-05' '重複結束放映' 'ok=False，訊息含「沒有進行中」' ("ok={0}，msg={1}" -f $r.ok, $r.message) $l5

$r = Api-Cmd 'next' $null
$pageAfter = (Api-State).page
$l6 = (-not $r.ok -and $pageAfter -eq $pageBefore -and $r.message -like '*尚未放映*')
Add-Result 'L-06' '結束放映後翻頁' 'ok=False，頁碼不變，訊息含「尚未放映」' ("ok={0}，page={1}，msg={2}" -f $r.ok, $pageAfter, $r.message) $l6

$lay = Api-Layout
$l7 = ($lay.thumbCols -ge 2 -and $lay.thumbTileW -ge 130 -and $lay.thumbTileW -le 240 -and $lay.thumbTileH -gt 70 -and $lay.thumbCount -eq 13)
Add-Result 'L-07' '版面端點欄數與圖磚' '欄數≥2、磚寬 130–240、13 張' ("{0} 欄，{1}×{2}，{3} 張" -f $lay.thumbCols, $lay.thumbTileW, $lay.thumbTileH, $lay.thumbCount) $l7

try {
    $shot = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/test/shot" -f $Port) -TimeoutSec 15
    $bytes = [Convert]::FromBase64String($shot.png)
    $png = ($bytes.Length -gt 20000 -and $bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47)
    Add-Result 'L-08' '自渲染截圖端點' 'PNG 檔頭、>20KB' ("{0} B，PNG={1}" -f $bytes.Length, $png) $png
} catch {
    Add-Result 'L-08' '自渲染截圖端點' 'PNG >20KB' ('失敗：' + $_.Exception.Message) $false
}

$thumbFiles = 0
if (Test-Path -LiteralPath $previewRoot) {
    $newest = Get-ChildItem -LiteralPath $previewRoot -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($newest) {
        $thumbFiles = (Get-ChildItem -LiteralPath $newest.FullName -Filter 'thumb_p*.png' -ErrorAction SilentlyContinue).Count
    }
}
Add-Result 'L-09' '主 PPT 縮圖預取' '13 張小縮圖已產生' ("{0} 張" -f $thumbFiles) ($thumbFiles -ge 13)

$layoutBefore = Api-Layout
[TestWin]::SetWindowPos($hwnd, [IntPtr]::Zero, 0, 0, [int]($swW * 0.9), $swH, 0x0040) | Out-Null
Start-Sleep -Milliseconds 1200
$layoutWide = Api-Layout
$l10 = ($layoutWide.thumbCols -gt $layoutBefore.thumbCols -and $layoutWide.thumbTileW -ge 130)
Add-Result 'L-10' '拉寬視窗自動增加縮圖欄數' '欄數變多、磚寬仍≥130' ("{0}→{1} 欄，磚 {2}→{3}" -f $layoutBefore.thumbCols, $layoutWide.thumbCols, $layoutBefore.thumbTileW, $layoutWide.thumbTileW) $l10

[TestWin]::SetWindowPos($hwnd, [IntPtr]::Zero, 0, 0, [int]($swW / 2), $swH, 0x0040) | Out-Null
Start-Sleep -Milliseconds 1200
$layoutBack = Api-Layout
$l11 = ([Math]::Abs($layoutBack.thumbCols - $layoutBefore.thumbCols) -le 1)
Add-Result 'L-11' '縮回視窗欄數復原' '欄數回到原值（±1）' ("{0}→{1}→{2} 欄" -f $layoutBefore.thumbCols, $layoutWide.thumbCols, $layoutBack.thumbCols) $l11

$modeBefore = Ui-Text 'ControllerText'
Invoke-Ui 'HandoverButton'
Start-Sleep -Milliseconds 400
$modeAfter = Ui-Text 'ControllerText'
$handoverBtn = Ui-Text 'HandoverButton'
$l12 = ($modeBefore -eq '操作員' -and $modeAfter -eq '講員' -and $handoverBtn -eq '收回控制')
Add-Result 'L-12' '控制者交給講員／收回' '操作員→講員，按鈕變「收回控制」' ("{0}→{1}，按鈕={2}" -f $modeBefore, $modeAfter, $handoverBtn) $l12
Invoke-Ui 'HandoverButton'
Start-Sleep -Milliseconds 300
$modeBack = Ui-Text 'ControllerText'
Add-Result 'L-12b' '收回控制' '回到操作員' ("{0}，按鈕={1}" -f $modeBack, (Ui-Text 'HandoverButton')) ($modeBack -eq '操作員')

$logEl = Ui-El 'LogBox'
$logBefore = if ($null -ne $logEl) { $logEl.Current.BoundingRectangle.Height } else { 0 }
Invoke-Ui 'LogToggleButton'
Start-Sleep -Milliseconds 400
$logEl = Ui-El 'LogBox'
$logAfter = if ($null -ne $logEl) { $logEl.Current.BoundingRectangle.Height } else { 0 }
Add-Result 'L-13' '日誌面板開關' '點「日誌」後日誌高度 > 50' ("{0}→{1}" -f $logBefore, $logAfter) ($logAfter -gt 50)
Invoke-Ui 'LogToggleButton'
Start-Sleep -Milliseconds 300

$monEl = Ui-El 'MonitorCombo'
$monBefore = if ($null -ne $monEl) { $monEl.Current.BoundingRectangle.Height } else { 0 }
Invoke-Ui 'ShowSectionHeader'
Start-Sleep -Milliseconds 400
$monEl2 = Ui-El 'MonitorCombo'
$monAfter = if ($null -ne $monEl2) { $monEl2.Current.BoundingRectangle.Height } else { 0 }
$headerText = Ui-Text 'ShowSectionHeader'
$l14 = ($monBefore -gt 20 -and $monAfter -eq 0)
Add-Result 'L-14' '右欄分區收合' '放映設定收合後螢幕下拉消失' ("高度 {0}→{1}，標題={2}" -f $monBefore, $monAfter, $headerText) $l14
Invoke-Ui 'ShowSectionHeader'
Start-Sleep -Milliseconds 300

# ---------------------------------------------------------------- 收尾
Get-Process ProjectionManager -ErrorAction SilentlyContinue | Stop-Process -Force
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Remove-Item Env:\PROJECTIONMANAGER_SETTINGS -ErrorAction SilentlyContinue

$pass = ($results | Where-Object { $_.Pass }).Count
$total = $results.Count
$md = New-Object System.Text.StringBuilder
[void]$md.AppendLine('# ProjectionManager 測試模式自動化結果')
[void]$md.AppendLine('')
[void]$md.AppendLine(('- 執行時間：{0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
[void]$md.AppendLine(('- 簡報：{0}' -f $Presentation))
[void]$md.AppendLine(('- 螢幕：主螢幕 {0}×{1}；模擬螢幕 2 = 右半 {2}×{1}' -f $swW, $swH, ($swW / 2)))
[void]$md.AppendLine(('- 結果：{0} / {1} 通過' -f $pass, $total))
[void]$md.AppendLine('')
[void]$md.AppendLine('| 編號 | 用例 | 預期 | 實際 | 結果 | 耗時 ms |')
[void]$md.AppendLine('|---|---|---|---|---|---|')
foreach ($x in $results) {
    [void]$md.AppendLine(('| {0} | {1} | {2} | {3} | {4} | {5} |' -f $x.Id, $x.Name, $x.Expect, $x.Actual, $(if ($x.Pass) { 'PASS' } else { 'FAIL' }), $x.Ms))
}
$mdPath = Join-Path $outDir '測試結果_自動化.md'
[IO.File]::WriteAllText($mdPath, $md.ToString(), (New-Object Text.UTF8Encoding($true)))

Write-Host ''
Write-Host ("=== {0} / {1} 通過，結果：{2} ===" -f $pass, $total, $mdPath)
if ($pass -lt $total) { exit 1 }
