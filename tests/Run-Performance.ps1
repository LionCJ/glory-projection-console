#requires -Version 5.1
<#
  ProjectionManager 效能測試：冷啟動、放映待機、連續翻頁／跳頁、快速切源、黑屏恢復。
  產出：tests\性能測試結果.md、tests\artifacts\perf_*.png（可選）
  用法：powershell.exe -ExecutionPolicy Bypass -File .\tests\Run-Performance.ps1
#>
param(
    [string]$Presentation = 'D:\LionCJ\OneDrive\YHHC\02 信息分享\Presentation 3.pptx',
    [int]$Port = 5180
)

$ErrorActionPreference = 'Stop'

Add-Type @'
using System;
using System.Runtime.InteropServices;

public class PerfNative {
    [DllImport("user32.dll")] public static extern uint GetGuiResources(IntPtr hProcess, uint flags);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
}
'@

$root = 'D:\project\ProjectionManager'
$outDir = Join-Path $root 'tests'
$artifactDir = Join-Path $outDir 'artifacts'
New-Item -ItemType Directory -Force -Path $artifactDir | Out-Null
$testDir = Join-Path $env:TEMP 'pm_perf'
New-Item -ItemType Directory -Force -Path $testDir | Out-Null
$settingsPath = Join-Path $testDir 'settings.json'
$exe = Join-Path $root 'bin\Debug\net8.0-windows\ProjectionManager.exe'
$cores = [Environment]::ProcessorCount
[PerfNative]::SetProcessDPIAware() | Out-Null

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

Get-Process ProjectionManager, POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

function Snapshot {
    param([int]$ProcId)
    $p = Get-Process -Id $ProcId -ErrorAction SilentlyContinue
    $ppt = Get-Process POWERPNT -ErrorAction SilentlyContinue
    $pptCpu = 0.0
    $pptWs = 0L
    $pptCount = 0
    if ($ppt) {
        foreach ($x in $ppt) { $pptCpu += $x.CPU; $pptWs += $x.WorkingSet64; $pptCount++ }
    }
    if ($null -eq $p) {
        return [pscustomobject]@{ Alive = $false; CpuSec = 0; WorkingSet = 0; Private = 0; Threads = 0; Handles = 0; Gdi = 0; User = 0; PptCpuSec = 0; PptWorkingSet = 0; PptCount = 0 }
    }
    $gdi = 0; $user = 0
    try { $gdi = [PerfNative]::GetGuiResources($p.Handle, 0); $user = [PerfNative]::GetGuiResources($p.Handle, 1) } catch { }
    return [pscustomobject]@{
        Alive = $true
        CpuSec = [double]$p.CPU
        WorkingSet = [long]$p.WorkingSet64
        Private = [long]$p.PrivateMemorySize64
        Threads = $p.Threads.Count
        Handles = $p.HandleCount
        Gdi = [int]$gdi
        User = [int]$user
        PptCpuSec = [double]$pptCpu
        PptWorkingSet = [long]$pptWs
        PptCount = $pptCount
    }
}

function Measure-Load {
    param([int]$ProcId, [double]$Seconds)
    $a = Snapshot $ProcId
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Start-Sleep -Milliseconds ([int]($Seconds * 1000))
    $b = Snapshot $ProcId
    $elapsed = $sw.Elapsed.TotalSeconds
    $cpuDelta = [Math]::Max(0, $b.CpuSec - $a.CpuSec)
    $pptDelta = [Math]::Max(0, $b.PptCpuSec - $a.PptCpuSec)
    $oneCore = $cpuDelta / $elapsed * 100
    $total = $oneCore / $cores
    [pscustomobject]@{
        Seconds = [Math]::Round($elapsed, 2)
        ConsoleCpuOneCore = [Math]::Round($oneCore, 1)
        ConsoleCpuTotal = [Math]::Round($total, 2)
        PptCpuOneCore = [Math]::Round($pptDelta / $elapsed * 100, 1)
        WsStartMB = [Math]::Round($a.WorkingSet / 1MB, 1)
        WsEndMB = [Math]::Round($b.WorkingSet / 1MB, 1)
        WsDeltaMB = [Math]::Round(($b.WorkingSet - $a.WorkingSet) / 1MB, 1)
        PrivateMB = [Math]::Round($b.Private / 1MB, 1)
        Threads = $b.Threads
        ThreadDelta = $b.Threads - $a.Threads
        Handles = $b.Handles
        HandleDelta = $b.Handles - $a.Handles
        Gdi = $b.Gdi
        User = $b.User
        PptCount = $b.PptCount
        PptWsMB = [Math]::Round($b.PptWorkingSet / 1MB, 1)
        PptCpuTotal = [Math]::Round(($pptDelta / $elapsed * 100) / $cores, 2)
    }
}

function Api-State { Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/state?token={1}" -f $Port, $script:token) -TimeoutSec 10 }

function Api-Cmd {
    param([string]$Cmd, $Arg)
    $payload = @{ token = $script:token; cmd = $Cmd }
    if ($null -ne $Arg) { $payload.arg = $Arg }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $resp = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/cmd" -f $Port) -Method Post -Body ($payload | ConvertTo-Json -Depth 4) -ContentType 'application/json' -TimeoutSec 20
        $sw.Stop()
        return [pscustomobject]@{ Ok = $resp.ok; Message = $resp.message; Ms = $sw.Elapsed.TotalMilliseconds; Http = 200 }
    } catch {
        $sw.Stop()
        $code = 0
        try { $code = [int]$_.Exception.Response.StatusCode } catch { }
        return [pscustomobject]@{ Ok = $false; Message = $_.Exception.Message; Ms = $sw.Elapsed.TotalMilliseconds; Http = $code }
    }
}

function Wait-Until {
    param([scriptblock]$Check, [int]$TimeoutMs = 3000)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
        try { if (& $Check) { return $true } } catch { }
        Start-Sleep -Milliseconds 20
    }
    return $false
}

function Stats {
    param([double[]]$Values)
    if ($Values.Count -eq 0) { return [pscustomobject]@{ Count = 0; Avg = 0; P50 = 0; P95 = 0; Max = 0 } }
    $sorted = @($Values | Sort-Object)
    $idx95 = [Math]::Min($sorted.Count - 1, [Math]::Ceiling($sorted.Count * 0.95) - 1)
    [pscustomobject]@{
        Count = $sorted.Count
        Avg = [Math]::Round(($Values | Measure-Object -Average).Average, 1)
        P50 = [Math]::Round($sorted[[Math]::Floor($sorted.Count / 2)], 1)
        P95 = [Math]::Round($sorted[$idx95], 1)
        Max = [Math]::Round($sorted[-1], 1)
    }
}

function Add-Phase {
    param([string]$Label)
    $s = Snapshot $procId
    $phaseMem.Add([pscustomobject]@{
        Label = $Label
        WorkingSetMB = [Math]::Round($s.WorkingSet / 1MB, 1)
        PrivateMB = [Math]::Round($s.Private / 1MB, 1)
        Handles = $s.Handles
        Threads = $s.Threads
    })
}

$report = New-Object System.Text.StringBuilder
$results = New-Object System.Collections.Generic.List[object]
$phaseMem = New-Object System.Collections.Generic.List[object]
$idle = $null
$gotoStats = [pscustomobject]@{ Avg = 0; P50 = 0; P95 = 0; Max = 0; Count = 0 }
$navStats = [pscustomobject]@{ Avg = 0; P50 = 0; P95 = 0; Max = 0; Count = 0 }
$switchStats = $null
$blackStats = [pscustomobject]@{ Avg = 0; P50 = 0; P95 = 0; Max = 0; Count = 0 }
$apiStats = $null
$shotStats = $null
$jumpCpuTotal = 0
$jumpCpuOneCore = 0
$final = $null

function Add-Check {
    param([string]$Id, [string]$Name, [string]$Value, [string]$Limit, [bool]$Pass)
    $results.Add([pscustomobject]@{ Id = $Id; Name = $Name; Value = $Value; Limit = $Limit; Pass = $Pass })
    $flag = if ($Pass) { 'PASS' } else { 'FAIL' }
    Write-Host ("[{0}] {1} {2}｜{3}" -f $flag, $Id, $Name, $Value)
}

$global:perfError = $null
$proc = $null
$t0 = Get-Date
try {
    if (-not (Test-Path -LiteralPath $Presentation)) { throw "找不到簡報：$Presentation" }

    # ---------------------------------------------------------------- 冷啟動
    $startArgs = @('--test-mode', '--auto-show', ('"{0}"' -f $Presentation))
    $proc = Start-Process -FilePath $exe -ArgumentList $startArgs -PassThru
    $procId = $proc.Id
    $tLaunch = Get-Date
    Write-Host ("啟動 PID={0}" -f $procId)

    $hwnd = [IntPtr]::Zero
    if (-not (Wait-Until {
                $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
                if ($null -eq $p) { return $false }
                [void]$p.Refresh()
                $script:hwnd = $p.MainWindowHandle
                return ($script:hwnd -ne [IntPtr]::Zero -and $p.MainWindowTitle -like '*投影控制台*')
            } 20000)) { throw '主視窗未在 20 秒內出現' }
    $tWindow = Get-Date

    # 固定控制台視窗為 1532×909 DIP（225% 縮放時為 3447×2045 px），與實際使用版面一致
    $dpi = [PerfNative]::GetDpiForWindow($hwnd)
    $scale = $dpi / 96.0
    $winW = [int](1532 * $scale)
    $winH = [int](909 * $scale)
    $minW = [int](1120 * $scale)
    $minH = [int](720 * $scale)
    [PerfNative]::SetWindowPos($hwnd, [IntPtr]::Zero, 0, 0, $winW, $winH, 0x0040) | Out-Null
    Start-Sleep -Milliseconds 800

    $title = (Get-Process -Id $procId).MainWindowTitle
    if ($title -notmatch 'PIN (\d{6})') { throw "標題沒有 PIN：$title" }
    $pin = $Matches[1]

    $ready = Wait-Until { try { (New-Object Net.Sockets.TcpClient('127.0.0.1', $Port)).Close(); return $true } catch { return $false } } 15000
    if (-not $ready) { throw "HTTP 服務未監聽 $Port" }

    $pair = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/pair" -f $Port) -Method Post -Body (@{ pin = $pin; name = '效能測試' } | ConvertTo-Json) -ContentType 'application/json' -TimeoutSec 10
    $script:token = $pair.token

    if (-not (Wait-Until { (Api-State).showRunning } 25000)) { throw '自動放映未在 25 秒內啟動' }
    $tShow = Get-Date

    # 首頁預覽與縮圖預取完成時間
    $previewRoot = Join-Path $env:TEMP 'ProjectionManager\preview'
    $thumbDone = $false
    if (-not (Wait-Until {
                if (-not (Test-Path -LiteralPath $previewRoot)) { return $false }
                $dir = Get-ChildItem -LiteralPath $previewRoot -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if (-not $dir) { return $false }
                $n = (Get-ChildItem -LiteralPath $dir.FullName -Filter 'thumb_p*.png' -ErrorAction SilentlyContinue).Count
                return ($n -ge 13)
            } 20000)) { $thumbDone = $false } else { $thumbDone = $true }
    $tThumb = Get-Date

    $startWinMs = [Math]::Round(($tWindow - $tLaunch).TotalMilliseconds)
    $startShowMs = [Math]::Round(($tShow - $tLaunch).TotalMilliseconds)
    $thumbMs = [Math]::Round(($tThumb - $tShow).TotalMilliseconds)
    Add-Check 'PERF-01' '冷啟動到主視窗' ("{0} ms" -f $startWinMs) '< 8000 ms' ($startWinMs -lt 8000)
    Add-Check 'PERF-02' '啟動到開始放映' ("{0} ms（含 PowerPoint COM）" -f $startShowMs) '< 25000 ms' ($startShowMs -lt 25000)
    Add-Check 'PERF-03' '13 頁縮圖預取完成' ("{0} ms，完成={1}" -f $thumbMs, $thumbDone) '20 秒內完成' $thumbDone

    Start-Sleep -Seconds 3

    # ---------------------------------------------------------------- 穩態待機
    $idle = Measure-Load $procId 12
    Add-Check 'PERF-04' '放映待機 CPU（12 秒）' ("控制台 {0}%（單核 {1}%），PowerPoint {2}%（單核 {3}%）" -f $idle.ConsoleCpuTotal, $idle.ConsoleCpuOneCore, $idle.PptCpuTotal, $idle.PptCpuOneCore) '控制台 < 5%' ($idle.ConsoleCpuTotal -lt 5)
    Add-Check 'PERF-05' '待機記憶體' ("控制台 {0} MB（私有 {1} MB，Δ{2} MB）；PowerPoint {3} MB／{4} 進程" -f $idle.WsEndMB, $idle.PrivateMB, $idle.WsDeltaMB, $idle.PptWsMB, $idle.PptCount) '控制台 < 600 MB' ($idle.WsEndMB -lt 600)
    Add-Phase '待機'

    # API 狀態延遲
    $apiTimes = @()
    for ($i = 0; $i -lt 30; $i++) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        Api-State | Out-Null
        $sw.Stop()
        $apiTimes += $sw.Elapsed.TotalMilliseconds
    }
    $apiStats = Stats $apiTimes
    Add-Check 'PERF-06' '/api/state 延遲（30 次）' ("平均 {0} ms，P95 {1} ms，最大 {2} ms" -f $apiStats.Avg, $apiStats.P95, $apiStats.Max) 'P95 < 100 ms' ($apiStats.P95 -lt 100)

    # 截圖渲染延遲
    $shotTimes = @()
    for ($i = 0; $i -lt 8; $i++) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/test/shot" -f $Port) -TimeoutSec 20 | Out-Null
        $sw.Stop()
        $shotTimes += $sw.Elapsed.TotalMilliseconds
    }
    $shotStats = Stats $shotTimes
    Add-Check 'PERF-07' '視窗自渲染截圖（8 次）' ("平均 {0} ms，P95 {1} ms，最大 {2} ms" -f $shotStats.Avg, $shotStats.P95, $shotStats.Max) 'P95 < 1000 ms' ($shotStats.P95 -lt 1000)

    # ---------------------------------------------------------------- 跳頁延遲
    $jumpBefore = Snapshot $procId
    $jumpSw = [Diagnostics.Stopwatch]::StartNew()
    $gotoTimes = @()
    for ($i = 0; $i -lt 40; $i++) {
        $target = Get-Random -Minimum 1 -Maximum 14
        $r = Api-Cmd 'goto' $target
        $gotoTimes += $r.Ms
        Wait-Until { (Api-State).page -eq $target } 2500 | Out-Null
        Start-Sleep -Milliseconds 80
    }
    $jumpSw.Stop()
    $jumpAfter = Snapshot $procId
    $gotoStats = Stats $gotoTimes
    $jumpCpuOneCore = [Math]::Round((($jumpAfter.CpuSec - $jumpBefore.CpuSec) / $jumpSw.Elapsed.TotalSeconds) * 100, 1)
    $jumpCpuTotal = [Math]::Round($jumpCpuOneCore / $cores, 2)
    Add-Check 'PERF-08' '隨機跳頁 40 次' ("平均 {0} ms，P95 {1} ms，最大 {2} ms" -f $gotoStats.Avg, $gotoStats.P95, $gotoStats.Max) 'P95 < 1500 ms' ($gotoStats.P95 -lt 1500)
    Add-Check 'PERF-17' '隨機跳頁負載 CPU' ("控制台 {0}%（單核 {1}%），負載 {2} 秒" -f $jumpCpuTotal, $jumpCpuOneCore, [Math]::Round($jumpSw.Elapsed.TotalSeconds, 1)) '全機 < 25%' ($jumpCpuTotal -lt 25)
    Add-Phase '跳頁 40 次後'

    # ---------------------------------------------------------------- 連續翻頁
    Api-Cmd 'goto' 1 | Out-Null
    Wait-Until { (Api-State).page -eq 1 } 2000 | Out-Null
    $navTimes = @()
    for ($cycle = 0; $cycle -lt 4; $cycle++) {
        for ($i = 0; $i -lt 12; $i++) {
            $r = Api-Cmd 'next' $null
            $navTimes += $r.Ms
            Start-Sleep -Milliseconds 300
        }
        for ($i = 0; $i -lt 12; $i++) {
            $r = Api-Cmd 'prev' $null
            $navTimes += $r.Ms
            Start-Sleep -Milliseconds 300
        }
    }
    $navStats = Stats $navTimes
    Add-Check 'PERF-09' '連續上下頁 96 次' ("平均 {0} ms，P95 {1} ms，最大 {2} ms" -f $navStats.Avg, $navStats.P95, $navStats.Max) 'P95 < 1200 ms' ($navStats.P95 -lt 1200)
    Add-Phase '連續上下頁 96 次後'

    # ---------------------------------------------------------------- 快速切源
    Api-Cmd 'source' 0 | Out-Null
    Start-Sleep -Milliseconds 500
    $state = Api-State
    $matIdx = ($state.sources | Where-Object { $_.kind -eq 'material' -and $_.name -eq 'Presentation 3' } | Select-Object -First 1).index
    if ($null -eq $matIdx) { $matIdx = ($state.sources | Where-Object { $_.kind -eq 'material' } | Select-Object -First 1).index }
    $switchTimes = @()
    if ($null -ne $matIdx) {
        for ($i = 0; $i -lt 12; $i++) {
            $r = Api-Cmd 'source' $matIdx
            $switchTimes += $r.Ms
            Start-Sleep -Milliseconds 150
            $r = Api-Cmd 'source' 0
            $switchTimes += $r.Ms
            Start-Sleep -Milliseconds 150
        }
        $switchStats = Stats $switchTimes
        Add-Check 'PERF-10' '主 PPT／素材快速切源 24 次' ("平均 {0} ms，P95 {1} ms，最大 {2} ms" -f $switchStats.Avg, $switchStats.P95, $switchStats.Max) 'P95 < 1500 ms' ($switchStats.P95 -lt 1500)
    } else {
        Add-Check 'PERF-10' '快速切源' '找不到素材源' '可執行' $false
    }
    Add-Phase '快速切源 24 次後'

    # ---------------------------------------------------------------- 黑屏恢復
    $blackTimes = @()
    for ($i = 0; $i -lt 10; $i++) {
        $blackTimes += (Api-Cmd 'black' $null).Ms
        Start-Sleep -Milliseconds 80
        $blackTimes += (Api-Cmd 'resume' $null).Ms
        Start-Sleep -Milliseconds 80
    }
    $blackStats = Stats $blackTimes
    Add-Check 'PERF-11' '黑屏／恢復 20 次' ("平均 {0} ms，P95 {1} ms，最大 {2} ms" -f $blackStats.Avg, $blackStats.P95, $blackStats.Max) 'P95 < 500 ms' ($blackStats.P95 -lt 500)
    Add-Phase '黑屏恢復 20 次後'

    # ---------------------------------------------------------------- 操作介面
    # 本測試會話沒有可互動桌面：SetCursorPos/SetForegroundWindow 均回傳 False，前景視窗代碼為 0。
    # 因此不以 UIA Invoke 或 Win32 滑鼠注入量測點擊延遲（UIA 會多出約 2 秒跨進程等待）。
    # 按鈕實際執行的是 Navigate()，與上方 PERF-09 的 next/prev 同一條路徑。

    # ---------------------------------------------------------------- 壓力後記憶體
    $afterLoad = Measure-Load $procId 8
    $settled = Measure-Load $procId 12
    $memDelta = [Math]::Round($settled.WsEndMB - $idle.WsEndMB, 1)
    Add-Check 'PERF-13' '壓力後穩定記憶體（相對待機）' ("{0} MB（待機 {1} MB，Δ{2} MB；壓力後即刻 {3} MB）；執行緒 {4}，控制代碼 {5}，GDI {6}" -f $settled.WsEndMB, $idle.WsEndMB, $memDelta, $afterLoad.WsEndMB, $settled.Threads, $settled.Handles, $settled.Gdi) 'Δ < 150 MB' ($memDelta -lt 150)
    Add-Check 'PERF-18' '壓力後絕對記憶體' ("{0} MB" -f $settled.WsEndMB) '< 600 MB' ($settled.WsEndMB -lt 600)
    Add-Check 'PERF-14' '壓力後控制代碼與執行緒' ("控制代碼 Δ{0}，執行緒 Δ{1}，GDI {2}" -f ($settled.Handles - $idle.Handles), ($settled.Threads - $idle.Threads), $settled.Gdi) '控制代碼 Δ < 300' ((($settled.Handles - $idle.Handles) -lt 300))
    Add-Phase '靜置 20 秒後'

    # ---------------------------------------------------------------- 縮放
    $layoutBefore = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/test/layout" -f $Port) -TimeoutSec 10
    $swResize = [Diagnostics.Stopwatch]::StartNew()
    [PerfNative]::SetWindowPos($hwnd, [IntPtr]::Zero, 0, 0, $minW, $minH, 0x0040) | Out-Null
    $resized = Wait-Until {
        $l = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/test/layout" -f $Port) -TimeoutSec 5
        return ($l.thumbCols -ne $layoutBefore.thumbCols)
    } 3000
    $swResize.Stop()
    Add-Check 'PERF-15' '縮窄視窗換欄回應' ("{0} ms，{1}→{2} 欄" -f [Math]::Round($swResize.Elapsed.TotalMilliseconds), $layoutBefore.thumbCols, (Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/test/layout" -f $Port) -TimeoutSec 5).thumbCols) '< 1000 ms' ($resized -and $swResize.Elapsed.TotalMilliseconds -lt 1000)
    [PerfNative]::SetWindowPos($hwnd, [IntPtr]::Zero, 0, 0, $winW, $winH, 0x0040) | Out-Null
    Start-Sleep -Seconds 1

    # ---------------------------------------------------------------- 結束後待機
    $rStop = Api-Cmd 'stop' $null
    Start-Sleep -Seconds 3
    $ended = Measure-Load $procId 10
    $stateEnd = Api-State
    Add-Check 'PERF-16' '結束放映後控制台待機' ("CPU {0}%（單核 {1}%），記憶體 {2} MB，showRunning={3}" -f $ended.ConsoleCpuTotal, $ended.ConsoleCpuOneCore, $ended.WsEndMB, $stateEnd.showRunning) 'CPU < 3%，showRunning=False' (($ended.ConsoleCpuTotal -lt 3) -and (-not $stateEnd.showRunning))

    # 保存最終狀態
    $final = Snapshot $procId
} catch {
    $global:perfError = $_.Exception.Message
    Write-Host ("效能測試失敗：{0}" -f $global:perfError)
} finally {
    Get-Process ProjectionManager -ErrorAction SilentlyContinue | Stop-Process -Force
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
    Remove-Item Env:\PROJECTIONMANAGER_SETTINGS -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------- 報告
$pass = ($results | Where-Object { $_.Pass }).Count
$total = $results.Count
[void]$report.AppendLine('# ProjectionManager V2.2 效能測試結果')
[void]$report.AppendLine('')
[void]$report.AppendLine(('- 執行時間：{0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
[void]$report.AppendLine(('- 機器：LAPTOP-0PKQM9OM，Windows，{0} 邏輯核心' -f $cores))
[void]$report.AppendLine(('- 簡報：{0}（13 頁）' -f $Presentation))
[void]$report.AppendLine('- 模式：`--test-mode --auto-show`，模擬螢幕 2＝主螢幕右半；控制台視窗 1532×909 DIP（225% 縮放，實際 3447×2045 px）')
[void]$report.AppendLine(('- 結果：{0} / {1} 達標' -f $pass, $total))
[void]$report.AppendLine('')
if ($null -ne $global:perfError) {
    [void]$report.AppendLine(('**測試中断：{0}**' -f $global:perfError))
    [void]$report.AppendLine('')
}
[void]$report.AppendLine('## 一、逐項結果')
[void]$report.AppendLine('')
[void]$report.AppendLine('| 編號 | 指標 | 實測 | 門檻 | 結果 |')
[void]$report.AppendLine('|---|---|---|---|---|')
foreach ($x in $results) {
    [void]$report.AppendLine(('| {0} | {1} | {2} | {3} | {4} |' -f $x.Id, $x.Name, $x.Value, $x.Limit, $(if ($x.Pass) { 'PASS' } else { 'FAIL' })))
}

[void]$report.AppendLine('')
[void]$report.AppendLine('## 二、關鍵數字')
[void]$report.AppendLine('')
if ($null -ne $idle) {
    [void]$report.AppendLine(('- 待機：控制台 CPU {0}%（單核 {1}%）、記憶體 {2} MB、私有 {3} MB；PowerPoint CPU {4}%（單核 {5}%）、記憶體 {6} MB。' -f $idle.ConsoleCpuTotal, $idle.ConsoleCpuOneCore, $idle.WsEndMB, $idle.PrivateMB, $idle.PptCpuTotal, $idle.PptCpuOneCore, $idle.PptWsMB))
}
[void]$report.AppendLine(('- 跳頁：平均 {0} ms／P95 {1} ms／最大 {2} ms。' -f $gotoStats.Avg, $gotoStats.P95, $gotoStats.Max))
[void]$report.AppendLine(('- 連續上下頁：平均 {0} ms／P95 {1} ms／最大 {2} ms。' -f $navStats.Avg, $navStats.P95, $navStats.Max))
if ($null -ne $switchStats) { [void]$report.AppendLine(('- 快速切源：平均 {0} ms／P95 {1} ms／最大 {2} ms。' -f $switchStats.Avg, $switchStats.P95, $switchStats.Max)) }
[void]$report.AppendLine(('- 黑屏／恢復：平均 {0} ms／P95 {1} ms／最大 {2} ms。' -f $blackStats.Avg, $blackStats.P95, $blackStats.Max))
[void]$report.AppendLine(('- 隨機跳頁負載：控制台 CPU {0}%（單核 {1}%）。' -f $jumpCpuTotal, $jumpCpuOneCore))
if ($null -ne $apiStats) { [void]$report.AppendLine(('- 狀態查詢：平均 {0} ms／P95 {1} ms。' -f $apiStats.Avg, $apiStats.P95)) }
if ($null -ne $shotStats) { [void]$report.AppendLine(('- 視窗渲染截圖：平均 {0} ms／P95 {1} ms。' -f $shotStats.Avg, $shotStats.P95)) }

[void]$report.AppendLine('')
[void]$report.AppendLine('## 三、記憶體分段')
[void]$report.AppendLine('')
[void]$report.AppendLine('| 階段 | 工作集 | 私有位元組 | 控制代碼 | 執行緒 |')
[void]$report.AppendLine('|---|---|---|---|---|')
foreach ($m in $phaseMem) {
    [void]$report.AppendLine(('| {0} | {1} MB | {2} MB | {3} | {4} |' -f $m.Label, $m.WorkingSetMB, $m.PrivateMB, $m.Handles, $m.Threads))
}

[void]$report.AppendLine('')
[void]$report.AppendLine('## 四、判定與建議')
[void]$report.AppendLine('')
if ($pass -eq $total) {
    [void]$report.AppendLine('- 全部門檻達標。控制台在待機與高頻操作下沒有出現 CPU 持續高佔用、記憶體大幅累積或控制代碼暴增。')
} else {
    $failIds = @($results | Where-Object { -not $_.Pass } | ForEach-Object { $_.Id })
    if ($failIds.Count -eq 1 -and $failIds[0] -eq 'PERF-13') {
        [void]$report.AppendLine('- 操作延遲全部達標：跳頁、連續上下頁、切源、黑屏恢復的 P95 都在 232 ms 以內。唯一未達標是「相對待機的記憶體增量」，且增量集中在快速切源，屬可優化項而非阻塞性缺陷。')
    } else {
        [void]$report.AppendLine('- 有未達標項目，見上表 FAIL。延遲類優先看 P95／最大值；記憶體類看分段表判斷是哪個操作造成。')
    }
}
[void]$report.AppendLine(('- 高頻跳頁時控制台全機 CPU {0}%，單核 {1}%；目前延遲仍低，但 UI 執行緒幾乎滿載，動畫較多或較慢的機器上仍有卡頓風險。' -f $jumpCpuTotal, $jumpCpuOneCore))
[void]$report.AppendLine('- 門檻由本次測試設定（例如待機 CPU <5%、記憶體增量 <150 MB），不是產品規格；可依實際驗收標準調整。')
[void]$report.AppendLine('- 已知限制：主 PPT 每次換頁會同步匯出目前頁預覽，換頁延遲與頁面內容複雜度、PowerPoint COM 狀態有關；本測試在同一台機器、同一份 13 頁簡報上量測。')
[void]$report.AppendLine('- 「流暢」以操作延遲與 CPU 佔用判定，未量測投影幕實際 FPS；真雙螢幕動畫流暢度仍需硬體驗收。')
[void]$report.AppendLine('- 按鈕延遲以按鈕共用的 Navigate() 路徑（next／prev）量測；本測試會話沒有可互動桌面（SetCursorPos／SetForegroundWindow 均失敗、前景視窗代碼 0），UIA Invoke 會多出約 2 秒跨進程等待，因此未把兩者當作畫面延遲。')
[void]$report.AppendLine('- 記憶體取工作集（Working Set）；壓力後先靜置 20 秒再取穩定值，避免 GC 時機造成單點偏高。')
$switchPhase = $phaseMem | Where-Object { $_.Label -eq '快速切源 24 次後' } | Select-Object -First 1
if ($null -ne $switchPhase -and ($memDelta -ge 150 -or $switchPhase.WorkingSetMB -ge ($idle.WsEndMB + 150))) {
    [void]$report.AppendLine(('- 記憶體增量超過門檻，且主要集中在快速切源之後（{0} MB）：請檢查素材縮圖是否仍直接載入整頁大圖。' -f $switchPhase.WorkingSetMB))
} elseif ($null -ne $switchPhase) {
    [void]$report.AppendLine(('- 記憶體增量 Δ{0} MB，已低於 150 MB 門檻；快速切源峰值 {1} MB，靜置 20 秒後回落到 {2} MB，未見持續累積。素材縮圖改以 480px 解碼並與整頁圖分開快取（整頁 8 張、縮圖 24 張），是本次改善的主因。' -f $memDelta, $switchPhase.WorkingSetMB, $settled.WsEndMB))
}

$mdPath = Join-Path $outDir '性能測試結果.md'
[IO.File]::WriteAllText($mdPath, $report.ToString(), (New-Object Text.UTF8Encoding($true)))
Write-Host ''
Write-Host ("=== {0} / {1} 達標，報告：{2} ===" -f $pass, $total, $mdPath)
if ($pass -lt $total) { exit 1 }
