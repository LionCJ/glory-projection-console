using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;

namespace ProjectionManager.Services;

/// <summary>一支已綁定的手機。Token 為配對時發出的長隨機字串，撤銷後立即失效。</summary>
public sealed class MobileDevice
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Role { get; set; } = "operator";   // operator | speaker
    public string Token { get; set; } = string.Empty;
    public string Created { get; set; } = string.Empty;
    public string LastSeen { get; set; } = string.Empty;
}

public sealed class MobileSourceInfo
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Kind { get; set; } = string.Empty;
    public string Group { get; set; } = string.Empty;
    public string GroupKey { get; set; } = string.Empty;
    public bool Live { get; set; }
    public bool Projected { get; set; }
    public int Page { get; set; }
    public int Total { get; set; }
}

public sealed class MobileState
{
    public bool Ok { get; set; } = true;
    public string Role { get; set; } = "operator";
    public string DeviceName { get; set; } = string.Empty;
    public bool ShowRunning { get; set; }
    public bool HasMain { get; set; }
    public bool Black { get; set; }
    public string Mode { get; set; } = "operator";
    public string SourceName { get; set; } = string.Empty;
    public string SourceKind { get; set; } = string.Empty;
    public bool Projected { get; set; }
    public int Page { get; set; }
    public int Total { get; set; }
    public string Status { get; set; } = string.Empty;
    public List<MobileSourceInfo> Sources { get; set; } = new();
}

public sealed class MobileCommandResult
{
    public bool Ok { get; set; }
    public string Message { get; set; } = string.Empty;

    public static MobileCommandResult Success(string message = "") => new() { Ok = true, Message = message };
    public static MobileCommandResult Fail(string message) => new() { Ok = false, Message = message };
}

/// <summary>
/// 手機遙控器服務：只綁區網、只認配對後的 token。
/// 用 TcpListener 自建最小 HTTP，避免 HttpListener 需要系統管理員註冊 URL 權限。
/// 指令一律回到 UI 執行緒，轉呼叫控制台原本的按鈕邏輯，行為與翻頁筆一致。
/// </summary>
public sealed class MobileControlServer : IDisposable
{
    private readonly Func<string> _pinProvider;
    private readonly Func<List<MobileDevice>> _devicesProvider;
    private readonly Action _saveDevices;
    private readonly Func<MobileState> _stateProvider;
    private readonly Func<MobileDevice, string, string?, MobileCommandResult> _executor;
    private readonly IPAddress _bindAddress;
    private readonly Action<string, string>? _testRoleSetter;
    private readonly Func<object>? _testLayoutProvider;
    private readonly Func<object>? _testShotProvider;
    private readonly Func<string, object>? _testActionProvider;

    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, long> _lastCommand =
        new(StringComparer.OrdinalIgnoreCase);

    private TcpListener? _listener;
    private CancellationTokenSource? _cts;

    public int Port { get; }
    public bool IsRunning => _listener is not null;
    public string? LastError { get; private set; }

    public MobileControlServer(
        int port,
        Func<string> pinProvider,
        Func<List<MobileDevice>> devicesProvider,
        Action saveDevices,
        Func<MobileState> stateProvider,
        Func<MobileDevice, string, string?, MobileCommandResult> executor,
        IPAddress? bindAddress = null,
        Action<string, string>? testRoleSetter = null,
        Func<object>? testLayoutProvider = null,
        Func<object>? testShotProvider = null,
        Func<string, object>? testActionProvider = null)
    {
        Port = port;
        _pinProvider = pinProvider;
        _devicesProvider = devicesProvider;
        _saveDevices = saveDevices;
        _stateProvider = stateProvider;
        _executor = executor;
        _bindAddress = bindAddress ?? IPAddress.Any;
        _testRoleSetter = testRoleSetter;
        _testLayoutProvider = testLayoutProvider;
        _testShotProvider = testShotProvider;
        _testActionProvider = testActionProvider;
    }

    public bool Start()
    {
        if (_listener is not null)
        {
            return true;
        }

        try
        {
            _listener = new TcpListener(_bindAddress, Port);
            _listener.Start();
            _cts = new CancellationTokenSource();
            _ = Task.Run(() => AcceptLoopAsync(_cts.Token));
            LastError = null;
            return true;
        }
        catch (Exception ex)
        {
            LastError = ex.Message;
            _listener = null;
            return false;
        }
    }

    public void Stop()
    {
        try
        {
            _cts?.Cancel();
            _listener?.Stop();
        }
        catch
        {
            // 關閉失敗不影響程式退出
        }
        finally
        {
            _listener = null;
            _cts = null;
        }
    }

    public void Dispose() => Stop();

    private async Task AcceptLoopAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            TcpClient client;
            try
            {
                client = await _listener!.AcceptTcpClientAsync(token);
            }
            catch
            {
                if (token.IsCancellationRequested)
                {
                    return;
                }

                continue;
            }

            _ = Task.Run(() => HandleClient(client), token);
        }
    }

    private void HandleClient(TcpClient client)
    {
        using (client)
        {
            try
            {
                client.ReceiveTimeout = 6000;
                client.SendTimeout = 6000;

                using var stream = client.GetStream();
                var request = ReadRequest(stream);
                if (request is null)
                {
                    return;
                }

                Route(stream, request.Value.Method, request.Value.Path, request.Value.Query, request.Value.Body);
            }
            catch
            {
                // 手機斷線、逾時、格式錯誤都直接丟棄，不影響主程式
            }
        }
    }

    private static (string Method, string Path, string Query, string Body)? ReadRequest(NetworkStream stream)
    {
        var buffer = new byte[8192];
        var data = new List<byte>(4096);
        int headerEnd = -1;

        while (headerEnd < 0)
        {
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read <= 0)
            {
                return null;
            }

            for (int i = 0; i < read; i++)
            {
                data.Add(buffer[i]);
            }

            headerEnd = FindHeaderEnd(data);
            if (data.Count > 65536)
            {
                return null;
            }
        }

        var headBytes = data.GetRange(0, headerEnd);
        var bodyBytes = data.GetRange(headerEnd + 4, data.Count - headerEnd - 4);
        var headText = Encoding.UTF8.GetString(headBytes.ToArray());
        var lines = headText.Split("\r\n", StringSplitOptions.None);
        var parts = lines[0].Split(' ');
        if (parts.Length < 2)
        {
            return null;
        }

        int contentLength = 0;
        foreach (var line in lines)
        {
            int colon = line.IndexOf(':');
            if (colon <= 0)
            {
                continue;
            }

            if (line[..colon].Equals("Content-Length", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(line[(colon + 1)..].Trim(), out int length))
            {
                contentLength = Math.Clamp(length, 0, 65536);
            }
        }

        while (bodyBytes.Count < contentLength)
        {
            int read = stream.Read(buffer, 0, Math.Min(buffer.Length, contentLength - bodyBytes.Count));
            if (read <= 0)
            {
                break;
            }

            for (int i = 0; i < read; i++)
            {
                bodyBytes.Add(buffer[i]);
            }
        }

        if (bodyBytes.Count > contentLength)
        {
            bodyBytes.RemoveRange(contentLength, bodyBytes.Count - contentLength);
        }

        var target = parts[1];
        int q = target.IndexOf('?');
        string path = q < 0 ? target : target[..q];
        string query = q < 0 ? string.Empty : target[(q + 1)..];

        return (parts[0].ToUpperInvariant(), path, query, Encoding.UTF8.GetString(bodyBytes.ToArray()));
    }

    private static int FindHeaderEnd(List<byte> data)
    {
        for (int i = 0; i + 3 < data.Count; i++)
        {
            if (data[i] == (byte)'\r' && data[i + 1] == (byte)'\n'
                && data[i + 2] == (byte)'\r' && data[i + 3] == (byte)'\n')
            {
                return i;
            }
        }

        return -1;
    }

    private void Route(NetworkStream stream, string method, string path, string query, string body)
    {
        if (method == "GET" && (path == "/" || path == "/index.html"))
        {
            Respond(stream, 200, "text/html", MobilePage.Html);
            return;
        }

        if (method == "GET" && path == "/api/state")
        {
            HandleState(stream, query);
            return;
        }

        if (method == "POST" && path == "/api/pair")
        {
            HandlePair(stream, body);
            return;
        }

        // 只在測試模式（且僅綁 127.0.0.1）開放：讀取介面實際尺寸，讓版面測試可量化。
        if (method == "GET" && path == "/api/test/layout" && _testLayoutProvider is not null)
        {
            try
            {
                RespondJson(stream, 200, _testLayoutProvider());
            }
            catch
            {
                RespondJson(stream, 500, new { ok = false, error = "layout_failed" });
            }

            return;
        }

        // 只在測試模式：由 WPF 自行渲染視窗截圖，桌面擷取失效時仍可取得畫面。
        if (method == "GET" && path == "/api/test/shot" && _testShotProvider is not null)
        {
            try
            {
                RespondJson(stream, 200, _testShotProvider());
            }
            catch
            {
                RespondJson(stream, 500, new { ok = false, error = "shot_failed" });
            }

            return;
        }

        // 只在測試模式：驅動「先加入再推送」等 UI 流程，不必操作複雜的彈窗。
        if (method == "POST" && path == "/api/test/action" && _testActionProvider is not null)
        {
            try
            {
                string action;
                using (var doc = JsonDocument.Parse(body))
                {
                    action = doc.RootElement.TryGetProperty("action", out var a)
                        ? a.GetString() ?? string.Empty
                        : string.Empty;
                }

                RespondJson(stream, 200, _testActionProvider(action));
            }
            catch (Exception ex)
            {
                RespondJson(stream, 500, new { ok = false, error = ex.Message });
            }

            return;
        }

        // 只在測試模式（且僅綁 127.0.0.1）開放：讓自動化切換角色，不必操作視窗。
        if (method == "POST" && path == "/api/test/role" && _testRoleSetter is not null)
        {
            try
            {
                using var doc = JsonDocument.Parse(body);
                string id = doc.RootElement.TryGetProperty("id", out var i) ? i.GetString() ?? string.Empty : string.Empty;
                string role = doc.RootElement.TryGetProperty("role", out var r) ? r.GetString() ?? string.Empty : string.Empty;
                if (string.IsNullOrWhiteSpace(id) || (role != "operator" && role != "speaker"))
                {
                    RespondJson(stream, 400, new { ok = false, error = "bad_role" });
                    return;
                }

                _testRoleSetter(id, role);
                RespondJson(stream, 200, new { ok = true, role });
            }
            catch
            {
                RespondJson(stream, 400, new { ok = false, error = "bad_json" });
            }

            return;
        }

        if (method == "POST" && path == "/api/cmd")
        {
            HandleCommand(stream, body);
            return;
        }

        if (method == "GET" && path == "/favicon.ico")
        {
            Respond(stream, 204, "text/plain", string.Empty);
            return;
        }

        RespondJson(stream, 404, new { ok = false, error = "not_found" });
    }

    private void HandleState(NetworkStream stream, string query)
    {
        var device = FindDevice(GetQueryValue(query, "token"));
        if (device is null)
        {
            RespondJson(stream, 401, new { ok = false, error = "unpaired" });
            return;
        }

        device.LastSeen = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var state = _stateProvider();
        state.Ok = true;
        state.Role = device.Role;
        state.DeviceName = device.Name;

        if (device.Role == "speaker")
        {
            state.Sources.Clear();
        }

        RespondJson(stream, 200, state);
    }

    private void HandlePair(NetworkStream stream, string body)
    {
        string pin = string.Empty;
        string name = string.Empty;

        try
        {
            using var doc = JsonDocument.Parse(body);
            pin = doc.RootElement.TryGetProperty("pin", out var p) ? p.GetString() ?? string.Empty : string.Empty;
            name = doc.RootElement.TryGetProperty("name", out var n) ? n.GetString() ?? string.Empty : string.Empty;
        }
        catch
        {
            RespondJson(stream, 400, new { ok = false, error = "bad_json" });
            return;
        }

        if (!pin.Trim().Equals(_pinProvider(), StringComparison.Ordinal))
        {
            RespondJson(stream, 403, new { ok = false, error = "pin" });
            return;
        }

        var devices = _devicesProvider();
        if (devices.Count >= 5)
        {
            RespondJson(stream, 409, new { ok = false, error = "full" });
            return;
        }

        name = string.IsNullOrWhiteSpace(name) ? "手機" : name.Trim();
        if (name.Length > 24)
        {
            name = name[..24];
        }

        var device = new MobileDevice
        {
            Id = Guid.NewGuid().ToString("N"),
            Name = name,
            Role = "operator",
            Token = Guid.NewGuid().ToString("N") + Guid.NewGuid().ToString("N"),
            Created = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            LastSeen = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
        };

        devices.Add(device);
        _saveDevices();

        RespondJson(stream, 200, new
        {
            ok = true,
            token = device.Token,
            id = device.Id,
            role = device.Role,
            name = device.Name,
        });
    }

    private void HandleCommand(NetworkStream stream, string body)
    {
        string token = string.Empty;
        string cmd = string.Empty;
        string? arg = null;

        try
        {
            using var doc = JsonDocument.Parse(body);
            token = doc.RootElement.TryGetProperty("token", out var t) ? t.GetString() ?? string.Empty : string.Empty;
            cmd = doc.RootElement.TryGetProperty("cmd", out var c) ? c.GetString() ?? string.Empty : string.Empty;
            if (doc.RootElement.TryGetProperty("arg", out var a))
            {
                arg = a.ValueKind == JsonValueKind.Number ? a.GetRawText() : a.GetString();
            }
        }
        catch
        {
            RespondJson(stream, 400, new { ok = false, message = "指令格式錯誤" });
            return;
        }

        var device = FindDevice(token);
        if (device is null)
        {
            RespondJson(stream, 401, new { ok = false, error = "unpaired", message = "尚未配對" });
            return;
        }

        cmd = cmd.Trim().ToLowerInvariant();
        device.LastSeen = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        // 講員端只給上下頁；其餘指令一律拒絕
        bool speakerAllowed = cmd is "next" or "prev";
        if (device.Role == "speaker" && !speakerAllowed)
        {
            RespondJson(stream, 403, new { ok = false, message = "講員端只允許上下頁" });
            return;
        }

        if (cmd is "next" or "prev")
        {
            long now = Environment.TickCount64;
            string key = device.Id + ":" + cmd;
            if (_lastCommand.TryGetValue(key, out long last) && now - last < 260)
            {
                RespondJson(stream, 200, new { ok = true, message = "已略過重複點擊" });
                return;
            }

            _lastCommand[key] = now;
        }

        var result = _executor(device, cmd, arg);
        RespondJson(stream, 200, new { ok = result.Ok, message = result.Message });
    }

    private MobileDevice? FindDevice(string token)
    {
        if (string.IsNullOrWhiteSpace(token))
        {
            return null;
        }

        return _devicesProvider().FirstOrDefault(d => d.Token.Equals(token, StringComparison.Ordinal));
    }

    private static string GetQueryValue(string query, string key)
    {
        foreach (var pair in query.Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            int eq = pair.IndexOf('=');
            string k = eq < 0 ? pair : pair[..eq];
            if (!k.Equals(key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string v = eq < 0 ? string.Empty : pair[(eq + 1)..];
            return Uri.UnescapeDataString(v.Replace('+', ' '));
        }

        return string.Empty;
    }

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    private static void RespondJson(NetworkStream stream, int code, object payload)
    {
        Respond(stream, code, "application/json", JsonSerializer.Serialize(payload, JsonOptions));
    }

    private static void Respond(NetworkStream stream, int code, string contentType, string body)
    {
        var bytes = Encoding.UTF8.GetBytes(body);
        string reason = code switch
        {
            200 => "OK",
            204 => "No Content",
            400 => "Bad Request",
            401 => "Unauthorized",
            403 => "Forbidden",
            404 => "Not Found",
            409 => "Conflict",
            _ => "OK",
        };

        string head = $"HTTP/1.1 {code} {reason}\r\n"
                      + $"Content-Type: {contentType}; charset=utf-8\r\n"
                      + $"Content-Length: {bytes.Length}\r\n"
                      + "Cache-Control: no-store\r\n"
                      + "Connection: close\r\n\r\n";

        var headBytes = Encoding.UTF8.GetBytes(head);
        stream.Write(headBytes, 0, headBytes.Length);
        if (bytes.Length > 0)
        {
            stream.Write(bytes, 0, bytes.Length);
        }

        stream.Flush();
    }
}

/// <summary>手機端單頁介面：內嵌在程式裡，不依賴網路上的任何資源。</summary>
internal static class MobilePage
{
    public const string Html = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#171B21">
<title>投影控制</title>
<style>
:root{--bg:#14181E;--panel:#232932;--line:#3A424D;--text:#F5F7FA;--dim:#C7D0DB;--faint:#96A1B0;--teal:#47D1C4;--amber:#F0B54A;--danger:#F4695F}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font-family:system-ui,-apple-system,"PingFang TC","Microsoft JhengHei",sans-serif;-webkit-user-select:none;user-select:none}
.wrap{max-width:520px;margin:0 auto;padding:16px 14px 28px}
h1{font-size:18px;margin:0 0 2px}
.sub{color:var(--faint);font-size:12px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px;margin-top:14px}
.state{display:flex;justify-content:space-between;align-items:center;gap:10px}
.src{font-size:15px;font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.page{font-size:28px;font-weight:700;color:var(--teal)}
.page span{font-size:14px;color:var(--dim);font-weight:400}
.badge{display:inline-block;padding:3px 8px;border-radius:6px;font-size:12px;border:1px solid var(--teal);color:var(--teal)}
.badge.black{border-color:var(--amber);color:var(--amber)}
.nav{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:14px}
button{font-family:inherit;font-size:17px;font-weight:600;color:var(--text);background:#2A313A;border:1px solid #4A5462;border-radius:12px;padding:20px 10px;cursor:pointer;touch-action:manipulation}
button:active{background:#35404C}
button.primary{border-color:var(--teal);color:var(--teal)}
button.danger{border-color:var(--danger);color:var(--danger)}
button:disabled{opacity:.35}
.row{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px}
.row button{padding:14px 8px;font-size:15px}
.action{display:flex;gap:10px;margin-top:10px}
.action button{flex:1;min-width:0;padding:14px 8px;font-size:15px}
.goto{display:flex;gap:8px;margin-top:12px}
.goto input{flex:1;min-width:0;background:#1B2027;border:1px solid var(--line);border-radius:10px;color:var(--text);font-size:17px;padding:12px}
.goto button{flex:0 0 88px;padding:0}
label{display:block;color:var(--dim);font-size:13px;margin:0 0 6px}
input[type=text],input[type=password]{width:100%;background:#1B2027;border:1px solid var(--line);border-radius:10px;color:var(--text);font-size:17px;padding:13px}
.hint{color:var(--faint);font-size:12px;line-height:1.6;margin-top:10px}
.pair button{width:100%;margin-top:12px;padding:16px;border-color:var(--teal);color:var(--teal)}
.srcs{margin-top:14px}
.src-search{display:flex;gap:8px;margin-bottom:10px}
.src-search input{flex:1;min-width:0;background:#1B2027;border:1px solid var(--line);border-radius:10px;color:var(--text);font-size:15px;padding:11px 12px}
.src-search button{flex:0 0 64px;padding:0;font-size:14px}
.src-count{color:var(--faint);font-size:12px;font-weight:400}
.srcs .group{display:flex;align-items:center;gap:8px;width:100%;text-align:left;padding:12px 13px;margin-top:8px;font-size:14px;font-weight:600;background:#20262E;border-color:#3E4854;color:var(--dim)}
.srcs .group .caret{color:var(--teal);font-size:12px;width:14px;flex:0 0 14px}
.srcs .group .gname{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.srcs .group .gcount{flex:0 0 auto;color:var(--faint);font-size:12px;font-weight:400}
.srcs .item{width:100%;text-align:left;padding:13px 13px 13px 24px;margin-top:8px;font-size:14px;font-weight:500}
.srcs .item.live{border-color:var(--teal);color:var(--teal)}
.srcs .item.projected{border-color:var(--amber);color:var(--amber)}
.msg{min-height:20px;color:var(--amber);font-size:13px;margin-top:10px}
.modal{position:fixed;inset:0;background:rgba(0,0,0,.68);display:flex;align-items:center;justify-content:center;padding:18px;z-index:20}
.modal-card{width:100%;max-width:420px;background:var(--panel);border:1px solid var(--teal);border-radius:14px;padding:18px}
.modal-card .src{white-space:normal}
.modal-card input{width:100%;background:#1B2027;border:1px solid var(--line);border-radius:10px;color:var(--text);font-size:18px;padding:12px;margin-top:10px}
.modal-card .row{margin-top:14px}
.modal-card .row button{padding:14px 8px}
.hidden{display:none}
@media (prefers-reduced-motion:no-preference){button{transition:background .12s}}
</style>
</head>
<body>
<div class="wrap">
  <h1>投影控制</h1>
  <div class="sub" id="conn">連線中…</div>

  <div id="pairView" class="card pair hidden">
    <label>配對碼（看控制台）</label>
    <input id="pin" type="text" inputmode="numeric" autocomplete="one-time-code" placeholder="6 位數字">
    <label style="margin-top:12px">這支手機叫什麼</label>
    <input id="name" type="text" placeholder="例：講員 iPhone">
    <button onclick="pair()">配對並連線</button>
    <div class="msg" id="pairMsg"></div>
    <div class="hint">手機要與電腦在同一個 Wi-Fi。配對成功後會記住這支手機，下次不必再輸配對碼。</div>
  </div>

  <div id="mainView" class="hidden">
    <div class="card">
      <div class="state">
        <div style="min-width:0">
          <div class="src" id="src">—</div>
          <div class="sub" id="kind" style="margin-top:4px">—</div>
        </div>
        <div style="text-align:right">
          <div class="page" id="page">—</div>
          <div class="sub" id="status">—</div>
        </div>
      </div>
    </div>

    <div class="nav">
      <button id="prevBtn" onclick="cmd('prev')">◀ 上一頁</button>
      <button id="nextBtn" class="primary" onclick="cmd('next')">下一頁 ▶</button>
    </div>

    <div class="action" id="actionRow">
      <button id="startShowBtn" class="primary" onclick="startMainShow()">開始主 PPT 放映</button>
      <button id="pushBtn" class="primary" onclick="pushProjection()">推送投影</button>
    </div>

    <div class="goto">
      <input id="gotoPage" type="number" min="1" inputmode="numeric" placeholder="跳到第幾頁">
      <button onclick="goTo()">跳頁</button>
    </div>

    <div class="row">
      <button onclick="cmd('black')">黑屏</button>
      <button onclick="cmd('resume')">恢復畫面</button>
    </div>
    <div class="row">
      <button class="danger" onclick="stopShow()">結束放映</button>
      <button onclick="refresh()">重新整理</button>
    </div>

    <div class="srcs card" id="srcsCard">
      <div class="src-search">
        <input id="sourceSearch" type="text" autocomplete="off" enterkeyhint="search" aria-label="搜尋投影源" placeholder="搜尋投影源名稱或分組" oninput="setSourceQuery(this.value)">
        <button type="button" onclick="clearSourceSearch()">清除</button>
      </div>
      <div class="sub">切換投影源 <span class="src-count" id="sourceSearchCount"></span></div>
      <div id="srcs"></div>
    </div>

    <div class="modal hidden" id="switchModal">
      <div class="modal-card">
        <div class="src" id="switchTitle">切換投影源</div>
        <div class="sub" id="switchMeta" style="margin-top:6px">—</div>
        <div id="switchPageWrap">
          <label style="margin-top:14px">選擇頁數</label>
          <input id="switchPage" type="number" min="1" inputmode="numeric" value="1">
        </div>
        <div class="sub" id="switchHint" style="margin-top:10px"></div>
        <div class="row">
          <button onclick="closeSwitch()">取消</button>
          <button class="primary" onclick="confirmSwitch()">確認切換</button>
        </div>
      </div>
    </div>

    <div class="msg" id="msg"></div>
    <div class="hint" id="roleHint"></div>
  </div>
</div>

<script>
const KEY='pm_mobile_token';
let token=localStorage.getItem(KEY)||'';
const qs=new URLSearchParams(location.search);
if(qs.get('token')){token=qs.get('token');localStorage.setItem(KEY,token)}
let timer=null;
let role='operator';

function $(id){return document.getElementById(id)}

async function api(path,body){
  const opt=body?{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)}:{};
  const res=await fetch(path,opt);
  let data={};
  try{data=await res.json()}catch(e){}
  return {status:res.status,data};
}

function show(unpaired){
  $('pairView').classList.toggle('hidden',!unpaired);
  $('mainView').classList.toggle('hidden',unpaired);
  $('conn').textContent=unpaired?'尚未配對':'已連線';
}

if(qs.get('pin')){$('pin').value=qs.get('pin')}

async function pair(){
  const pin=$('pin').value.trim();
  const name=$('name').value.trim();
  if(!pin){$('pairMsg').textContent='請輸入配對碼';return}
  const r=await api('/api/pair',{pin,name});
  if(r.data && r.data.ok){
    token=r.data.token; role=r.data.role||'operator';
    localStorage.setItem(KEY,token);
    $('pairMsg').textContent='';
    await refresh();
  }else{
    $('pairMsg').textContent=r.data.error==='pin'?'配對碼不正確':'配對失敗，請重試';
  }
}

async function refresh(){
  if(!token){show(true);return}
  const r=await api('/api/state?token='+encodeURIComponent(token));
  if(r.status===401){token='';localStorage.removeItem(KEY);show(true);return}
  if(!r.data || !r.data.ok){$('conn').textContent='連線中斷，重試中…';return}
  show(false);
  const s=r.data;
  role=s.role||'operator';
  $('conn').textContent='已連線 · '+(s.deviceName||'手機');
  $('src').textContent=s.sourceName||'（尚無投影源）';
  $('kind').textContent=kindName(s.sourceKind);
  $('page').innerHTML=(s.page>0?s.page:'—')+(s.total>0?'<span> / '+s.total+'</span>':'');
  $('status').textContent=s.black?'黑屏':(s.showRunning?'放映中':(s.status||'未開始放映'));
  $('status').className='sub';
  updateProjectionButtons(s);
  $('roleHint').textContent=role==='speaker'?'講員模式：只開放上下頁。':'操作員模式：可使用全部控制。';
  $('srcsCard').classList.toggle('hidden',role==='speaker'||!s.sources||!s.sources.length);
  if(role!=='speaker'&&s.sources){
    renderSourceGroups(s.sources);
  }
}

let lastSources=[];
const expandedGroups=new Set();
let sourceQuery='';

function setSourceQuery(value){
  sourceQuery=(value||'').trim().toLowerCase();
  renderSourceGroups(lastSources);
}

function clearSourceSearch(){
  sourceQuery='';
  const input=$('sourceSearch');
  if(input)input.value='';
  renderSourceGroups(lastSources);
}

function updateProjectionButtons(s){
  const start=$('startShowBtn');
  const push=$('pushBtn');
  const startHidden=!!s.showRunning;
  start.classList.toggle('hidden',startHidden);
  start.disabled=!s.hasMain;
  start.textContent=s.hasMain?'開始主 PPT 放映':'尚未加入主 PPT';
  start.title=s.hasMain?'':"請先在電腦端按「加入主 PPT」";

  const canPush=!!s.sourceKind&&s.sourceKind!=='main';
  push.textContent=s.projected?'斷開投影':'推送投影';
  push.className=s.projected?'danger':'primary';
  push.classList.toggle('hidden',!canPush);

  $('actionRow').classList.toggle('hidden',startHidden&&!canPush);
}

function startMainShow(){
  if(!confirm('確定開始主 PPT 放映？'))return;
  cmd('start');
}

function pushProjection(){
  if(!confirm($('pushBtn').textContent+'？'))return;
  cmd('push');
}
function renderSourceGroups(items){
  lastSources=items||[];
  const box=$('srcs');
  box.innerHTML='';
  if(!lastSources.length)return;

  const groups=[];
  const byKey={};
  lastSources.forEach(function(item){
    const key=item.groupKey||item.group||item.kind||'other';
    const groupName=item.group||kindName(item.kind)||'其他';
    if(sourceQuery){
      const haystack=(String(item.name||'')+' '+groupName+' '+kindName(item.kind)).toLowerCase();
      if(haystack.indexOf(sourceQuery)<0)return;
    }
    let group=byKey[key];
    if(!group){
      group={key:key,name:groupName,items:[]};
      byKey[key]=group;
      groups.push(group);
    }
    group.items.push(item);
  });

  const total=groups.reduce(function(n,group){return n+group.items.length},0);
  $('sourceSearchCount').textContent=sourceQuery?'找到 '+total+' 項':'';
  if(!groups.length){
    const empty=document.createElement('div');
    empty.className='sub';
    empty.style.marginTop='10px';
    empty.textContent='找不到符合「'+(sourceQuery||'')+'」的投影源';
    box.appendChild(empty);
    return;
  }

  groups.forEach(function(group){
    const expanded=sourceQuery?true:expandedGroups.has(group.key);
    const header=document.createElement('button');
    header.type='button';
    header.className='group';
    header.setAttribute('aria-expanded',expanded?'true':'false');
    const caret=document.createElement('span');
    caret.className='caret';
    caret.textContent=expanded?'▾':'▸';
    const name=document.createElement('span');
    name.className='gname';
    name.textContent=group.name;
    const count=document.createElement('span');
    count.className='gcount';
    count.textContent=group.items.length+' 項';
    header.appendChild(caret);header.appendChild(name);header.appendChild(count);
    header.onclick=function(){
      if(sourceQuery)return;
      if(expandedGroups.has(group.key))expandedGroups.delete(group.key);
      else expandedGroups.add(group.key);
      renderSourceGroups(lastSources);
    };
    box.appendChild(header);

    if(!expanded)return;
    group.items.forEach(function(item){
      const b=document.createElement('button');
      b.type='button';
      b.className='item'+(item.live?' live':'')+(item.projected?' projected':'');
      const pages=item.total>0?'　'+(item.page||1)+'/'+item.total+' 頁':'';
      b.textContent=item.name+'　'+kindName(item.kind)+pages+(item.projected?'　投影中':'');
      b.onclick=function(){openSwitch(item)};
      box.appendChild(b);
    });
  });
}

let pendingSource=null;
function openSwitch(item){
  pendingSource=item;
  $('switchTitle').textContent='切換到「'+item.name+'」';
  $('switchMeta').textContent=kindName(item.kind)+(item.live?'　（目前投影源）':'');
  const hasPages=item.total>0;
  $('switchPageWrap').classList.toggle('hidden',!hasPages);
  $('switchPage').value=hasPages?String(item.page>0?item.page:1):'';
  $('switchPage').max=hasPages?String(item.total):'';
  $('switchHint').textContent=hasPages
    ? '可先選起始頁；經文／網頁／Zoom 未推送前不會上螢幕 2。'
    : '此來源沒有頁碼；切換後仍需按「推送投影」才會上螢幕 2。';
  $('switchModal').classList.remove('hidden');
}
function closeSwitch(){
  pendingSource=null;
  $('switchModal').classList.add('hidden');
}
function confirmSwitch(){
  if(!pendingSource)return;
  let page=1;
  if(pendingSource.total>0){
    page=parseInt($('switchPage').value,10);
    if(!page||page<1||page>pendingSource.total){
      $('switchHint').textContent='請輸入 1 到 '+pendingSource.total+' 的頁碼。';
      return;
    }
  }
  const item=pendingSource;
  const label=item.name+'（'+kindName(item.kind)+'）'+(item.total>0?' 第 '+page+' 頁':'');
  if(!confirm('再次確認：確定切換投影源到\n'+label+'？')){
    return;
  }
  closeSwitch();
  cmd('source',String(item.index)+'|'+String(page));
}

function kindName(k){
  return k==='main'?'主 PPT':k==='material'?'素材':k==='scripture'?'經文':k==='web'?'網頁':k==='zoom'?'Zoom':'';
}

async function cmd(name,arg){
  if(!token){show(true);return}
  const r=await api('/api/cmd',{token,cmd:name,arg:arg});
  if(r.status===401){token='';localStorage.removeItem(KEY);show(true);return}
  $('msg').textContent=(r.data&&r.data.message)?r.data.message:'';
  await refresh();
}

function goTo(){
  const n=parseInt($('gotoPage').value,10);
  if(!n||n<1){$('msg').textContent='請輸入頁碼';return}
  cmd('goto',String(n));
  $('gotoPage').value='';
}

function stopShow(){
  if(confirm('確定結束放映？講員將無法繼續翻頁。')){cmd('stop')}
}

show(!!token?false:true);
if(token){refresh()}
timer=setInterval(function(){if(token&&(!document.hidden)){refresh()}},900);
document.addEventListener('visibilitychange',function(){if(!document.hidden&&token){refresh()}});
</script>
</body>
</html>
""";
}
