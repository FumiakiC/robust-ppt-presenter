# ==============================================================================
#  対話型プレゼンテーション実行スクリプト V7.3 (Network Error Fix)
#  （クライアント切断によるコンソール強制終了バグ修正版）
# ==============================================================================

param(
    [string]$TargetFolderPath = $(if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }),
    [string]$FinishFolderName = "finish",
    [int]$WebPort = 8090
)

# エラー発生時に停止する設定ですが、送信エラーは個別にcatchして無視します
$ErrorActionPreference = 'Stop'

# アセンブリの読み込み
Add-Type -AssemblyName System.Web

# ============================================================================== 
# コンソールウィンドウ制御（誤操作防止）
# ============================================================================== 
#region ConsoleWindow API
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class ConsoleWindow {
    [DllImport("kernel32.dll", ExactSpelling = true)]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

    [DllImport("user32.dll")]
    public static extern int RemoveMenu(IntPtr hMenu, int nPosition, int wFlags);

    public const int SC_CLOSE = 0xF060;
    public const int MF_BYCOMMAND = 0x00000000;

    public static void DisableCloseButton() {
        IntPtr hWnd = GetConsoleWindow();
        if (hWnd != IntPtr.Zero) {
            IntPtr hMenu = GetSystemMenu(hWnd, false);
            if (hMenu != IntPtr.Zero) {
                RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND);
            }
        }
    }
}
"@
#endregion

# ==============================================================================
# セキュリティ：ワンタイムPIN認証
# ==============================================================================
$script:AuthPin = Get-Random -Minimum 100000 -Maximum 999999
$script:SessionToken = [guid]::NewGuid().ToString('N')
$script:LastAuthFailedTime = [DateTime]::MinValue

# ============================================================================== 
# HTML/CSS/JSテンプレート集約
# ============================================================================== 
#region HTML/CSS/JS Templates
$script:HtmlTemplates = @{
    # 共通HTMLヘッダー + CSS (パラメータ: {0}=Title, {1}=BgColor)
    HtmlHeader = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
    <title>{0}</title>
    <style>
        body {{
            font-family: 'Segoe UI', system-ui, sans-serif;
            background: #000000;
            color: #ffffff;
            text-align: center;
            padding: 20px;
            margin: 0;
            min-height: 100vh;
            position: relative;
            overflow-x: hidden;
            height: 100vh;
            overflow: hidden;
            display: flex;
            flex-direction: column;
        }}
        .container {{
            max-width: 600px;
            width: 100%;
            margin: 0 auto;
            position: relative;
            z-index: 10;
            flex: 1;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            height: 100%;
        }}
        .card {{
            background: #1e1e1e;
            border: 1px solid #333333;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
        }}
        h2 {{ color: #ffffff; margin: 0 0 5px 0; font-size: 1.3rem; }}
        p {{ color: #dcdcdc; font-size: 0.9rem; margin: 5px 0; }}
        .btn {{
            display: block;
            width: 100%;
            padding: 16px;
            margin: 10px 0;
            font-size: 1.1rem;
            border: none;
            border-radius: 12px;
            cursor: pointer;
            color: #ffffff;
            font-weight: bold;
            transition: filter 0.2s ease;
        }}
        .btn:hover {{ filter: brightness(1.15); }}
        .btn-start {{ background: #0d6efd; color: #ffffff; font-size: 1.2rem; padding: 20px; }}
        .btn-stop  {{ background: #dc3545; color: #ffffff; font-size: 1.2rem; padding: 20px; }}
        .btn-next  {{ background: #198754; color: #ffffff; padding: 20px; font-size: 1.2rem; }}
        .btn-retry {{ background: #ffc107; color: #000000; }}
        .btn-list  {{ background: #0dcaf0; color: #000000; }}
        .btn-exit  {{ background: #495057; color: #ffffff; opacity: 0.95; margin-top: 20px; margin-bottom: 50px; }}
        .btn-file {{ background: #2b2b2b; text-align: left; padding: 12px 15px; font-size: 1rem; margin: 5px 0; border-left: 5px solid #0d6efd; color: #ffffff; }}
        .btn-finished {{ background: #121212; border-left: 5px solid #495057; color: #6c757d; }}
        .list-container {{
            text-align: left;
            margin-top: 20px;
            flex-grow: 1;
            overflow-y: auto;
            overflow-x: hidden;
            word-wrap: break-word;
            overflow-wrap: break-word;
            white-space: normal;
        }}
        .list-container::-webkit-scrollbar {{ width: 10px; }}
        .list-container::-webkit-scrollbar-track {{ background: #111111; border-radius: 8px; }}
        .list-container::-webkit-scrollbar-thumb {{ background: #343a40; border-radius: 8px; }}
        .list-container::-webkit-scrollbar-thumb:hover {{ background: #495057; }}
        .loader {{ border: 5px solid #333; border-top: 5px solid #00d2ff; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 20px auto; }}
        .playing-icon {{
            font-size: 3rem;
            margin: 10px;
            animation: pulse 2s infinite;
            color: #198754;
        }}
        .end-icon {{
            font-size: 4rem;
            margin: 20px 0;
            color: #dc3545;
        }}
        #offline-overlay {{ display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.85); z-index: 9999; flex-direction: column; justify-content: center; align-items: center; color: #fff; backdrop-filter: blur(5px); }}
        #offline-overlay.active {{ display: flex; }}
        .offline-icon {{ font-size: 4rem; margin-bottom: 10px; color: #dc3545; animation: pulse 2s infinite; }}
        @keyframes spin {{ 0% {{ transform: rotate(0deg); }} 100% {{ transform: rotate(360deg); }} }}
        @keyframes pulse {{ 0% {{ transform: scale(1); opacity: 1; }} 50% {{ transform: scale(1.1); opacity: 0.8; }} 100% {{ transform: scale(1); opacity: 1; }} }}
    </style>
</head>
<body>
    <div id="offline-overlay">
        <div class="offline-icon">⚠️</div>
        <h2>Connection Lost</h2>
        <p>Connection unstable.<br>Attempting to reconnect...</p>
    </div>
    <div class="container">
"@

    # プレゼンテーション実行中画面 (パラメータ: {0}=FileName)
    NowPlayingView = @"
    <div class="card" style="border: 1px solid #28a745;">
        <div class="playing-icon">▶</div>
        <h2>Now Presenting</h2>
        <p style="font-weight:bold; color:#fff;">{0}</p>
        <p>Controlling slides on PC...</p>
    </div>
    <form method="post" action="/stop">
        <button class="btn btn-stop">■ Stop Presentation</button>
    </form>
    
    <script>
        (function() {{
            var overlay = document.getElementById('offline-overlay');
            var defaultDelay = 1500;
            var currentDelay = defaultDelay;
            var maxDelay = 5000;
            var backoffMultiplier = 1.5;
            
            function pollStatus() {{
                var showOverlayTimer = setTimeout(function() {{
                    if (overlay) overlay.classList.add('active');
                }}, 3000);
                
                fetch('/status?t=' + Date.now())
                .then(response => {{
                    clearTimeout(showOverlayTimer);
                    if (overlay) overlay.classList.remove('active');
                    if (response.ok) {{ return response.text(); }}
                    throw new Error('Network error');
                }})
                .then(text => {{
                    currentDelay = defaultDelay;
                    if (text !== 'running') {{
                        window.location.reload();
                    }} else {{
                        setTimeout(pollStatus, currentDelay);
                    }}
                }})
                .catch(error => {{
                    clearTimeout(showOverlayTimer);
                    if (overlay) overlay.classList.add('active');
                    currentDelay = Math.min(currentDelay * backoffMultiplier, maxDelay);
                    setTimeout(pollStatus, currentDelay);
                }});
            }}
            pollStatus();
        }})();
    </script>
</div></body></html>
"@

    # Lobby画面（スライド一覧） (パラメータ: {0}=stBtn, {1}=nextTxt, {2}=listHtml)
    LobbyView = @"
        <div class="card"><h2>Select Slide</h2><p>Select from list or press Start</p></div>
        <form method="post" action="/start"><button class="btn btn-start" {0}>Start: {1}</button></form>
        {2}
        <form method="post" action="/exit" onsubmit="return confirm('本当にシステムを終了しますか？\n（PC上のプレゼンテーションも強制終了されます）');"><button class="btn btn-exit">Exit System</button></form>
"@

    # プレゼンテーション終了後のダイアログ画面 (パラメータ: {0}=CurrentFileName, {1}=nxtSt, {2}=nxtLbl)
    DialogView = @"
        <div class="card"><h2>Presentation Ended</h2><p>{0}</p></div>
        <form method="post" action="/next"><button class="btn btn-next" {1}>{2}</button></form>
        <form method="post" action="/retry"><button class="btn btn-retry">Play Again</button></form>
        <form method="post" action="/lobby"><button class="btn btn-list">Back to List</button></form>
        <form method="post" action="/exit" onsubmit="return confirm('本当にシステムを終了しますか？\n（PC上のプレゼンテーションも強制終了されます）');"><button class="btn btn-exit">Exit System</button></form>
"@

    # ポーリングスクリプト（Lobby/Dialog用）
    PollingScript = @"
    <script>
        (function() {{
            var overlay = document.getElementById('offline-overlay');
            var defaultDelay = 300;
            var currentDelay = defaultDelay;
            var maxDelay = 5000;
            var backoffMultiplier = 1.5;
            var isPolling = true;
            
            function pollStatus() {{
                if (!isPolling) return;
                
                var showOverlayTimer = setTimeout(function() {{
                    if (overlay) overlay.classList.add('active');
                }}, 3000);
                
                fetch('/status?t=' + Date.now())
                .then(r => {{
                    clearTimeout(showOverlayTimer);
                    if (overlay) overlay.classList.remove('active');
                    return r.text();
                }})
                .then(status => {{
                    currentDelay = defaultDelay;
                    if (status === 'stopping') {{
                        isPolling = false;
                        window.location.href = '/exit';
                    }} else if (status === 'changing' || status === 'starting' || status === 'running') {{
                        isPolling = false;
                        window.location.href = '/';
                    }} else {{
                        setTimeout(pollStatus, currentDelay);
                    }}
                }})
                .catch(e => {{
                    clearTimeout(showOverlayTimer);
                    if (overlay) overlay.classList.add('active');
                    currentDelay = Math.min(currentDelay * backoffMultiplier, maxDelay);
                    setTimeout(pollStatus, currentDelay);
                }});
            }}
            
            document.addEventListener('DOMContentLoaded', function() {{
                var forms = document.querySelectorAll('form');
                forms.forEach(function(form) {{
                    form.addEventListener('submit', function(e) {{
                        if (overlay && overlay.classList.contains('active')) {{
                            e.preventDefault();
                            return;
                        }}
                        isPolling = false;
                    }});
                }});
            }});
            pollStatus();
        }})();
    </script>
"@

    # 処理中画面
    ProcessingView = @"
    <div style="margin-top:50px;"><div class="loader"></div><h2>Processing...</h2><p>Screen will refresh</p></div>
    <script>
        (function() {{
            var overlay = document.getElementById('offline-overlay');
            var defaultDelay = 500;
            var currentDelay = defaultDelay;
            var maxDelay = 5000;
            var backoffMultiplier = 1.5;
            var checkCount = 0;
            var maxRetries = 60;
            var errorCount = 0;
            var maxErrors = 40;
            
            function pollStatus() {{
                var showOverlayTimer = setTimeout(function() {{
                    if (overlay) overlay.classList.add('active');
                }}, 3000);
                
                fetch('/status?t=' + Date.now())
                .then(r => {{
                    clearTimeout(showOverlayTimer);
                    if (overlay) overlay.classList.remove('active');
                    return r.text();
                }})
                .then(status => {{
                    currentDelay = defaultDelay;
                    if (status === 'running' || (status === 'waiting' && checkCount > 2)) {{
                        window.location.href = '/';
                    }} else {{
                        checkCount++;
                        if (checkCount > maxRetries) {{
                            window.location.href = '/';
                        }} else {{
                            setTimeout(pollStatus, currentDelay);
                        }}
                    }}
                }})
                .catch(e => {{
                    clearTimeout(showOverlayTimer);
                    if (overlay) overlay.classList.add('active');
                    currentDelay = Math.min(currentDelay * backoffMultiplier, maxDelay);
                    checkCount++;
                    errorCount++;
                    if (errorCount > maxErrors || checkCount > maxRetries) {{
                        window.location.href = '/';
                    }} else {{
                        setTimeout(pollStatus, currentDelay);
                    }}
                }});
            }}
            pollStatus();
        }})();
    </script>
</body></html>
"@

    # 終了画面
    ExitView = @"
    <div style="margin-top:50px;">
        <div class="end-icon">✔</div>
        <h1>System Shutdown</h1>
        <p style="font-size:1.2rem; color:#fff;">Please close this tab<br>or window.</p>
        <p style="color:#666; margin-top:20px;">Server has been shut down.</p>
    </div>
</body></html>
"@

    # PIN認証画面 (パラメータ: {0}=BgColor, {1}=ErrorFlag "error" or "")
    AuthView = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
    <title>Authentication Required</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #000000;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
            overflow: hidden;
        }}
        .auth-container {{
            position: relative;
            z-index: 10;
            background: #1e1e1e;
            border: 1px solid #333333;
            border-radius: 12px;
            padding: 50px 40px;
            max-width: 450px;
            width: 90%;
            text-align: center;
        }}
        .lock-icon {{
            font-size: 4rem;
            margin-bottom: 20px;
            color: #0d6efd;
        }}
        h1 {{
            color: #fff;
            font-size: 1.8rem;
            margin-bottom: 10px;
            font-weight: 600;
        }}
        .subtitle {{
            color: #aaa;
            font-size: 0.95rem;
            margin-bottom: 40px;
        }}
        .pin-inputs {{
            display: flex;
            justify-content: center;
            gap: 10px;
            margin-bottom: 30px;
        }}
        .pin-inputs.shake {{
            animation: shake 0.5s;
        }}
        @keyframes shake {{
            0%, 100% {{ transform: translateX(0); }}
            10%, 30%, 50%, 70%, 90% {{ transform: translateX(-8px); }}
            20%, 40%, 60%, 80% {{ transform: translateX(8px); }}
        }}
        .pin-box {{
            width: 55px;
            height: 65px;
            font-size: 2rem;
            text-align: center;
            border: 2px solid #444;
            border-radius: 12px;
            background: #2b2b2b;
            color: #fff;
            outline: none;
            transition: all 0.3s;
            caret-color: #0d6efd;
        }}
        .pin-box:focus {{
            border-color: #0d6efd;
            background: #2b2b2b;
        }}
        .pin-box.error {{
            border-color: #dc3545;
            background: rgba(220, 53, 69, 0.1);
        }}
        .error-msg {{
            color: #dc3545;
            font-size: 0.9rem;
            margin-top: -20px;
            margin-bottom: 20px;
            opacity: 0;
            transition: opacity 0.3s;
        }}
        .error-msg.show {{
            opacity: 1;
        }}
        .btn-submit {{
            width: 100%;
            padding: 18px;
            font-size: 1.1rem;
            font-weight: 600;
            border: none;
            border-radius: 12px;
            background: #0d6efd;
            color: #ffffff;
            cursor: pointer;
            transition: filter 0.2s ease;
        }}
        .btn-submit:hover {{
            filter: brightness(1.15);
        }}
        .btn-submit:active {{
            filter: brightness(1.0);
        }}
        .btn-submit:disabled {{
            opacity: 0.5;
            cursor: not-allowed;
        }}
    </style>
</head>
<body>
    <div class="auth-container">
        <div class="lock-icon">🔒</div>
        <h1>Enter PIN Code</h1>
        <p class="subtitle">Please check your PC console for 6-digit PIN</p>
        
        <form method="post" action="/auth" id="authForm">
            <div class="pin-inputs {1}" id="pinInputs">
                <input type="text" class="pin-box {1}" maxlength="1" inputmode="numeric" pattern="[0-9]" autocomplete="off" id="pin1">
                <input type="text" class="pin-box {1}" maxlength="1" inputmode="numeric" pattern="[0-9]" autocomplete="off" id="pin2">
                <input type="text" class="pin-box {1}" maxlength="1" inputmode="numeric" pattern="[0-9]" autocomplete="off" id="pin3">
                <input type="text" class="pin-box {1}" maxlength="1" inputmode="numeric" pattern="[0-9]" autocomplete="off" id="pin4">
                <input type="text" class="pin-box {1}" maxlength="1" inputmode="numeric" pattern="[0-9]" autocomplete="off" id="pin5">
                <input type="text" class="pin-box {1}" maxlength="1" inputmode="numeric" pattern="[0-9]" autocomplete="off" id="pin6">
            </div>
            <div class="error-msg {1}" id="errorMsg">❌ Invalid PIN. Please try again.</div>
            <input type="hidden" name="pin" id="pinValue">
            <button type="submit" class="btn-submit" id="submitBtn" disabled>Unlock</button>
        </form>
    </div>

    <script>
        var boxes = [document.getElementById('pin1'), document.getElementById('pin2'), document.getElementById('pin3'), 
                     document.getElementById('pin4'), document.getElementById('pin5'), document.getElementById('pin6')];
        var submitBtn = document.getElementById('submitBtn');
        var pinValue = document.getElementById('pinValue');
        var form = document.getElementById('authForm');
        var errorMsg = document.getElementById('errorMsg');
        var pinInputsDiv = document.getElementById('pinInputs');
        
        // エラー状態の場合は表示
        var hasError = '{1}' === 'error';
        if (hasError) {{
            errorMsg.classList.add('show');
        }}
        
        boxes.forEach(function(box, index) {{
            // 数字のみ入力可能
            box.addEventListener('input', function(e) {{
                var val = e.target.value;
                if (!/^[0-9]$/.test(val)) {{
                    e.target.value = '';
                    return;
                }}
                
                // エラー状態をクリア
                box.classList.remove('error');
                errorMsg.classList.remove('show');
                pinInputsDiv.classList.remove('shake');
                
                // 次のボックスにフォーカス
                if (val && index < 5) {{
                    boxes[index + 1].focus();
                }}
                
                // すべて入力されたら送信ボタンを有効化
                checkComplete();
            }});
            
            // Backspaceで前のボックスに戻る
            box.addEventListener('keydown', function(e) {{
                if (e.key === 'Backspace' && !e.target.value && index > 0) {{
                    boxes[index - 1].focus();
                }}
            }});
            
            // ペースト対応
            box.addEventListener('paste', function(e) {{
                e.preventDefault();
                var pasteData = e.clipboardData.getData('text').replace(/[^0-9]/g, '').substring(0, 6);
                for (var j = 0; j < boxes.length; j++) {{
                    boxes[j].value = '';
                }}
                for (var i = 0; i < pasteData.length && i < 6; i++) {{
                    boxes[i].value = pasteData[i];
                }}
                if (pasteData.length < 6) {{
                    boxes[pasteData.length].focus();
                }} else {{
                    boxes[5].focus();
                }}
                checkComplete();
            }});
        }});
        
        function checkComplete() {{
            var complete = boxes.every(function(b) {{ return b.value.length === 1; }});
            submitBtn.disabled = !complete;
        }}
        
        // フォーム送信時に6桁を結合
        form.addEventListener('submit', function() {{
            pinValue.value = boxes.map(function(b) {{ return b.value; }}).join('');
        }});
        
        // 最初のボックスにフォーカス
        boxes[0].focus();
    </script>
</body>
</html>
"@
}
#endregion

# ------------------------------------------------------------
# 1. ユーティリティ関数
# ------------------------------------------------------------
function Get-LocalActiveIPs {
    try {
        # Get all adapters that are up and not virtual
        $adapters = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object {
            $_.Status -eq "Up" -and
            -not $_.Virtual -and
            $_.InterfaceAlias -notlike "*Loopback*" -and
            $_.InterfaceAlias -notlike "*vEthernet*" -and
            $_.InterfaceAlias -notlike "*VMware*" -and
            $_.InterfaceAlias -notlike "*VirtualBox*" -and
            $_.InterfaceAlias -notlike "*Tailscale*" -and
            $_.InterfaceAlias -notlike "*ZeroTier*" -and
            $_.InterfaceDescription -notlike "*Loopback*" -and
            $_.InterfaceDescription -notlike "*vEthernet*" -and
            $_.InterfaceDescription -notlike "*VMware*" -and
            $_.InterfaceDescription -notlike "*VirtualBox*" -and
            $_.InterfaceDescription -notlike "*Tailscale*" -and
            $_.InterfaceDescription -notlike "*ZeroTier*"
        }
        
        # Get IP addresses for each adapter
        $results = @()
        foreach ($adapter in $adapters) {
            $ipAddresses = Get-NetIPAddress -InterfaceIndex $adapter.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {
                $_.IPAddress -notlike "169.254.*" -and $_.IPAddress -notlike "0.0.0.0"
            }
            
            foreach ($ipAddr in $ipAddresses) {
                $results += @{
                    InterfaceAlias = $adapter.InterfaceAlias
                    IPAddress = $ipAddr.IPAddress
                }
            }
        }
        
        # Fallback if no valid IPs found
        if ($results.Count -eq 0) {
            $results = @(@{ InterfaceAlias = "Local"; IPAddress = "localhost" })
        }
        
        return $results
    } catch {
        # Fallback on error
        return @(@{ InterfaceAlias = "Local"; IPAddress = "localhost" })
    }
}

function Release-ComObject {
    param([object]$obj)
    if ($obj) { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {} }
}

# 安全にレスポンスを返すためのラッパー関数 (今回の修正の要)
function Send-HttpResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [string]$Content,
        [string]$ContentType = "text/html; charset=utf-8"
    )

    try {
        if ($Response.OutputStream.CanWrite) {
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($Content)
            $Response.ContentType = $ContentType
            $Response.ContentLength64 = $buffer.Length
            $Response.KeepAlive = $false
            $Response.OutputStream.Write($buffer, 0, $buffer.Length)
            $Response.OutputStream.Close()
        }
    } catch {
        # ここで "The specified network name is no longer available" を握りつぶす
        # クライアントが切断されているため、エラーを出さずに無視してよい
    } finally {
        try { $Response.Close() } catch {}
    }
}

# ============================================================================== 
# 2. 共通HTMLヘッダー・スタイル
# ============================================================================== 
function Get-HtmlHeader {
    param([string]$Title, [string]$BgColor="#1a1a1a")
    return $script:HtmlTemplates.HtmlHeader -f $Title, $BgColor
}

# ============================================================================== 
# 3. 発表中の監視関数
# ==============================================================================
function Watch-RunningPresentation {
    param (
        [object]$PptApp,
        [object]$TargetFileItem,
        [System.Net.HttpListener]$Listener
    )

    $head = Get-HtmlHeader -Title "Now Playing" -BgColor "#000000"
    $bodyHtml = $script:HtmlTemplates.NowPlayingView -f $TargetFileItem.Name
    $fullHtml = $head + $bodyHtml

    $status = "NormalEnd"

    try {
        $isFileOpen = $true
        while ($isFileOpen) {
            # 1. Webリクエスト確認
            if ($script:ContextTask -and $script:ContextTask.AsyncWaitHandle.WaitOne(100)) {
                $context = $script:ContextTask.Result
                $req = $context.Request
                $res = $context.Response
                $path = $req.Url.LocalPath.ToLower()

                if ($path -eq "/status") {
                    # JSからの生存確認用：発表中は "running" を返す
                    Send-HttpResponse -Response $res -Content "running" -ContentType "text/plain"
                } 
                elseif ($path -eq "/stop" -and $req.HttpMethod -eq "POST") {
                    $status = "ManualStop"
                    try {
                        $res.StatusCode = 302
                        $res.KeepAlive = $false
                        $res.AddHeader("Location", "/")
                        $res.Close()
                    } catch {}
                    $script:ContextTask = $Listener.GetContextAsync()
                    break 
                } 
                else {
                    # その他のアクセスには「発表中画面」を返す
                    Send-HttpResponse -Response $res -Content $fullHtml
                }
                
                # 次のリクエスト待ち準備
                $script:ContextTask = $Listener.GetContextAsync()
            }

            # 2. コンソール入力確認 (Qキー)
            if ([Console]::KeyAvailable) {
                $k = [Console]::ReadKey($true).Key.ToString().ToUpper()
                if ($k -eq "Q" -or $k -eq "ESCAPE") {
                    $status = "ManualStop"
                    break
                }
            }

            # 3. PowerPointの状態確認
            $stillOpen = $false
            try {
                foreach ($p in $PptApp.Presentations) {
                    if ($p.FullName -eq $TargetFileItem.FullName) {
                        $stillOpen = $true
                        break
                    }
                }
            } catch { $stillOpen = $false }
            
            if (-not $stillOpen) {
                $status = "NormalEnd"
                break
            }
        }
    } finally {
        # HttpListener はメインフロー内で一元管理するため、ここでは Stop/Close しない
    }

    return $status
}

# ------------------------------------------------------------
# 4. 入力待機関数 (Lobby / Dialog)
# ------------------------------------------------------------
function Get-UserAction {
    param (
        [string]$Mode,
        [string]$CurrentFileName = "",
        [array]$ActiveFiles = @(),
        [array]$FinishedFiles = @(),
        [string]$NextFileName = "",
        [System.Net.HttpListener]$Listener
    )

    # ページング用変数
    $currentPage = 0
    $itemsPerPage = 9
    $needsRedraw = $true

    # 画面表示関数
    function Show-ConsolePage {
        Clear-Host
        $adapters = Get-LocalActiveIPs
        $line = "━" * 70
        Write-Host $line -ForegroundColor DarkCyan
        Write-Host "  [ ppt-orchestrator ] " -ForegroundColor Cyan -NoNewline
        Write-Host "v7.4 - Presentation Controller" -ForegroundColor DarkGray
        Write-Host $line -ForegroundColor DarkCyan
        Write-Host ""
        Write-Host "   🔐 PIN CODE: " -NoNewline -ForegroundColor Yellow
        Write-Host $script:AuthPin -ForegroundColor White -BackgroundColor DarkRed
        Write-Host ""
        foreach ($adapter in $adapters) {
            Write-Host " [Web URL - $($adapter.InterfaceAlias)] http://$($adapter.IPAddress):$($WebPort)/" -ForegroundColor Yellow
        }
        Write-Host " [Status]   $Mode" -ForegroundColor White
        Write-Host ""
        Write-Host " --- PC Control Menu ---" -ForegroundColor Gray
        if ($Mode -eq "Lobby") {
            Write-Host " [Enter] Start" -ForegroundColor Green
            Write-Host " [1-9]   Select Slide by Number" -ForegroundColor Cyan
            
            # ページング情報の表示
            $totalActiveFiles = if ($ActiveFiles) { @($ActiveFiles).Count } else { 0 }
            $totalFinishedFiles = if ($FinishedFiles) { @($FinishedFiles).Count } else { 0 }
            $totalFiles = $totalActiveFiles + $totalFinishedFiles
            $totalPages = [Math]::Ceiling($totalFiles / $itemsPerPage)
            
            if ($totalPages -gt 1) {
                Write-Host " [N]     Next Page  [P] Previous Page" -ForegroundColor Magenta
            }
            Write-Host " [Q]     Exit System" -ForegroundColor Red
            Write-Host "   * Note: To close a presentation, please click the 'X' button on the PowerPoint window." -ForegroundColor DarkGray
            Write-Host ""
            
            if ($totalPages -gt 1) {
                Write-Host " --- Available Slides (Page $($currentPage + 1)/$totalPages) ---" -ForegroundColor Gray
            } else {
                Write-Host " --- Available Slides ---" -ForegroundColor Gray
            }
            
            # ページング計算
            $startIdx = $currentPage * $itemsPerPage
            $endIdx = $startIdx + $itemsPerPage - 1
            
            # Pending スライド表示
            $displayIndex = 1
            $currentFileIndex = 0
            $activeSectionShown = $false
            
            if ($ActiveFiles -and @($ActiveFiles).Count -gt 0) {
                foreach ($f in $ActiveFiles) {
                    if ($currentFileIndex -ge $startIdx -and $currentFileIndex -le $endIdx) {
                        if (-not $activeSectionShown) {
                            Write-Host " [Pending]" -ForegroundColor Green
                            $activeSectionShown = $true
                        }
                        Write-Host "  [$displayIndex] $($f.Name)" -ForegroundColor White
                        $displayIndex++
                    }
                    $currentFileIndex++
                }
            }
            
            # Completed スライド表示
            $finishedSectionShown = $false
            if ($FinishedFiles -and @($FinishedFiles).Count -gt 0) {
                foreach ($f in $FinishedFiles) {
                    if ($currentFileIndex -ge $startIdx -and $currentFileIndex -le $endIdx) {
                        if (-not $finishedSectionShown) {
                            Write-Host " [Completed]" -ForegroundColor DarkGray
                            $finishedSectionShown = $true
                        }
                        Write-Host "  [$displayIndex] $($f.Name)" -ForegroundColor DarkGray
                        $displayIndex++
                    }
                    $currentFileIndex++
                }
            }
        } else {
            Write-Host " [Enter] Next" -ForegroundColor Green
            Write-Host " [R]     Retry" -ForegroundColor Yellow
            Write-Host " [L]     Back to Lobby" -ForegroundColor Cyan
            Write-Host " [Q]     Exit System" -ForegroundColor Red
            Write-Host "   * Note: To close a presentation, please click the 'X' button on the PowerPoint window." -ForegroundColor DarkGray
        }
            Write-Host ""
            Write-Host " ▶ Waiting for command... (Press a key to execute immediately)" -ForegroundColor Green
            Write-Host $line -ForegroundColor DarkCyan
            Write-Host "  Copyright (c) 2026 FumiakiC" -ForegroundColor DarkGray
            Write-Host ""
        }
    
        # 初回表示
        Show-ConsolePage
    $head = Get-HtmlHeader -Title "Controller" -BgColor $(if($Mode -eq 'Lobby'){"#1a1a1a"}else{"#000000"})
    
    # Bodyコンテンツ
    $bodyContent = ""
    if ($Mode -eq 'Lobby') {
        $nextTxt = if ($ActiveFiles) { $ActiveFiles[0].Name } else { "None" }
        $stBtn = if ($ActiveFiles) { "" } else { "disabled style='opacity:0.5;'" }
        
        $listHtml = "<div class='list-container'><h3 style='color:#28a745;font-size:0.9rem;border-bottom:1px solid #555;'>Pending</h3>"
        if (!$ActiveFiles) { $listHtml += "<p>None</p>" }
        foreach ($f in $ActiveFiles) {
            $listHtml += "<form method='post' action='/select' style='margin:0;'><input type='hidden' name='filename' value='$($f.Name)'><button class='btn btn-file'>$($f.Name)</button></form>"
        }
        $listHtml += "<h3 style='color:#6c757d;font-size:0.9rem;border-bottom:1px solid #555;margin-top:20px;'>Completed</h3>"
        foreach ($f in $FinishedFiles) {
            $listHtml += "<form method='post' action='/select' style='margin:0;'><input type='hidden' name='filename' value='$($f.Name)'><button class='btn btn-file btn-finished'>$($f.Name)</button></form>"
        }
        $listHtml += "</div>"

        $bodyContent = $script:HtmlTemplates.LobbyView -f $stBtn, $nextTxt, $listHtml
    } else {
        $nxtLbl = if ($NextFileName) { "Start Next Slide<br><span style='font-size:0.8rem;font-weight:normal'>$NextFileName</span>" } else { "No slides in queue" }
        $nxtSt = if ($NextFileName) { "" } else { "disabled style='opacity:0.5;'" }

        $bodyContent = $script:HtmlTemplates.DialogView -f $CurrentFileName, $nxtSt, $nxtLbl
    }
    
    # 完全なHTMLページを構築
    $mainHtml = $head + $bodyContent + $script:HtmlTemplates.PollingScript + "</div></body></html>"

    # --- 画面状態のHTML ---
    $processingHtml = $head + $script:HtmlTemplates.ProcessingView
    
    $exitHtml = $head + $script:HtmlTemplates.ExitView

    $resultAction = $null
    $resultFile = $null
    $actionSetTime = $null  # アクション設定時刻を記録
    
    # シャットダウン制御用
    $shuttingDown = $false
    $shutdownDeadline = $null
    $waitingExitConfirm = $false

    while ($true) {
        
        # --- Web確認 ---
        if ($script:ContextTask.AsyncWaitHandle.WaitOne(100)) {
            $context = $script:ContextTask.Result
            $req = $context.Request
            $res = $context.Response
            $url = $req.Url.LocalPath.ToLower()
            
            # --- 認証ミドルウェア：Cookie確認 ---
            $isAuthenticated = $false
            if ($req.Cookies["SessionToken"]) {
                if ($req.Cookies["SessionToken"].Value -eq $script:SessionToken) {
                    $isAuthenticated = $true
                }
            }
            
            # 認証が必要なパス（/authと/statusは除外）
            if (-not $isAuthenticated -and $url -ne "/auth" -and $url -ne "/status") {
                $authHtml = $script:HtmlTemplates.AuthView -f "#0f2027", ""
                Send-HttpResponse -Response $res -Content $authHtml
                $script:ContextTask = $Listener.GetContextAsync()
                continue
            }
            
            # /auth POSTリクエスト処理
            if ($url -eq "/auth" -and $req.HttpMethod -eq "POST") {
                # レートリミット処理：認証失敗後1秒以内のリクエストを即座にエラーで返す
                $currentTime = Get-Date
                if ($currentTime -lt $script:LastAuthFailedTime.AddSeconds(1)) {
                    $authHtml = $script:HtmlTemplates.AuthView -f "#0f2027", "error"
                    Send-HttpResponse -Response $res -Content $authHtml
                    $script:ContextTask = $Listener.GetContextAsync()
                    continue
                }
                
                if ($req.HasEntityBody) {
                    $r = New-Object System.IO.StreamReader($req.InputStream, $req.ContentEncoding)
                    $body = $r.ReadToEnd(); $r.Close()
                    
                    if ([System.Web.HttpUtility]::UrlDecode($body) -match "pin=([0-9]{6})") {
                        $submittedPin = $matches[1]
                        if ($submittedPin -eq $script:AuthPin.ToString()) {
                            # 認証成功：Cookieをセットしてリダイレクト
                            $res.Headers.Add("Set-Cookie", "SessionToken=$script:SessionToken; HttpOnly; Path=/; SameSite=Strict")
                            $res.StatusCode = 302
                            $res.Headers.Add("Location", "/")
                            Send-HttpResponse -Response $res -Content ""
                            $script:ContextTask = $Listener.GetContextAsync()
                            continue
                        }
                    }
                }
                # 認証失敗：エラー表示
                $script:LastAuthFailedTime = Get-Date
                $authHtml = $script:HtmlTemplates.AuthView -f "#0f2027", "error"
                Send-HttpResponse -Response $res -Content $authHtml
                $script:ContextTask = $Listener.GetContextAsync()
                continue
            }
            
            $resHtml = $mainHtml
            
            if ($url -eq "/status") {
                # 状態を返す：stopping（終了中）/ changing（状態変更中）/ waiting（待機中）
                $statusText = if ($shuttingDown) { 
                    "stopping" 
                } elseif ($resultAction -ne $null) { 
                    "changing" 
                } else { 
                    "waiting" 
                }
                Send-HttpResponse -Response $res -Content $statusText -ContentType "text/plain"
                
                $script:ContextTask = $Listener.GetContextAsync()
                continue
            }

            if ($req.HttpMethod -eq "POST") {
                if ($req.HasEntityBody) {
                    $r = New-Object System.IO.StreamReader($req.InputStream, $req.ContentEncoding)
                    $body = $r.ReadToEnd(); $r.Close()
                }
                switch ($url) {
                    "/start"  { $resultAction = "Start"; $actionSetTime = Get-Date; $resHtml = $processingHtml }
                    "/next"   { $resultAction = "Next";  $actionSetTime = Get-Date; $resHtml = $processingHtml }
                    "/retry"  { $resultAction = "Retry"; $actionSetTime = Get-Date; $resHtml = $processingHtml }
                    "/lobby"  { $resultAction = "Lobby"; $actionSetTime = Get-Date; $resHtml = $processingHtml }
                    "/exit"   { 
                        $resultAction = "Exit"
                        $actionSetTime = Get-Date
                        $shuttingDown = $true 
                        $resHtml = $exitHtml 
                    }
                    "/select" {
                        if ([System.Web.HttpUtility]::UrlDecode($body) -match "filename=(.*)") { 
                            $resultAction = "Select"; $resultFile = $matches[1]; $actionSetTime = Get-Date
                        }
                        $resHtml = $processingHtml
                    }
                }
            } elseif ($url -eq "/exit") {
                $resHtml = $exitHtml
            }

            # 安全にレスポンスを返す
            Send-HttpResponse -Response $res -Content $resHtml

            if ($resultAction -eq "Exit" -and -not $shutdownDeadline) {
                 Start-Sleep -Milliseconds 500
                 break
            }

            $script:ContextTask = $Listener.GetContextAsync()
        }

        # --- コンソール確認 ---
        if ((!$shuttingDown) -and ($resultAction -eq $null) -and [Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true).Key.ToString().ToUpper()
            
            # 終了確認待ちの場合
            if ($waitingExitConfirm) {
                if ($k -eq "Y") {
                    $shuttingDown = $true
                    $shutdownDeadline = (Get-Date).AddSeconds(5)
                    Write-Host ""
                    Write-Host " [System] Shutting down... (Notifying web clients / Will exit in 5 seconds)" -ForegroundColor Magenta
                } else {
                    # キャンセル
                    $waitingExitConfirm = $false
                    Show-ConsolePage
                }
            } else {
                # 通常コマンド処理
                
                # コンソールからの終了要求
                if ($k -eq "Q" -or $k -eq "ESCAPE") {
                    $waitingExitConfirm = $true
                    Write-Host ""
                    Write-Host " Are you sure you want to exit? [Y] Confirm / [N] Cancel : " -ForegroundColor Yellow -NoNewline
                }

                if ($Mode -eq "Lobby") {
                    if ($k -eq "ENTER" -or $k -eq "S") { $resultAction = "Start"; $actionSetTime = Get-Date }
                    
                    # ページング操作
                    $totalActiveFiles = if ($ActiveFiles) { @($ActiveFiles).Count } else { 0 }
                    $totalFinishedFiles = if ($FinishedFiles) { @($FinishedFiles).Count } else { 0 }
                    $totalFiles = $totalActiveFiles + $totalFinishedFiles
                    $totalPages = [Math]::Ceiling($totalFiles / $itemsPerPage)
                    
                    if ($k -eq "N") {
                        # 次のページ
                        if ($currentPage -lt ($totalPages - 1)) {
                            $currentPage++
                            Show-ConsolePage
                        }
                    }
                    elseif ($k -eq "P") {
                        # 前のページ
                        if ($currentPage -gt 0) {
                            $currentPage--
                            Show-ConsolePage
                        }
                    }
                    
                    # 数字キーでスライド選択（ページオフセットを考慮）
                    if ($k -match "^D([0-9])$" -or $k -match "^NUMPAD([0-9])$") {
                        # D1-D9 および NUMPAD1-NUMPAD9 形式のキー
                        $num = [int]$matches[1]
                        if ($num -ge 1 -and $num -le 9) {
                            $absoluteIndex = $currentPage * $itemsPerPage + ($num - 1)
                            
                            # 全ファイルリストを作成
                            $allFiles = @()
                            if ($ActiveFiles) { $allFiles += $ActiveFiles }
                            if ($FinishedFiles) { $allFiles += $FinishedFiles }
                            
                            if ($absoluteIndex -lt $allFiles.Count) {
                                $resultAction = "Select"
                                $resultFile = $allFiles[$absoluteIndex].Name
                                $actionSetTime = Get-Date
                            }
                        }
                    }
                } else {
                    if ($k -eq "ENTER" -or $k -eq "N") { $resultAction = "Next"; $actionSetTime = Get-Date }
                    if ($k -eq "R") { $resultAction = "Retry"; $actionSetTime = Get-Date }
                    if ($k -eq "L" -or $k -eq "BACKSPACE") { $resultAction = "Lobby"; $actionSetTime = Get-Date }
                }
            }
        }

        # --- 終了判定 ---
        if ($resultAction -ne $null -and $resultAction -ne "Exit") {
            # Web側が状態変化を検知できるよう、800ms待機してからbreak
            if ($actionSetTime -and ((Get-Date) - $actionSetTime).TotalMilliseconds -gt 800) {
                break
            }
        }
        if ($shuttingDown -and $shutdownDeadline) {
            if ((Get-Date) -gt $shutdownDeadline) {
                $resultAction = "Exit"
                break
            }
        }
    }

    # HttpListener はメインフロー内で一元管理するため、ここでは Stop/Close しない
    
    return @{ Action = $resultAction; FileName = $resultFile }
}

# ============================================================================== 
# 5. メインフロー
# ==============================================================================
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Administrator privileges required. Please run PowerShell as Administrator."
    Start-Sleep 3
    exit
}

# コンソールの「閉じる」ボタンを無効化（誤操作防止）
[ConsoleWindow]::DisableCloseButton()

# Ctrl+Cをスクリプト終了ではなく通常のキー入力として処理
[console]::TreatControlCAsInput = $true

if (-not (Test-Path $TargetFolderPath)) { Write-Error "Target Folder Not Found"; exit }
$finishFolderPath = Join-Path $TargetFolderPath $FinishFolderName
if (-not (Test-Path $finishFolderPath)) { New-Item -Path $finishFolderPath -ItemType Directory | Out-Null }

Write-Host "Starting PowerPoint..." -ForegroundColor Cyan
try {
    $pptApp = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
} catch {
    Write-Error "Failed to start PowerPoint"
    exit
}

try {
    $exitLoop = $false
    $autoPlayTarget = $null
    
    # HttpListener を作成・Start（メインフロー内で一元管理）
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://+:$WebPort/")
    try {
        $listener.Start()
        $script:ContextTask = $listener.GetContextAsync()
    } catch {
        Write-Warning "Web control is unavailable due to port conflict. Only keyboard operations are available."
    }

    while (-not $exitLoop) {
        
        $activeFiles = Get-ChildItem -Path $TargetFolderPath -File | Where-Object { $_.Extension -in @('.ppt', '.pptx') -and $_.Name -notlike '~$*' } | Sort-Object Name
        $finishedFiles = Get-ChildItem -Path $finishFolderPath -File | Where-Object { $_.Extension -in @('.ppt', '.pptx') -and $_.Name -notlike '~$*' } | Sort-Object Name
        
        $targetFileItem = $null

        # --- A. 選択 ---
        if ($autoPlayTarget) {
            $targetFileItem = $autoPlayTarget
            $autoPlayTarget = $null
        } else {
            $result = Get-UserAction -Mode "Lobby" -ActiveFiles $activeFiles -FinishedFiles $finishedFiles -Listener $listener
            
            switch ($result.Action) {
                "Exit" { $exitLoop = $true; break }
                "Start" { if ($activeFiles) { $targetFileItem = $activeFiles[0] } }
                "Select" {
                    $name = $result.FileName
                    $targetFileItem = $activeFiles | Where-Object { $_.Name -eq $name } | Select-Object -First 1
                    if (!$targetFileItem) { $targetFileItem = $finishedFiles | Where-Object { $_.Name -eq $name } | Select-Object -First 1 }
                }
            }
        }

        if (-not $targetFileItem) { continue }
        if ($exitLoop) { break }

        # --- B. プレゼン実行 ---
        $presentation = $null
        $status = "NormalEnd"

        try {
            Write-Host " >> Opening: $($targetFileItem.Name)" -ForegroundColor Cyan
            $presentation = $pptApp.Presentations.Open($targetFileItem.FullName, $false, $false, $true)
            
            # PowerPointのファイル読み込み（COMオブジェクト）が完了し、プロセスが安定するのを待機
            Start-Sleep -Milliseconds 100
            $presentation.SlideShowSettings.Run() | Out-Null
            
            # 発表中の監視
            $status = Watch-RunningPresentation -PptApp $pptApp -TargetFileItem $targetFileItem -Listener $listener
            
            # 手動終了(ManualStop)の場合はここで閉じる
            if ($status -eq "ManualStop") {
                Write-Host " >> Manually stopped." -ForegroundColor Yellow
                try { $presentation.Close() } catch {}
            }

            # --- C. 移動判定 ---
            if ($targetFileItem.DirectoryName -ne $finishFolderPath) {
                try {
                    Write-Host " >> Moving to finished folder..." -ForegroundColor Gray
                    $targetFileItem = Move-Item -LiteralPath $targetFileItem.FullName -Destination $finishFolderPath -Force -PassThru
                } catch { Write-Warning "Move failed: $_" }
            }
            
            # --- D. 終了後の画面遷移 ---
            if ($status -eq "ManualStop") {
                # 手動終了の場合：ダイアログを出さずにLobbyへ戻る
                $autoPlayTarget = $null
                continue 
            }

            # 正常終了の場合：ダイアログ(次へ/リトライ)を表示
            $activeFiles = Get-ChildItem -Path $TargetFolderPath -File | Where-Object { $_.Extension -in @('.ppt', '.pptx') -and $_.Name -notlike '~$*' } | Sort-Object Name
            $nextName = if ($activeFiles) { $activeFiles[0].Name } else { "" }

            $postResult = Get-UserAction -Mode "Dialog" -CurrentFileName $targetFileItem.Name -NextFileName $nextName -Listener $listener

            switch ($postResult.Action) {
                "Next"  { if ($activeFiles) { $autoPlayTarget = $activeFiles[0] } }
                "Retry" { $autoPlayTarget = $targetFileItem }
                "Lobby" { $autoPlayTarget = $null }
                "Exit"  { $exitLoop = $true }
            }

        } catch {
            Write-Error "Error: $($_.Exception.Message)"
            if ($presentation) { try { $presentation.Close() } catch {} }
            Start-Sleep 2
        } finally {
            # COMオブジェクトの確実なクリーンアップ
            if ($presentation) {
                try { $presentation.Close() } catch {}
                Release-ComObject -obj $presentation
                $presentation = $null
            }
            # ガベージコレクションを強制実行してCOM参照を完全に解放
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }

} finally {
    # HttpListener を停止・閉鎖
    if ($listener -and $listener.IsListening) {
        try { $listener.Stop(); $listener.Close(); Start-Sleep -Milliseconds 200 } catch {}
    }
    
    # コンソール画面をクリアして終了処理を明示
    Clear-Host
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "  [System] Shutting down..." -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""
    
    # PowerPointアプリケーション全体の終了処理
    if ($pptApp) { 
        try { $pptApp.Quit() } catch {}
        Release-ComObject -obj $pptApp
        $pptApp = $null
    }
    
    # ガベージコレクションを強制実行してPOWERPNT.EXEプロセスを確実に終了
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "System terminated." -ForegroundColor Green
    Write-Host ""
    
    [Environment]::Exit(0)
}