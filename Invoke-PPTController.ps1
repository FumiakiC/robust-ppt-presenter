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

# ------------------------------------------------------------
# 1. ユーティリティ関数
# ------------------------------------------------------------
function Get-LocalIPAddress {
    try {
        $ip = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { 
            $_.InterfaceAlias -notlike "*Loopback*" -and $_.IPAddress -notlike "169.254.*" -and $_.IPAddress -notlike "0.0.0.0"
        } | Select-Object -ExpandProperty IPAddress -First 1)
        
        if (-not $ip) {
            $ip = ([System.Net.Dns]::GetHostAddresses([System.Net.Dns]::GetHostName()) | 
                   Where-Object { $_.AddressFamily -eq 'InterNetwork' } | 
                   Select-Object -ExpandProperty IPAddressToString -First 1)
        }
        return $ip
    } catch {
        return "localhost"
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
            $Response.OutputStream.Write($buffer, 0, $buffer.Length)
        }
    } catch {
        # ここで "The specified network name is no longer available" を握りつぶす
        # クライアントが切断されているため、エラーを出さずに無視してよい
    } finally {
        try { $Response.Close() } catch {}
    }
}

# ------------------------------------------------------------
# 2. 共通HTMLヘッダー・スタイル
# ------------------------------------------------------------
function Get-HtmlHeader {
    param([string]$Title, [string]$BgColor="#1a1a1a")
    return @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
    <title>$Title</title>
    <style>
        body { font-family: sans-serif; background: $BgColor; color: #fff; text-align: center; padding: 20px; margin: 0; }
        .container { max-width: 600px; margin: 0 auto; }
        .card { background: #333; padding: 20px; border-radius: 15px; margin-bottom: 20px; }
        h2 { color: #00d2ff; margin: 0 0 5px 0; font-size: 1.3rem; }
        p { color: #ccc; font-size: 0.9rem; margin: 5px 0; }
        .btn { display: block; width: 100%; padding: 16px; margin: 10px 0; font-size: 1.1rem; border: none; border-radius: 10px; cursor: pointer; color: white; font-weight: bold; }
        .btn-start { background: linear-gradient(135deg, #007bff, #0056b3); font-size: 1.2rem; padding: 20px; }
        .btn-stop  { background: linear-gradient(135deg, #dc3545, #a71d2a); font-size: 1.2rem; padding: 20px; box-shadow: 0 4px 10px rgba(220,53,69,0.4); }
        .btn-next  { background: linear-gradient(135deg, #28a745, #218838); padding: 20px; font-size: 1.2rem; }
        .btn-retry { background: linear-gradient(135deg, #ffc107, #e0a800); color: #000; }
        .btn-list  { background: #17a2b8; }
        .btn-exit  { background: #6c757d; opacity: 0.8; margin-top: 30px; }
        .btn-file { background: #444; text-align: left; padding: 12px 15px; font-size: 1rem; margin: 5px 0; border-left: 5px solid #00d2ff; }
        .btn-finished { background: #2a2a2a; border-left: 5px solid #6c757d; color: #aaa; }
        .list-container { text-align: left; margin-top: 20px; max-height: 40vh; overflow-y: auto; }
        .loader { border: 5px solid #333; border-top: 5px solid #00d2ff; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 20px auto; }
        .playing-icon { font-size: 3rem; color: #28a745; margin: 10px; animation: pulse 2s infinite; }
        .end-icon { font-size: 4rem; color: #dc3545; margin: 20px 0; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        @keyframes pulse { 0% { transform: scale(1); opacity: 1; } 50% { transform: scale(1.1); opacity: 0.8; } 100% { transform: scale(1); opacity: 1; } }
    </style>
</head>
<body>
    <div class="container">
"@
}

# ------------------------------------------------------------
# 3. 発表中の監視関数
# ------------------------------------------------------------
function Watch-RunningPresentation {
    param (
        [object]$PptApp,
        [object]$TargetFileItem
    )

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://+:$WebPort/")
    try {
        $listener.Start()
    } catch {
        Write-Warning "Web control is unavailable due to port conflict. Only keyboard operations are available."
    }

    $head = Get-HtmlHeader -Title "Now Playing" -BgColor "#000000"
    $bodyHtml = @"
    <div class="card" style="border: 1px solid #28a745;">
        <div class="playing-icon">▶</div>
        <h2>Now Presenting</h2>
        <p style="font-weight:bold; color:#fff;">$($TargetFileItem.Name)</p>
        <p>Controlling slides on PC...</p>
    </div>
    <form method="post" action="/stop">
        <button class="btn btn-stop">■ Stop Presentation</button>
    </form>
    
    <script>
        // 定期的に状態を確認し、発表が終わっていたらリロードする
        setInterval(function(){
            // キャッシュ対策で時刻を付与
            fetch('/status?t=' + Date.now())
            .then(response => {
                if (response.ok) { return response.text(); }
                throw new Error('Network error');
            })
            .then(text => {
                // サーバーから 'running' 以外（waiting/stoppingなど）が返ってきたら画面遷移(リロード)
                if (text !== 'running') {
                    window.location.reload();
                }
            })
            .catch(error => {
                // サーバー停止/切り替え時もリロードを試みる
                setTimeout(() => window.location.reload(), 1000);
            });
        }, 1500);
    </script>
</div></body></html>
"@
    $fullHtml = $head + $bodyHtml

    $status = "NormalEnd"
    $contextTask = if ($listener.IsListening) { $listener.GetContextAsync() } else { $null }

    try {
        $isFileOpen = $true
        while ($isFileOpen) {
            # 1. Webリクエスト確認
            if ($contextTask -and $contextTask.AsyncWaitHandle.WaitOne(100)) {
                $context = $contextTask.Result
                $req = $context.Request
                $res = $context.Response
                $path = $req.Url.LocalPath.ToLower()

                if ($path -eq "/status") {
                    # JSからの生存確認用：発表中は "running" を返す
                    Send-HttpResponse -Response $res -Content "running" -ContentType "text/plain"
                } 
                elseif ($path -eq "/stop" -and $req.HttpMethod -eq "POST") {
                    # 強制終了ボタン
                    $status = "ManualStop"
                    try {
                        $res.StatusCode = 302
                        $res.AddHeader("Location", "/")
                        $res.Close()
                    } catch {}
                    break 
                } 
                else {
                    # その他のアクセスには「発表中画面」を返す
                    Send-HttpResponse -Response $res -Content $fullHtml
                }
                
                # 次のリクエスト待ち準備
                $contextTask = $listener.GetContextAsync()
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
        if ($listener.IsListening) {
            $listener.Stop()
            $listener.Close()
            Start-Sleep -Milliseconds 200
        }
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
        [string]$NextFileName = ""
    )

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://+:$WebPort/")
    try { $listener.Start() } catch { Write-Error "ポート $WebPort 使用不可"; exit }

    # ページング用変数
    $currentPage = 0
    $itemsPerPage = 9
    $needsRedraw = $true

    # 画面表示関数
    function Show-ConsolePage {
        Clear-Host
        $ip = Get-LocalIPAddress
        $line = "=" * 70
        Write-Host $line -ForegroundColor Cyan
        Write-Host "   Presentation Controller V7.3" -ForegroundColor White -BackgroundColor DarkCyan
        Write-Host $line -ForegroundColor Cyan
        Write-Host " [Web URL]  http://$($ip):$($WebPort)/" -ForegroundColor Yellow
        Write-Host " [Status]   $Mode" -ForegroundColor White
        Write-Host ""
        Write-Host " --- PC Control Menu ---" -ForegroundColor Gray
        if ($Mode -eq "Lobby") {
            Write-Host " [Enter] Start" -ForegroundColor Green
            Write-Host " [1-9]   Select Slide by Number" -ForegroundColor Cyan
            
            # ページング情報の表示
            $totalActiveFiles = if ($ActiveFiles) { $ActiveFiles.Count } else { 0 }
            $totalFinishedFiles = if ($FinishedFiles) { $FinishedFiles.Count } else { 0 }
            $totalFiles = $totalActiveFiles + $totalFinishedFiles
            $totalPages = [Math]::Ceiling($totalFiles / $itemsPerPage)
            
            if ($totalPages -gt 1) {
                Write-Host " [N]     Next Page  [P] Previous Page" -ForegroundColor Magenta
            }
            Write-Host " [Q]     Exit System" -ForegroundColor Red
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
            
            if ($ActiveFiles -and $ActiveFiles.Count -gt 0) {
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
            if ($FinishedFiles -and $FinishedFiles.Count -gt 0) {
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
        }
        Write-Host $line -ForegroundColor Cyan
    }
    
    # 初回表示
    Show-ConsolePage

    # HTML生成
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

        $bodyContent = @"
        <div class="card"><h2>Select Slide</h2><p>Select from list or press Start</p></div>
        <form method="post" action="/start"><button class="btn btn-start" $stBtn>Start: $nextTxt</button></form>
        $listHtml
        <form method="post" action="/exit"><button class="btn btn-exit">Exit System</button></form>
"@
    } else {
        $nxtLbl = if ($NextFileName) { "Start Next Slide<br><span style='font-size:0.8rem;font-weight:normal'>$NextFileName</span>" } else { "No slides in queue" }
        $nxtSt = if ($NextFileName) { "" } else { "disabled style='opacity:0.5;'" }

        $bodyContent = @"
        <div class="card"><h2>Presentation Ended</h2><p>$CurrentFileName</p></div>
        <form method="post" action="/next"><button class="btn btn-next" $nxtSt>$nxtLbl</button></form>
        <form method="post" action="/retry"><button class="btn btn-retry">Play Again</button></form>
        <form method="post" action="/lobby"><button class="btn btn-list">Back to List</button></form>
        <form method="post" action="/exit"><button class="btn btn-exit">Exit All</button></form>
"@
    }
    
    # ポーリング用スクリプト
    $pollingScript = @"
    <script>
        // ポーリングタイマーを変数に格納して制御可能にする
        var pollingTimer = setInterval(function(){
            fetch('/status?t=' + Date.now())
            .then(r => r.text())
            .then(status => {
                if (status === 'stopping') {
                    clearInterval(pollingTimer);
                    window.location.href = '/exit';
                } else if (status === 'changing' || status === 'starting' || status === 'running') {
                    // プレゼンテーション開始や状態変化時にリロード
                    // running も検知対象に追加（changing を取り逃がした場合の対処）
                    clearInterval(pollingTimer);
                    window.location.href = '/';
                }
            })
            .catch(e => console.log('Waiting connection...'));
        }, 300);  // 300msごとにチェック（800ms待機時間内に確実に検知）
        
        // フォーム送信時にポーリングを停止して画面遷移の競合を防止
        document.addEventListener('DOMContentLoaded', function() {
            var forms = document.querySelectorAll('form');
            forms.forEach(function(form) {
                form.addEventListener('submit', function() {
                    clearInterval(pollingTimer);
                });
            });
            
            // ボタンクリック時も念のため停止
            var buttons = document.querySelectorAll('.btn');
            buttons.forEach(function(btn) {
                btn.addEventListener('click', function() {
                    clearInterval(pollingTimer);
                });
            });
        });
    </script>
"@

    $mainHtml = $head + $bodyContent + $pollingScript + "</div></body></html>"

    # --- 画面状態のHTML ---
    $processingHtml = $head + @"
    <div style="margin-top:50px;"><div class="loader"></div><h2>Processing...</h2><p>Screen will refresh</p></div>
    <script>
        var checkCount = 0;
        var maxRetries = 60; // 最大30秒待機（500ms * 60）
        var errorCount = 0;
        var maxErrors = 40; // 接続エラー時は最大20秒待機（500ms * 40）
        var checkInterval = setInterval(function(){
            fetch('/status?t=' + Date.now())
            .then(r => r.text())
            .then(status => {
                // running（発表中）またはwaiting（待機中）になったらリロード
                if (status === 'running' || (status === 'waiting' && checkCount > 2)) {
                    clearInterval(checkInterval);
                    window.location.href = '/';
                } else {
                    checkCount++;
                    if (checkCount > maxRetries) {
                        // タイムアウト時は強制リロード
                        clearInterval(checkInterval);
                        window.location.href = '/';
                    }
                }
            })
            .catch(e => {
                // 接続エラー時、プレゼンテーション起動を待って再試行
                checkCount++;
                errorCount++;
                if (errorCount > maxErrors || checkCount > maxRetries) {
                    clearInterval(checkInterval);
                    window.location.href = '/';
                }
            });
        }, 500);
    </script>
</body></html>
"@
    
    $exitHtml = $head + @"
    <div style="margin-top:50px;">
        <div class="end-icon">✔</div>
        <h1>System Shutdown</h1>
        <p style="font-size:1.2rem; color:#fff;">Please close this tab<br>or window.</p>
        <p style="color:#666; margin-top:20px;">Server has been shut down.</p>
    </div>
</body></html>
"@

    $resultAction = $null
    $resultFile = $null
    $actionSetTime = $null  # アクション設定時刻を記録
    
    # シャットダウン制御用
    $shuttingDown = $false
    $shutdownDeadline = $null

    $contextTask = $listener.GetContextAsync()

    while ($true) {
        
        # --- Web確認 ---
        if ($contextTask.AsyncWaitHandle.WaitOne(100)) {
            $context = $contextTask.Result
            $req = $context.Request
            $res = $context.Response
            $url = $req.Url.LocalPath.ToLower()
            
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
                
                $contextTask = $listener.GetContextAsync()
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

            $contextTask = $listener.GetContextAsync()
        }

        # --- コンソール確認 ---
        if ((!$shuttingDown) -and ($resultAction -eq $null) -and [Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true).Key.ToString().ToUpper()
            
            # コンソールからの終了要求（5秒待機ロジック）
            if ($k -eq "Q" -or $k -eq "ESCAPE") {
                $shuttingDown = $true
                $shutdownDeadline = (Get-Date).AddSeconds(5)
                Write-Host ""
                Write-Host " [System] Shutting down... (Notifying web clients / Will exit in 5 seconds)" -ForegroundColor Magenta
            }

            if ($Mode -eq "Lobby") {
                if ($k -eq "ENTER" -or $k -eq "S") { $resultAction = "Start"; $actionSetTime = Get-Date }
                
                # ページング操作
                $totalActiveFiles = if ($ActiveFiles) { $ActiveFiles.Count } else { 0 }
                $totalFinishedFiles = if ($FinishedFiles) { $FinishedFiles.Count } else { 0 }
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

    $listener.Stop()
    $listener.Close()
    Start-Sleep -Milliseconds 200
    
    return @{ Action = $resultAction; FileName = $resultFile }
}

Add-Type -AssemblyName System.Web

# ------------------------------------------------------------
# 5. メインフロー
# ------------------------------------------------------------
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Administrator privileges required. Please run PowerShell as Administrator."
    Start-Sleep 3
    exit
}

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

    while (-not $exitLoop) {
        
        $activeFiles = Get-ChildItem -Path $TargetFolderPath -File | Where-Object { $_.Extension -in @('.ppt', '.pptx') -and $_.Name -notlike '~$*' } | Sort-Object Name
        $finishedFiles = Get-ChildItem -Path $finishFolderPath -File | Where-Object { $_.Extension -in @('.ppt', '.pptx') -and $_.Name -notlike '~$*' } | Sort-Object Name
        
        $targetFileItem = $null

        # --- A. 選択 ---
        if ($autoPlayTarget) {
            $targetFileItem = $autoPlayTarget
            $autoPlayTarget = $null
        } else {
            $result = Get-UserAction -Mode "Lobby" -ActiveFiles $activeFiles -FinishedFiles $finishedFiles
            
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
            $status = Watch-RunningPresentation -PptApp $pptApp -TargetFileItem $targetFileItem
            
            # 手動終了(ManualStop)の場合はここで閉じる
            if ($status -eq "ManualStop") {
                Write-Host " >> Manually stopped." -ForegroundColor Yellow
                try { $presentation.Close() } catch {}
            }
            
            # オブジェクト破棄
            $presentation = $null
            [GC]::Collect()

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

            $postResult = Get-UserAction -Mode "Dialog" -CurrentFileName $targetFileItem.Name -NextFileName $nextName

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
        }
    }

} finally {
    if ($pptApp) { try { $pptApp.Quit() } catch {}; Release-ComObject $pptApp }
    Write-Host "System terminated." -ForegroundColor Red
}