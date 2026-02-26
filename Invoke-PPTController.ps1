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
<html lang="ja">
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
        Write-Warning "ポート競合等のためWeb制御(中断)ができません。キーボード操作のみ有効です。"
    }

    $head = Get-HtmlHeader -Title "Now Playing" -BgColor "#000000"
    $bodyHtml = @"
    <div class="card" style="border: 1px solid #28a745;">
        <div class="playing-icon">▶</div>
        <h2>発表中</h2>
        <p style="font-weight:bold; color:#fff;">$($TargetFileItem.Name)</p>
        <p>PC側でスライド操作中...</p>
    </div>
    <form method="post" action="/stop">
        <button class="btn btn-stop">■ 発表を終了する (Finish)</button>
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

    # 画面表示
    Clear-Host
    $ip = Get-LocalIPAddress
    $line = "=" * 70
    Write-Host $line -ForegroundColor Cyan
    Write-Host "   プレゼンテーション コントローラー V7.3" -ForegroundColor White -BackgroundColor DarkCyan
    Write-Host $line -ForegroundColor Cyan
    Write-Host " [Web URL]  http://$($ip):$($WebPort)/" -ForegroundColor Yellow
    Write-Host " [状態]     $Mode" -ForegroundColor White
    Write-Host ""
    Write-Host " --- PC操作メニュー ---" -ForegroundColor Gray
    if ($Mode -eq "Lobby") {
        Write-Host " [Enter] 開始 (Start)" -ForegroundColor Green
        Write-Host " [Q]     システム終了 (Exit)" -ForegroundColor Red
    } else {
        Write-Host " [Enter] 次へ (Next)" -ForegroundColor Green
        Write-Host " [R]     リトライ (Retry)" -ForegroundColor Yellow
        Write-Host " [L]     リストへ (Lobby)" -ForegroundColor Cyan
        Write-Host " [Q]     システム終了 (Exit)" -ForegroundColor Red
    }
    Write-Host $line -ForegroundColor Cyan

    # HTML生成
    $head = Get-HtmlHeader -Title "Controller" -BgColor $(if($Mode -eq 'Lobby'){"#1a1a1a"}else{"#000000"})
    
    # Bodyコンテンツ
    $bodyContent = ""
    if ($Mode -eq 'Lobby') {
        $nextTxt = if ($ActiveFiles) { $ActiveFiles[0].Name } else { "なし" }
        $stBtn = if ($ActiveFiles) { "" } else { "disabled style='opacity:0.5;'" }
        
        $listHtml = "<div class='list-container'><h3 style='color:#28a745;font-size:0.9rem;border-bottom:1px solid #555;'>待機中</h3>"
        if (!$ActiveFiles) { $listHtml += "<p>なし</p>" }
        foreach ($f in $ActiveFiles) {
            $listHtml += "<form method='post' action='/select' style='margin:0;'><input type='hidden' name='filename' value='$($f.Name)'><button class='btn btn-file'>$($f.Name)</button></form>"
        }
        $listHtml += "<h3 style='color:#6c757d;font-size:0.9rem;border-bottom:1px solid #555;margin-top:20px;'>完了済み</h3>"
        foreach ($f in $FinishedFiles) {
            $listHtml += "<form method='post' action='/select' style='margin:0;'><input type='hidden' name='filename' value='$($f.Name)'><button class='btn btn-file btn-finished'>$($f.Name)</button></form>"
        }
        $listHtml += "</div>"

        $bodyContent = @"
        <div class="card"><h2>スライド選択</h2><p>リスト選択 または 開始ボタン</p></div>
        <form method="post" action="/start"><button class="btn btn-start" $stBtn>開始: $nextTxt</button></form>
        $listHtml
        <form method="post" action="/exit"><button class="btn btn-exit">システム終了</button></form>
"@
    } else {
        $nxtLbl = if ($NextFileName) { "次のスライドを開始<br><span style='font-size:0.8rem;font-weight:normal'>$NextFileName</span>" } else { "待機リストなし" }
        $nxtSt = if ($NextFileName) { "" } else { "disabled style='opacity:0.5;'" }

        $bodyContent = @"
        <div class="card"><h2>発表終了</h2><p>$CurrentFileName</p></div>
        <form method="post" action="/next"><button class="btn btn-next" $nxtSt>$nxtLbl</button></form>
        <form method="post" action="/retry"><button class="btn btn-retry">もう一度再生</button></form>
        <form method="post" action="/lobby"><button class="btn btn-list">リストに戻る</button></form>
        <form method="post" action="/exit"><button class="btn btn-exit">全て終了</button></form>
"@
    }
    
    # ポーリング用スクリプト
    $pollingScript = @"
    <script>
        setInterval(function(){
            fetch('/status?t=' + Date.now())
            .then(r => r.text())
            .then(status => {
                if (status === 'stopping') {
                    window.location.href = '/exit';
                }
            })
            .catch(e => console.log('Waiting connection...'));
        }, 1000);
    </script>
"@

    $mainHtml = $head + $bodyContent + $pollingScript + "</div></body></html>"

    # --- 画面状態のHTML ---
    $processingHtml = $head + @"
    <div style="margin-top:50px;"><div class="loader"></div><h2>処理中...</h2><p>画面が切り替わります</p></div>
    <script>setTimeout(function(){ window.location.href='/'; }, 1000);</script>
</body></html>
"@
    
    $exitHtml = $head + @"
    <div style="margin-top:50px;">
        <div class="end-icon">✔</div>
        <h1>システム終了</h1>
        <p style="font-size:1.2rem; color:#fff;">このタブまたはウィンドウを<br>閉じてください。</p>
        <p style="color:#666; margin-top:20px;">Server has been shut down.</p>
    </div>
</body></html>
"@

    $resultAction = $null
    $resultFile = $null
    
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
                $statusText = if ($shuttingDown) { "stopping" } else { "waiting" }
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
                    "/start"  { $resultAction = "Start"; $resHtml = $processingHtml }
                    "/next"   { $resultAction = "Next";  $resHtml = $processingHtml }
                    "/retry"  { $resultAction = "Retry"; $resHtml = $processingHtml }
                    "/lobby"  { $resultAction = "Lobby"; $resHtml = $processingHtml }
                    "/exit"   { 
                        $resultAction = "Exit"
                        $shuttingDown = $true 
                        $resHtml = $exitHtml 
                    }
                    "/select" {
                        if ([System.Web.HttpUtility]::UrlDecode($body) -match "filename=(.*)") { 
                            $resultAction = "Select"; $resultFile = $matches[1] 
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
                Write-Host " [System] 終了準備中... (Web画面へ通知中 / 5秒後に終了します)" -ForegroundColor Magenta
            }

            if ($Mode -eq "Lobby") {
                if ($k -eq "ENTER" -or $k -eq "S") { $resultAction = "Start" }
            } else {
                if ($k -eq "ENTER" -or $k -eq "N") { $resultAction = "Next" }
                if ($k -eq "R") { $resultAction = "Retry" }
                if ($k -eq "L" -or $k -eq "BACKSPACE") { $resultAction = "Lobby" }
            }
        }

        # --- 終了判定 ---
        if ($resultAction -ne $null -and $resultAction -ne "Exit") {
            break
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
    Write-Warning "管理者権限が必要です。PowerShellを管理者として実行してください。"
    Start-Sleep 3
    exit
}

if (-not (Test-Path $TargetFolderPath)) { Write-Error "Target Folder Not Found"; exit }
$finishFolderPath = Join-Path $TargetFolderPath $FinishFolderName
if (-not (Test-Path $finishFolderPath)) { New-Item -Path $finishFolderPath -ItemType Directory | Out-Null }

Write-Host "PowerPointを起動しています..." -ForegroundColor Cyan
try {
    $pptApp = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
} catch {
    Write-Error "PowerPoint起動失敗"
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
            Write-Host " >> 開いています: $($targetFileItem.Name)" -ForegroundColor Cyan
            $presentation = $pptApp.Presentations.Open($targetFileItem.FullName, $false, $false, $true)
            $presentation.SlideShowSettings.Run() | Out-Null
            
            # 発表中の監視
            $status = Watch-RunningPresentation -PptApp $pptApp -TargetFileItem $targetFileItem
            
            # 手動終了(ManualStop)の場合はここで閉じる
            if ($status -eq "ManualStop") {
                Write-Host " >> 手動で終了されました。" -ForegroundColor Yellow
                try { $presentation.Close() } catch {}
            }
            
            # オブジェクト破棄
            $presentation = $null
            [GC]::Collect()

            # --- C. 移動判定 ---
            if ($targetFileItem.DirectoryName -ne $finishFolderPath) {
                try {
                    Write-Host " >> 完了フォルダへ移動..." -ForegroundColor Gray
                    $targetFileItem = Move-Item -LiteralPath $targetFileItem.FullName -Destination $finishFolderPath -Force -PassThru
                } catch { Write-Warning "移動失敗: $_" }
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
            Write-Error "エラー: $($_.Exception.Message)"
            if ($presentation) { try { $presentation.Close() } catch {} }
            Start-Sleep 2
        }
    }

} finally {
    if ($pptApp) { try { $pptApp.Quit() } catch {}; Release-ComObject $pptApp }
    Write-Host "システムを終了しました。" -ForegroundColor Red
}