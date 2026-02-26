@echo off
cd /d %~dp0

:: ===============================================
:: 対話型プレゼン Webリモコン 前処理 + 起動バッチ（完全版）
::  - URLACL 登録（http.sys）
::  - Windows Defender Firewall <WEB_PORT>/TCP 許可（異サブネット含む）
::  - 管理者昇格
::  - 64bit PowerShell で PowerShell スクリプト起動
:: ===============================================

:: 文字コードをUTF-8 (65001) に
chcp 65001 > nul

setlocal

:: -----------------------------------------------
:: 設定
:: -----------------------------------------------
set "SCRIPT_NAME=Invoke-PPTController.ps1"
set "WEB_PORT=8090"
set "FW_RULE_NAME=PresentController TCP %WEB_PORT% In"
:: すべてのNICで待受するURL予約（実IPに固定したい場合は http://<IP>:%WEB_PORT%/ に変更）
set "URLACL_URL=http://+:%WEB_PORT%/"

:: 64bit PowerShell のフルパス（Office 64bit と揃えるために推奨）
set "POWERSHELL_X64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

:: -----------------------------------------------
:: PowerShellスクリプトの存在確認
:: -----------------------------------------------
if not exist "%SCRIPT_NAME%" (
    echo [エラー] ファイルが見つかりません: %SCRIPT_NAME%
    echo このバッチと同じフォルダに PowerShell スクリプトを配置してください。
    pause
    exit /b 1
)

:: -----------------------------------------------
:: 管理者権限チェック＆昇格（net session は管理者でのみ成功）
:: -----------------------------------------------
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 管理者権限でバッチを再起動します...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

:: -----------------------------------------------
:: 64bit PowerShell が見つからない場合のフォールバック
:: -----------------------------------------------
if not exist "%POWERSHELL_X64%" (
    echo [警告] 64bit PowerShell が見つかりません。既定の powershell.exe を使用します。
    set "POWERSHELL_X64=powershell.exe"
)

echo.
echo ==== 前処理開始: URLACL と Firewall の設定を行います ====
echo.

:: -----------------------------------------------
:: URLACL: 既存の URL 予約を削除 → 再登録
::  - 競合や不整合を避けるため、毎回クリーンに作り直します
:: -----------------------------------------------
echo [URLACL] 削除: %URLACL_URL%（存在しなくてもOK）
netsh http delete urlacl url=%URLACL_URL% >nul 2>&1

echo [URLACL] 追加: %URLACL_URL%
:: URLACL の user は UPN（whoami /upn）を優先、取れない場合は ドメイン\ユーザー を使用
set "CURRENT_UPN="
for /f "tokens=1,* delims=:" %%A in ('whoami /upn 2^>nul ^| find ":"') do set "CURRENT_UPN=%%B"
if defined CURRENT_UPN (
    for /f "tokens=* delims= " %%Z in ("%CURRENT_UPN%") do set "CURRENT_UPN=%%Z"
) else (
    set "CURRENT_UPN=%USERDOMAIN%\%USERNAME%"
)

netsh http add urlacl url=%URLACL_URL% user="%CURRENT_UPN%" listen=yes
if %errorlevel% neq 0 (
    echo [警告] URLACL の追加に失敗しました。続行しますが、後で権限を確認してください。
)

:: -----------------------------------------------
:: Windows Defender Firewall: <WEB_PORT>/TCP 着信許可
::  - RemoteAddress=Any, Profile=Any で「異サブネットからも許可」
::  - 既存の同名ルールは削除してから追加
:: -----------------------------------------------
echo [FW] 既存ルール削除（同名）: %FW_RULE_NAME%
"%POWERSHELL_X64%" -NoProfile -Command ^
  "Get-NetFirewallRule -DisplayName '%FW_RULE_NAME%' -ErrorAction SilentlyContinue | Remove-NetFirewallRule -ErrorAction SilentlyContinue" >nul 2>&1

echo [FW] 新規ルール追加: %FW_RULE_NAME%
"%POWERSHELL_X64%" -NoProfile -Command ^
  "New-NetFirewallRule -DisplayName '%FW_RULE_NAME%' -Direction Inbound -Action Allow -Protocol TCP -LocalPort %WEB_PORT% -RemoteAddress Any -Profile Any" >nul 2>&1

if %errorlevel% neq 0 (
    echo [警告] Firewall ルールの作成でエラーが発生しました。GPOなどで制限されている可能性があります。
)

:: -----------------------------------------------
:: （参考情報）現時点の LISTEN 状況を表示
::  - PowerShell側のWebリスナー起動前なので、空でも正常
:: -----------------------------------------------
echo.
echo [INFO] ポート %WEB_PORT% の現行バインド状況（開始前の参考情報）:
netstat -ano | findstr ":%WEB_PORT%" || echo   (該当なし)
echo.

:: -----------------------------------------------
:: PowerShell スクリプトを 64bit で起動（実行ポリシー回避）
::  - 既にこのバッチ自体が管理者で動作中のため、改めて RunAs は不要
:: -----------------------------------------------
echo 管理者権限で PowerShell スクリプト（64bit）を起動します...
echo.

"%POWERSHELL_X64%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0%SCRIPT_NAME%"

:: -----------------------------------------------
:: バッチはここで終了
:: -----------------------------------------------
echo.
echo すべての処理が完了しました。ウィンドウを閉じてもOKです。
exit /b 0
``