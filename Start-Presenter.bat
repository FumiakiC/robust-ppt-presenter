@echo off
cd /d %~dp0

REM ===============================================
REM Interactive presentation web remote: prep + launcher
REM  - URLACL registration (http.sys)
REM  - Windows Defender Firewall <WEB_PORT>/TCP allow (includes other subnets)
REM  - Admin elevation
REM  - Launch PowerShell script
REM ===============================================

REM Set code page to UTF-8 (65001)
chcp 65001 > nul

setlocal

REM -----------------------------------------------
REM Settings
REM -----------------------------------------------
set "SCRIPT_NAME=Invoke-PPTController.ps1"
set "WEB_PORT=8090"
set "FW_RULE_NAME=PresentController TCP %WEB_PORT% In"
REM Reserve URL for all NICs (use http://<IP>:%WEB_PORT%/ to bind a fixed IP)
set "URLACL_URL=http://+:%WEB_PORT%/"

REM PowerShell command
set "POWERSHELL=powershell.exe"

REM -----------------------------------------------
REM Ensure PowerShell script exists
REM -----------------------------------------------
if not exist "%SCRIPT_NAME%" (
    echo [Error] File not found: %SCRIPT_NAME%
    echo Please place the PowerShell script in the same folder as this batch file.
    pause
    exit /b 1
)

REM -----------------------------------------------
REM Admin check and elevation (net session succeeds only as admin)
REM -----------------------------------------------
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Restarting batch with administrator privileges...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo.
echo ==== Starting Pre-processing: Configuring URLACL and Firewall ====
echo.

REM -----------------------------------------------
REM URLACL: delete existing reservation and re-add
REM  - Recreate each time to avoid conflicts/inconsistencies
REM -----------------------------------------------
echo [URLACL] Removing: %URLACL_URL% (OK if it doesn't exist)
netsh http delete urlacl url=%URLACL_URL% >nul 2>&1

echo [URLACL] Adding: %URLACL_URL%
REM URLACL user: prefer UPN (whoami /upn), fallback to DOMAIN\USERNAME
set "CURRENT_UPN="
for /f "tokens=1,* delims=:" %%A in ('whoami /upn 2^>nul ^| find ":"') do set "CURRENT_UPN=%%B"
if defined CURRENT_UPN (
    for /f "tokens=* delims= " %%Z in ("%CURRENT_UPN%") do set "CURRENT_UPN=%%Z"
) else (
    set "CURRENT_UPN=%USERDOMAIN%\%USERNAME%"
)

netsh http add urlacl url=%URLACL_URL% user="%CURRENT_UPN%" listen=yes
if %errorlevel% neq 0 (
    echo [Warning] Failed to add URLACL. Continuing, but please check permissions later.
)

REM -----------------------------------------------
REM Windows Defender Firewall: allow inbound <WEB_PORT>/TCP
REM  - RemoteAddress=Any, Profile=Any to allow other subnets
REM  - Remove same-name rule before adding
REM -----------------------------------------------
echo [FW] Removing existing rule (same name): %FW_RULE_NAME%
"%POWERSHELL%" -NoProfile -Command ^
  "Get-NetFirewallRule -DisplayName '%FW_RULE_NAME%' -ErrorAction SilentlyContinue | Remove-NetFirewallRule -ErrorAction SilentlyContinue" >nul 2>&1

echo [FW] Adding new rule: %FW_RULE_NAME%
"%POWERSHELL%" -NoProfile -Command ^
  "New-NetFirewallRule -DisplayName '%FW_RULE_NAME%' -Direction Inbound -Action Allow -Protocol TCP -LocalPort %WEB_PORT% -RemoteAddress Any -Profile Any" >nul 2>&1

if %errorlevel% neq 0 (
    echo [Warning] Error creating Firewall rule. It may be restricted by GPO or other policies.
)

REM -----------------------------------------------
REM Reference: current LISTEN status (may be empty before listener starts)
REM -----------------------------------------------
echo.
echo [INFO] Current binding status for port %WEB_PORT% (reference before start):
netstat -ano | findstr ":%WEB_PORT%" || echo   (None found)
echo.

REM -----------------------------------------------
REM Start PowerShell script (execution policy bypass)
REM  - Batch already runs as admin, no additional RunAs needed
REM -----------------------------------------------
echo Starting PowerShell script with administrator privileges...
echo.

"%POWERSHELL%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0%SCRIPT_NAME%"

REM -----------------------------------------------
REM Cleanup: Remove URLACL and Firewall rule
REM -----------------------------------------------
echo.
echo ==== Starting Post-processing: Cleaning up system configuration ====
echo.
echo [Cleanup] Removing URLACL: %URLACL_URL%
netsh http delete urlacl url=%URLACL_URL% >nul 2>&1
if %errorlevel% equ 0 (
    echo [Cleanup] URLACL successfully removed.
) else (
    echo [Cleanup] Warning: Could not remove URLACL [may not have existed].
)
echo.
echo [Cleanup] Removing Firewall rule: %FW_RULE_NAME%
"%POWERSHELL%" -NoProfile -Command "Remove-NetFirewallRule -DisplayName '%FW_RULE_NAME%' -ErrorAction SilentlyContinue" >nul 2>&1

if %errorlevel% equ 0 (
    echo [Cleanup] Firewall rule successfully removed.
) else (
    echo [Cleanup] Warning: Could not remove Firewall rule [may not have existed].
)
echo.
echo ==== Post-processing completed ====
echo.

REM -----------------------------------------------
REM End of batch
REM -----------------------------------------------
echo All processes completed. Press any key to close this window...
pause >nul
exit /b 0
``