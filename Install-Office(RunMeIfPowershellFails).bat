@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

:: Enhanced batch file for users who have PowerShell execution policy issues
echo.
echo ===============================================================================
echo                    MICROSOFT OFFICE AUTO INSTALLER
echo                         Batch File Launcher
echo ===============================================================================
echo.
echo This batch file will help you run the Office installer if PowerShell 
echo execution policies are preventing the script from running normally.
echo.
echo What this does:
echo   - Bypasses PowerShell execution policy restrictions
echo   - Automatically requests administrator privileges
echo   - Launches the main Office installer script
echo.
echo This is completely safe and standard for PowerShell script execution.
echo.

:: Check if we're running as administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo [OK] Running with administrator privileges
    echo.
) else (
    echo [WARNING] Not running as administrator
    echo.
    echo Administrator privileges are required for Office installation.
    echo This batch file will request elevation automatically.
    echo.
    pause
)

:: Locate the PowerShell script
SET SCRIPT_DIR=%~dp0
SET SCRIPT_NAME=Install-Office.ps1
SET FULL_PATH="%SCRIPT_DIR%%SCRIPT_NAME%"

echo Launching Office installer...
echo Script location: %FULL_PATH%
echo.

:: Check if the PowerShell script exists
if not exist %FULL_PATH% (
    echo ERROR: PowerShell script not found!
    echo Expected location: %FULL_PATH%
    echo.
    echo Please make sure both files are in the same folder:
    echo   - Install-Office.ps1
    echo   - Install-Office(RunMeIfPowershellFails).bat
    echo.
    pause
    exit /b 1
)

echo Starting PowerShell with elevated privileges and bypassed execution policy...
echo.

:: Run the PowerShell script with elevation and execution policy bypass
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File %FULL_PATH%' -Verb RunAs"

if %errorLevel% == 0 (
    echo.
    echo [SUCCESS] PowerShell script launched successfully!
    echo.
    echo A new PowerShell window should have opened with the Office installer.
    echo You can close this window now.
    echo.
    echo If no new window appeared, try:
    echo   1. Right-click on Install-Office.ps1
    echo   2. Select "Run with PowerShell"
    echo   3. Click "Yes" when Windows asks for administrator permission
    echo.
) else (
    echo.
    echo [ERROR] Failed to launch PowerShell script!
    echo.
    echo This might happen if:
    echo   - PowerShell is not installed or corrupted
    echo   - Windows security settings are very restrictive
    echo   - User Account Control (UAC) is disabled
    echo.
    echo Manual solution:
    echo   1. Right-click on Install-Office.ps1
    echo   2. Select "Run with PowerShell"
    echo   3. Click "Yes" when prompted for administrator access
    echo.
)

echo.
echo Press any key to exit...
pause >nul