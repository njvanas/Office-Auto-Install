@echo off
REM Office Auto Installer - Batch File Launcher
REM This batch file bypasses PowerShell execution policy issues

echo.
echo ================================================================
echo   MICROSOFT OFFICE AUTO INSTALLER - SAFE LAUNCHER
echo ================================================================
echo.
echo This launcher will run the PowerShell script safely, bypassing
echo execution policy restrictions that prevent the script from running.
echo.
echo What this does:
echo   1. Requests administrator privileges (you'll see a UAC prompt)
echo   2. Runs the PowerShell script with execution policy bypass
echo   3. Keeps the window open so you can see results
echo.
echo This is completely safe - it just allows the script to run!
echo.
pause

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with administrator privileges...
    goto :run_script
) else (
    echo Requesting administrator privileges...
    echo You'll see a Windows security prompt - click "Yes" to continue.
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:run_script
echo.
echo ================================================================
echo   LAUNCHING OFFICE INSTALLER...
echo ================================================================
echo.

REM Run the PowerShell script with execution policy bypass
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Office.ps1"

echo.
echo ================================================================
echo   SCRIPT FINISHED
echo ================================================================
echo.
echo The Office installer has finished running.
echo Check the messages above to see if installation was successful.
echo.
pause