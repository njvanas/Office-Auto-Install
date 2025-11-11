@echo off
REM Office Auto Installer - GUI Launcher
REM Simple launcher for the single-file GUI application

title Microsoft Office Auto Installer - Launcher

echo.
echo ================================================================
echo   MICROSOFT OFFICE AUTO INSTALLER
echo   Modern GUI Edition - Single File Application
echo ================================================================
echo.
echo Launching the Office installer GUI...
echo.

REM Check for admin and launch
net session >nul 2>&1
if %errorLevel% == 0 (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Office-GUI-WPF.ps1"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%~dp0Install-Office-GUI-WPF.ps1\"' -Verb RunAs"
)
