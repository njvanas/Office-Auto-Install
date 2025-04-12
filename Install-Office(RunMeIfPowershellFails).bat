@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

:: Locate script
SET SCRIPT_DIR=%~dp0
SET SCRIPT_NAME=Install-Office.ps1
SET FULL_PATH="%SCRIPT_DIR%%SCRIPT_NAME%"

:: Run the PowerShell script with elevation
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File %FULL_PATH%' -Verb RunAs"

pause
