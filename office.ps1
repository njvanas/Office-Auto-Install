#Requires -Version 5.1
<#
.SYNOPSIS
    Office Auto Install - remote bootstrap (standard PowerShell: fetch script, run with Invoke-Expression).

.DESCRIPTION
    Downloads the latest Install-Office-GUI-WPF.ps1 or Install-Office.ps1 from GitHub and runs it in this
    PowerShell session. The bootstrap is published on GitHub Pages for a shorter URL; raw.githubusercontent.com
    is the fallback host for the same file.

.EXAMPLE
    # Stable — GUI (recommended). Run in Windows Terminal or PowerShell as Administrator:
    irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex

.EXAMPLE
    # Stable — raw GitHub (if Pages is unreachable):
    irm "https://raw.githubusercontent.com/njvanas/Office-Auto-Install/main/office.ps1" | iex

.EXAMPLE
    # Dev branch payloads (bootstrap still from main; set branch before irm):
    $env:OFFICE_AUTO_INSTALL_BRANCH = "dev"
    irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex

.EXAMPLE
    # Console wizard instead of GUI:
    $env:OFFICE_AUTO_INSTALL_USE_CONSOLE = "1"
    irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex

.NOTES
    Forks: set OFFICE_AUTO_INSTALL_REPO (e.g. "yourname/Office-Auto-Install"). Optional: OFFICE_AUTO_INSTALL_BRANCH.
    Only run irm ... | iex from sources you trust. Administrator rights are required for Office setup.
#>

$ErrorActionPreference = "Stop"

# Default upstream (change in your fork if you publish your own copy of this file)
$DefaultRepo = "njvanas/Office-Auto-Install"
$DefaultBranch = "main"

if (-not ([Net.ServicePointManager]::SecurityProtocol -band [Net.SecurityProtocolType]::Tls12)) {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch { }
}

$repo = if ($env:OFFICE_AUTO_INSTALL_REPO) { $env:OFFICE_AUTO_INSTALL_REPO.Trim() } else { $DefaultRepo }
$branch = if ($env:OFFICE_AUTO_INSTALL_BRANCH) { $env:OFFICE_AUTO_INSTALL_BRANCH.Trim() } else { $DefaultBranch }
$useConsole = ($env:OFFICE_AUTO_INSTALL_USE_CONSOLE -eq "1")
$scriptFile = if ($useConsole) { "Install-Office.ps1" } else { "Install-Office-GUI-WPF.ps1" }
$uri = "https://raw.githubusercontent.com/$repo/$branch/$scriptFile"

Write-Host ""
Write-Host " =================================================================" -ForegroundColor Cyan
Write-Host "  Microsoft Office Auto Installer" -ForegroundColor White
Write-Host "  Mode: $(if ($useConsole) { 'Console' } else { 'GUI (WPF)' })" -ForegroundColor Gray
Write-Host "  Payload: $uri" -ForegroundColor DarkGray
Write-Host " =================================================================" -ForegroundColor Cyan
Write-Host ""

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host " Administrator rights are required for Office setup." -ForegroundColor Yellow
    Write-Host " Open Terminal or PowerShell as Administrator, then run this command again." -ForegroundColor Yellow
    Write-Host ""
}

try {
    $response = Invoke-WebRequest -Uri $uri -UseBasicParsing -MaximumRedirection 5 -TimeoutSec 120
    $body = $response.Content
} catch {
    Write-Host " ERROR: Could not download the installer from GitHub." -ForegroundColor Red
    Write-Host " $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host " Check repo/branch (OFFICE_AUTO_INSTALL_REPO / OFFICE_AUTO_INSTALL_BRANCH) or your network." -ForegroundColor Gray
    if ($Host.Name -eq "ConsoleHost") {
        try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { Read-Host "Press Enter to exit" }
    }
    exit 1
}

if ([string]::IsNullOrWhiteSpace($body)) {
    Write-Host " ERROR: Download returned an empty file." -ForegroundColor Red
    exit 1
}

Write-Host " Starting installer..." -ForegroundColor Green
Write-Host ""

Invoke-Expression $body
