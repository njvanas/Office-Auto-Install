#Requires -Version 5.1
<#
.SYNOPSIS
    Single-command bootstrap: downloads the right scripts from GitHub and runs the Microsoft 365 Apps installer.

.DESCRIPTION
    One liner for everyone:
    - Default: graphical installer for home and business users (same experience as always).
    - Optional: text-only wizard, or IT silent deploy using presets from configs\ (AVD, physical, uninstall).
      Office Deployment Tool (setup.exe) is fetched from Microsoft's official CDN when install runs — not stored in git.

    Set OFFICE_AUTO_INSTALL_MODE before irm | iex:
      gui     — WPF installer (default)
      console — prompts in the terminal
      deploy  — runs Deploy-Microsoft365Apps.ps1 with presets from configs\ (downloads full toolkit)

    Fork: OFFICE_AUTO_INSTALL_REPO, OFFICE_AUTO_INSTALL_BRANCH

.EXAMPLE
    # End user — graphical installer (default)
    irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex

.EXAMPLE
    # Same as default; explicit mode
    $env:OFFICE_AUTO_INSTALL_MODE = "gui"
    irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex

.EXAMPLE
    # Text-only wizard
    $env:OFFICE_AUTO_INSTALL_MODE = "console"
    irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex

.EXAMPLE
    # IT — AVD / shared PC preset (silent ODT deploy; run PowerShell as Administrator)
    $env:OFFICE_AUTO_INSTALL_MODE = "deploy"
    $env:OFFICE_AUTO_INSTALL_PRESET = "O365ProPlus-VDI"
    $env:OFFICE_AUTO_INSTALL_LANGUAGE = "en-us"
    irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex

.NOTES
    Legacy: OFFICE_AUTO_INSTALL_USE_CONSOLE=1 still selects console mode if MODE is not set.
#>

$ErrorActionPreference = "Stop"

$DefaultRepo = "njvanas/Office-Auto-Install"
$DefaultBranch = "main"

# Bundled preset XML names under configs/ on GitHub — keep in sync with the repository
$script:OaiConfigFiles = @(
    'O365ProPlus.xml',
    'O365ProPlus-VDI.xml',
    'O365Business.xml',
    'O365Business-VDI.xml',
    'O365ProPlusVisioProject.xml',
    'O365ProPlusVisioProject-VDI.xml',
    'Uninstall-Microsoft365Apps.xml'
)

if (-not ([Net.ServicePointManager]::SecurityProtocol -band [Net.SecurityProtocolType]::Tls12)) {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch { }
}

$repo = if ($env:OFFICE_AUTO_INSTALL_REPO) { $env:OFFICE_AUTO_INSTALL_REPO.Trim() } else { $DefaultRepo }
$branch = if ($env:OFFICE_AUTO_INSTALL_BRANCH) { $env:OFFICE_AUTO_INSTALL_BRANCH.Trim() } else { $DefaultBranch }
$base = "https://raw.githubusercontent.com/$repo/$branch"

# Mode: gui | console | deploy (backward compat: USE_CONSOLE=1 → console)
$mode = if ($env:OFFICE_AUTO_INSTALL_MODE) {
    $env:OFFICE_AUTO_INSTALL_MODE.Trim().ToLowerInvariant()
} elseif ($env:OFFICE_AUTO_INSTALL_USE_CONSOLE -eq "1") {
    'console'
} else {
    'gui'
}

if ($mode -notin @('gui', 'console', 'deploy')) {
    Write-Host " ERROR: OFFICE_AUTO_INSTALL_MODE must be gui, console, or deploy (got '$mode')." -ForegroundColor Red
    exit 1
}

$scriptFile = switch ($mode) {
    'gui' { 'Install-Office-GUI-WPF.ps1' }
    'console' { 'Install-Office.ps1' }
    'deploy' { 'Deploy-Microsoft365Apps.ps1' }
}

$safeRepo = ($repo -replace '/', '_')
$stage = Join-Path $env:TEMP "OfficeAutoInstall\$safeRepo\$branch"
New-Item -ItemType Directory -Path $stage -Force | Out-Null

$modulePath = Join-Path $stage "Microsoft365AppsDeployment.psm1"
$payloadPath = Join-Path $stage $scriptFile

Write-Host ""
Write-Host " =================================================================" -ForegroundColor Cyan
Write-Host "  Microsoft 365 Apps - Office Auto Install" -ForegroundColor White
if ($mode -eq 'gui') {
    Write-Host "  Profile: Standard (graphical installer)" -ForegroundColor Gray
    Write-Host "  Tip: Run PowerShell as Administrator when prompted." -ForegroundColor DarkGray
} elseif ($mode -eq 'console') {
    Write-Host "  Profile: Step-by-step (terminal)" -ForegroundColor Gray
} else {
    Write-Host "  Profile: IT - silent deploy (ODT + presets)" -ForegroundColor Yellow
    Write-Host "  Requires: Administrator PowerShell" -ForegroundColor DarkGray
}
Write-Host "  Source: $repo @ $branch" -ForegroundColor DarkGray
Write-Host " =================================================================" -ForegroundColor Cyan
Write-Host ""

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if ($mode -eq 'deploy') {
    if (-not $isAdmin) {
        Write-Host " Deploy mode must run elevated. Open PowerShell as Administrator and run the same command." -ForegroundColor Red
        Write-Host ""
        if ($Host.Name -eq "ConsoleHost") {
            try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { Read-Host "Press Enter to exit" }
        }
        exit 1
    }
} elseif (-not $isAdmin) {
    Write-Host " Administrator rights are required to install Office." -ForegroundColor Yellow
    Write-Host " Open Terminal or PowerShell as Administrator, then run this command again." -ForegroundColor Yellow
    Write-Host ""
}

function Get-OaiWebFile {
    param([string]$Uri, [string]$OutFile)
    Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -MaximumRedirection 5 -TimeoutSec 120
}

try {
    Get-OaiWebFile -Uri "$base/Microsoft365AppsDeployment.psm1" -OutFile $modulePath

    if ($mode -in 'gui', 'console') {
        Get-OaiWebFile -Uri "$base/$scriptFile" -OutFile $payloadPath
        $configsDir = Join-Path $stage 'configs'
        New-Item -ItemType Directory -Path $configsDir -Force | Out-Null
        foreach ($f in $script:OaiConfigFiles) {
            Get-OaiWebFile -Uri "$base/configs/$f" -OutFile (Join-Path $configsDir $f)
        }
    } else {
        $configsDir = Join-Path $stage 'configs'
        New-Item -ItemType Directory -Path $configsDir -Force | Out-Null
        foreach ($f in $script:OaiConfigFiles) {
            Get-OaiWebFile -Uri "$base/configs/$f" -OutFile (Join-Path $configsDir $f)
        }
        Get-OaiWebFile -Uri "$base/Deploy-Microsoft365Apps.ps1" -OutFile $payloadPath
    }
} catch {
    Write-Host " ERROR: Download from GitHub failed." -ForegroundColor Red
    Write-Host " $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host " Check repo/branch (OFFICE_AUTO_INSTALL_REPO / OFFICE_AUTO_INSTALL_BRANCH) and network." -ForegroundColor Gray
    if ($Host.Name -eq "ConsoleHost") {
        try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { Read-Host "Press Enter to exit" }
    }
    exit 1
}

if (-not (Test-Path -LiteralPath $modulePath) -or -not (Test-Path -LiteralPath $payloadPath)) {
    Write-Host " ERROR: Required files are missing after download." -ForegroundColor Red
    exit 1
}

if ($mode -in 'gui', 'console', 'deploy') {
    foreach ($f in $script:OaiConfigFiles) {
        $p = Join-Path $stage "configs\$f"
        if (-not (Test-Path -LiteralPath $p)) {
            Write-Host " ERROR: Missing config: $f" -ForegroundColor Red
            exit 1
        }
    }
}

Write-Host " Files ready. Starting..." -ForegroundColor Green
Write-Host ""

if ($mode -in 'gui', 'console') {
    & $payloadPath
    exit $LASTEXITCODE
}

# --- Deploy mode: map environment to Deploy-Microsoft365Apps.ps1 parameters ---
$uninstall = ($env:OFFICE_AUTO_INSTALL_UNINSTALL -eq '1')
$preset = if ($env:OFFICE_AUTO_INSTALL_PRESET) { $env:OFFICE_AUTO_INSTALL_PRESET.Trim() } else { 'O365ProPlus' }
$lang = if ($env:OFFICE_AUTO_INSTALL_LANGUAGE) { $env:OFFICE_AUTO_INSTALL_LANGUAGE.Trim() } elseif ($env:OFFICE_AUTO_INSTALL_LANGUAGEID) { $env:OFFICE_AUTO_INSTALL_LANGUAGEID.Trim() } else { 'en-us' }

$deployArgs = @{
    LanguageId = $lang
}
if ($env:OFFICE_AUTO_INSTALL_CHANNEL) {
    $deployArgs.Channel = $env:OFFICE_AUTO_INSTALL_CHANNEL.Trim()
}
if ($env:OFFICE_AUTO_INSTALL_ARCH -match '^(32|64)$') {
    $deployArgs.OfficeClientEdition = $env:OFFICE_AUTO_INSTALL_ARCH
}
if ($env:OFFICE_AUTO_INSTALL_WORKING_DIRECTORY) {
    $deployArgs.WorkingDirectory = $env:OFFICE_AUTO_INSTALL_WORKING_DIRECTORY.Trim()
}
if ($env:OFFICE_AUTO_INSTALL_SKIP_PREREQ -eq '1') {
    $deployArgs.SkipPrerequisiteTest = $true
}
if ($env:OFFICE_AUTO_INSTALL_SKIP_ADMIN -eq '1') {
    $deployArgs.SkipAdministratorCheck = $true
}

if ($uninstall) {
    Write-Host " Running uninstall (Microsoft 365 Apps removal)..." -ForegroundColor Yellow
    & $payloadPath @deployArgs -Uninstall
} elseif ($env:OFFICE_AUTO_INSTALL_CONFIGURATION_FILE) {
    $cfg = $env:OFFICE_AUTO_INSTALL_CONFIGURATION_FILE.Trim()
    if (-not (Test-Path -LiteralPath $cfg)) {
        Write-Host " ERROR: OFFICE_AUTO_INSTALL_CONFIGURATION_FILE not found: $cfg" -ForegroundColor Red
        exit 1
    }
    & $payloadPath @deployArgs -ConfigurationFile $cfg
} else {
    Write-Host " Preset: $preset  |  Language: $lang" -ForegroundColor Gray
    & $payloadPath @deployArgs -Preset $preset
}

exit $LASTEXITCODE
