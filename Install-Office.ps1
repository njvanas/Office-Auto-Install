#Requires -Version 5.1
<#
.SYNOPSIS
    Interactive Microsoft 365 / Office installer using the Office Deployment Tool.

.DESCRIPTION
    Console wizard: choose a deployment profile (bundled preset XML under .\configs) or a custom interactive
    configuration. Primary languages match Microsoft 365 Apps (see Get-M365AppsSupportedLanguages).
    For unattended deploys use Deploy-Microsoft365Apps.ps1. ODT is downloaded from Microsoft's CDN.

.NOTES
    Requires elevation. Example: powershell -ExecutionPolicy Bypass -File .\Install-Office.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$corePath = Join-Path $PSScriptRoot 'M365AppsCore.ps1'
if (-not (Test-Path -LiteralPath $corePath)) {
    throw "Missing M365AppsCore.ps1 (expected next to this script): $corePath"
}
. $corePath

function Request-AdminElevation {
    if (Test-M365AppsAdministrator) { return }
    Write-Host 'Administrator rights are required. Launching elevated PowerShell...' -ForegroundColor Yellow
    $scriptPath = if ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } else { $PSCommandPath }
    if ($scriptPath) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$scriptPath`"") -Verb RunAs
    } else {
        throw 'Cannot elevate: save this script to a file and run again.'
    }
    exit 0
}

function Read-Choice {
    param([string]$Prompt, [string[]]$Valid, [string]$Default)
    do {
        $r = Read-Host "$Prompt [$Default]"
        if ([string]::IsNullOrWhiteSpace($r)) { $r = $Default }
        if ($Valid -contains $r) { return $r }
        Write-Host "Invalid. Enter one of: $($Valid -join ', ')" -ForegroundColor Red
    } while ($true)
}

function Get-PresetNameFromMenuChoice {
    param([string]$Choice)
    $map = @{
        '1' = 'O365ProPlus'
        '2' = 'O365ProPlus-VDI'
        '3' = 'O365Business'
        '4' = 'O365Business-VDI'
        '5' = 'O365ProPlusVisioProject'
        '6' = 'O365ProPlusVisioProject-VDI'
    }
    if ($map.ContainsKey($Choice)) { return $map[$Choice] }
    return $null
}

function Resolve-ConsoleChannelOverride {
    <#
    Preset path: 1 = use XML default (no override), 2 = Current, 3 = SemiAnnualEnterprise.
    Custom path: 1 = Current, 2 = SemiAnnualEnterprise.
    #>
    param(
        [bool]$IsCustomProfile,
        [string]$Choice
    )
    if ($IsCustomProfile) {
        if ($Choice -eq '2') { return 'SemiAnnualEnterprise' }
        return 'Current'
    }
    switch ($Choice) {
        '1' { return $null }
        '2' { return 'Current' }
        '3' { return 'SemiAnnualEnterprise' }
        default { return $null }
    }
}

function Read-ValidatedLanguageId {
    Write-Host ''
    Write-Host ' Primary language for Office (Microsoft 365 Apps culture ID, e.g. en-us, de-de, ja-jp).' -ForegroundColor White
    Write-Host ' Run Get-M365AppsSupportedLanguages in PowerShell for the full list with display names.' -ForegroundColor DarkGray
    do {
        $raw = Read-Host ' LanguageId [en-us]'
        try {
            return Resolve-M365AppsLanguageId -Text $raw
        } catch {
            Write-Host " $($_.Exception.Message)" -ForegroundColor Red
        }
    } while ($true)
}

Request-AdminElevation

$installerFolder = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'OfficeInstaller' } else { Join-Path $env:TEMP 'OfficeInstaller' }
if (Test-Path $installerFolder) {
    Remove-Item -Path (Join-Path $installerFolder '*') -Recurse -Force -ErrorAction SilentlyContinue
} else {
    New-Item -ItemType Directory -Path $installerFolder -Force | Out-Null
}

$logFile = Join-Path $installerFolder 'installer.log'
function Write-Log([string]$Message) {
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

Write-Log 'Install-Office.ps1 started'

try {
    Test-M365AppsPrerequisites
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    Read-Host 'Press Enter to exit'
    exit 1
}

Write-Host ''
Write-Host ' Microsoft 365 Apps — interactive setup' -ForegroundColor Cyan
Write-Host ' Uses Microsoft Office Deployment Tool (official).' -ForegroundColor Gray
Write-Host ''

Write-Host ' Deployment mode:' -ForegroundColor White
Write-Host '  [1] Deployment profile — uses preset XML (same as GUI; recommended for M365 business/enterprise, VDI, etc.)'
Write-Host '  [2] Custom — Office 2024 / LTSC / M365 retail XML built in this session (no preset file)'
$deployMode = Read-Choice -Prompt 'Select' -Valid @('1', '2') -Default '1'
$usePreset = ($deployMode -eq '1')

if ($usePreset) {
    try {
        $null = Get-M365AppsPresetConfigurationPath -Preset 'O365ProPlus'
    } catch {
        Write-Host ' Preset configurations are missing. Ensure configs\ is next to this script and M365AppsCore.ps1.' -ForegroundColor Red
        Write-Host " $($_.Exception.Message)" -ForegroundColor Yellow
        Read-Host 'Press Enter to exit'
        exit 1
    }
}

Write-Host ''
$arch = Read-Choice -Prompt 'Architecture: 1=64-bit  2=32-bit' -Valid @('1', '2') -Default '1'
$bit = if ($arch -eq '2') { '32' } else { '64' }

$presetName = $null
$isCustom = $false

if ($usePreset) {
    Write-Host ''
    Write-Host ' Deployment profile (matches files under configs\):' -ForegroundColor White
    Write-Host '  [1] Microsoft 365 Apps — enterprise (physical / desktop)'
    Write-Host '  [2] Microsoft 365 Apps — enterprise (VDI / shared PC)'
    Write-Host '  [3] Microsoft 365 Apps — business (physical / desktop)'
    Write-Host '  [4] Microsoft 365 Apps — business (VDI / shared PC)'
    Write-Host '  [5] M365 enterprise + Visio & Project (physical / desktop)'
    Write-Host '  [6] M365 enterprise + Visio & Project (VDI / shared PC)'
    $prof = Read-Choice -Prompt 'Select' -Valid @('1', '2', '3', '4', '5', '6') -Default '1'
    $presetName = Get-PresetNameFromMenuChoice -Choice $prof
    $isCustom = $false
} else {
    $isCustom = $true
    Write-Host ''
    Write-Host ' Office product:' -ForegroundColor White
    Write-Host '  [1] Office 2024 Pro Plus (Retail)'
    Write-Host '  [2] Office LTSC 2021 (Volume)'
    Write-Host '  [3] Microsoft 365 Apps (Retail)'
    $ed = Read-Choice -Prompt 'Select' -Valid @('1', '2', '3') -Default '3'
    $editionMap = @{
        '1' = 'ProPlus2024Retail'
        '2' = 'ProPlus2021Volume'
        '3' = 'O365ProPlusRetail'
    }
    $productId = $editionMap[$ed]

    Write-Host ''
    $visio = Read-Choice -Prompt 'Include Visio Professional 2021 (volume)? 1=Yes 2=No' -Valid @('1', '2') -Default '2'
    $project = Read-Choice -Prompt 'Include Project Professional 2021 (volume)? 1=Yes 2=No' -Valid @('1', '2') -Default '2'
}

Write-Host ''
if ($usePreset) {
    Write-Host ' Update channel: preset XML has a default; override only if you need to.' -ForegroundColor DarkGray
    $ch = Read-Choice -Prompt 'Channel: 1=Use preset default  2=Current  3=SemiAnnualEnterprise' -Valid @('1', '2', '3') -Default '1'
} else {
    Write-Host ' Update channel: Current = frequent; SemiAnnualEnterprise = less frequent.' -ForegroundColor DarkGray
    $ch = Read-Choice -Prompt 'Channel: 1=Current  2=SemiAnnualEnterprise' -Valid @('1', '2') -Default '1'
}
$channelOverride = Resolve-ConsoleChannelOverride -IsCustomProfile $isCustom -Choice $ch

$languageId = Read-ValidatedLanguageId
try {
    if ($usePreset) {
        Assert-M365AppsLanguageCompatibleWithDeployment -LanguageId $languageId -Preset $presetName
    } else {
        Assert-M365AppsLanguageCompatibleWithDeployment -LanguageId $languageId `
            -CustomIncludeVisio:($visio -eq '1') -CustomIncludeProject:($project -eq '1')
    }
} catch {
    Write-Host " $($_.Exception.Message)" -ForegroundColor Red
    Read-Host 'Press Enter to exit'
    exit 1
}
$languageLabel = (Get-M365AppsSupportedLanguages | Where-Object { $_.Id -eq $languageId } | Select-Object -First 1 -ExpandProperty Display)
if (-not $languageLabel) { $languageLabel = $languageId }

Write-Host ''
$ui = Read-Choice -Prompt 'Installer UI: 1=Show progress  2=Quiet' -Valid @('1', '2') -Default '1'
$displayLevel = if ($ui -eq '1') { 'Full' } else { 'None' }

$autoActivate = $false
if ($isCustom) {
    $aa = Read-Choice -Prompt 'Set AUTOACTIVATE=1 (typical for personal retail)? 1=Yes 2=No' -Valid @('1', '2') -Default '2'
    $autoActivate = ($aa -eq '1')
}

$configPath = Join-Path $installerFolder 'config.xml'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if ($usePreset) {
    Write-Log "Building config from preset $presetName"
    $src = Get-M365AppsPresetConfigurationPath -Preset $presetName
    Copy-M365AppsConfigurationWithOverrides -SourcePath $src -DestinationPath $configPath `
        -OfficeClientEdition $bit -LanguageId $languageId -Channel $channelOverride
    Set-M365AppsConfigurationDisplayLevel -Path $configPath -Level $displayLevel
} else {
    $xml = New-M365AppsInteractiveConfiguration -ProductId $productId -LanguageId $languageId -OfficeClientEdition $bit `
        -Channel $channelOverride -DisplayLevel $displayLevel -IncludeVisio:($visio -eq '1') -IncludeProject:($project -eq '1') -AutoActivate:$autoActivate
    [System.IO.File]::WriteAllText($configPath, $xml, $utf8NoBom)
}
Write-Log "Wrote configuration to $configPath"

$setupExe = Join-Path $installerFolder 'setup.exe'
Write-Host ''
Write-Host 'Downloading Office Deployment Tool...' -ForegroundColor Cyan
try {
    Save-M365AppsOfficeDeploymentTool -DestinationPath $setupExe
} catch {
    Write-Log "Download failed: $_"
    Write-Host $_ -ForegroundColor Red
    Read-Host 'Press Enter to exit'
    exit 1
}

Write-Host 'Starting setup (this may take a long time)...' -ForegroundColor Green
Write-Log "Running setup /configure (working directory: $installerFolder)"
Set-Location -LiteralPath $installerFolder
$exitCode = Start-M365AppsSetup -SetupExePath $setupExe -ConfigurationPath $configPath -Wait
Write-Log "setup.exe exit code: $exitCode"

if ($exitCode -eq 0) {
    Write-Host 'Setup finished successfully (exit 0).' -ForegroundColor Green
} else {
    Write-Warning "Setup exited with code $exitCode. Office may still be installed; check logs and Start Menu."
}

Write-Log 'Done.'
Read-Host 'Press Enter to close'
