#Requires -Version 5.1
<#
.SYNOPSIS
    Deploy Microsoft 365 Apps using the Office Deployment Tool and preset XML (AVD, physical, business).

.DESCRIPTION
    Intended for automation: Azure Virtual Desktop, Windows 365, MECM, Intune Win32 wrapper scripts, golden images.
    Preset XML is maintained under .\configs.

.PARAMETER Preset
    Bundled configuration name under .\configs (see repository).

.PARAMETER ConfigurationFile
    Overrides Preset with a full path to any valid ODT configuration XML.

.PARAMETER Uninstall
    Runs removal using configs\Uninstall-Microsoft365Apps.xml

.PARAMETER ExcludeApp
    One or more ODT ExcludeApp IDs (e.g. Teams, OneDrive, Access) merged into the suite product in preset or custom XML.
    Valid IDs match Microsoft’s documentation (see M365AppsCore.ps1). Ignored when -Uninstall is used.

.EXAMPLE
    .\Deploy-Microsoft365Apps.ps1 -Preset O365ProPlus-VDI -LanguageId en-us

.EXAMPLE
    .\Deploy-Microsoft365Apps.ps1 -ConfigurationFile 'C:\Deploy\custom.xml'

.EXAMPLE
    .\Deploy-Microsoft365Apps.ps1 -Uninstall

.EXAMPLE
    .\Deploy-Microsoft365Apps.ps1 -Preset O365ProPlus -LanguageId en-us -ExcludeApp Teams,OneDrive,Access
#>
[CmdletBinding(DefaultParameterSetName = 'Deploy')]
param(
    [Parameter(ParameterSetName = 'Deploy')]
    [ValidateSet(
        'O365ProPlus',
        'O365ProPlus-VDI',
        'O365Business',
        'O365Business-VDI',
        'O365ProPlusVisioProject',
        'O365ProPlusVisioProject-VDI'
    )]
    [string]$Preset = 'O365ProPlus',

    [Parameter(ParameterSetName = 'Deploy')]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [string]$ConfigurationFile,

    [Parameter(ParameterSetName = 'Uninstall')]
    [switch]$Uninstall,

    [ValidateSet('32', '64')]
    [string]$OfficeClientEdition = '64',

    [ValidateSet('Current', 'MonthlyEnterprise', 'SemiAnnualEnterprise', 'SemiAnnualPreview')]
    [string]$Channel,

    [string]$LanguageId = 'en-us',

    [Parameter(ParameterSetName = 'Deploy')]
    [string[]]$ExcludeApp = @(),

    [string]$WorkingDirectory,

    [switch]$SkipPrerequisiteTest,

    [switch]$SkipAdministratorCheck
)

$ErrorActionPreference = 'Stop'

$corePath = Join-Path $PSScriptRoot 'M365AppsCore.ps1'
if (-not (Test-Path -LiteralPath $corePath)) {
    throw "Missing M365AppsCore.ps1 (expected next to this script): $corePath"
}
. $corePath

if (-not $SkipAdministratorCheck -and -not (Test-M365AppsAdministrator)) {
    throw 'Run this script from an elevated PowerShell session (Run as Administrator).'
}

if (-not $SkipPrerequisiteTest) {
    Test-M365AppsPrerequisites
}

if ($PSCmdlet.ParameterSetName -eq 'Deploy' -and -not $ConfigurationFile) {
    Assert-M365AppsLanguageCompatibleWithDeployment -LanguageId $LanguageId -Preset $Preset
}

$excludeNormalized = @()
if ($PSCmdlet.ParameterSetName -eq 'Deploy' -and $ExcludeApp -and $ExcludeApp.Count -gt 0) {
    $excludeNormalized = Resolve-M365AppsExcludeAppIdList -CommaSeparatedText ($ExcludeApp -join ',')
}

$work = if ($WorkingDirectory) {
    $WorkingDirectory
} else {
    Join-Path $env:TEMP 'M365AppsDeploy'
}
if (-not (Test-Path -LiteralPath $work)) {
    New-Item -ItemType Directory -Path $work -Force | Out-Null
}

$setupExe = Join-Path $work 'setup.exe'
Save-M365AppsOfficeDeploymentTool -DestinationPath $setupExe

if ($Uninstall) {
    $cfg = Get-M365AppsUninstallConfigurationPath
    Copy-Item -LiteralPath $cfg -Destination (Join-Path $work 'config.xml') -Force
} elseif ($ConfigurationFile) {
    Copy-M365AppsConfigurationWithOverrides -SourcePath $ConfigurationFile -DestinationPath (Join-Path $work 'config.xml') -OfficeClientEdition $OfficeClientEdition -Channel $Channel -LanguageId $LanguageId -ExcludeAppIds $excludeNormalized
} else {
    $src = Get-M365AppsPresetConfigurationPath -Preset $Preset
    $ch = if ($Channel) { $Channel } else { $null }
    Copy-M365AppsConfigurationWithOverrides -SourcePath $src -DestinationPath (Join-Path $work 'config.xml') -OfficeClientEdition $OfficeClientEdition -Channel $ch -LanguageId $LanguageId -ExcludeAppIds $excludeNormalized
}

$configPath = Join-Path $work 'config.xml'
$code = Start-M365AppsSetup -SetupExePath $setupExe -ConfigurationPath $configPath -Wait
if ($code -ne 0) {
    Write-Warning "setup.exe exited with code $code. Review Office logs under %TEMP% or the path set in your configuration."
}
exit $code
