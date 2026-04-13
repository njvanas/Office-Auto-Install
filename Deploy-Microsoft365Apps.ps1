#Requires -Version 5.1
<#
.SYNOPSIS
    Deploy Microsoft 365 Apps using the Office Deployment Tool (ODT).

.DESCRIPTION
    Intended for automation: Azure Virtual Desktop, Windows 365, MECM, Intune Win32 wrapper scripts, golden images.
    All configuration XML (retail profiles, Visio + Project bundles, uninstall) is generated from parameters.

.PARAMETER RetailProfile
    One of four retail suite profiles: enterprise/business combined with physical desktop or VDI/shared PC.
    Default channel is chosen per profile unless -Channel is set.

.PARAMETER Channel
    Optional override for the ODT Add Channel attribute. Must match values documented by Microsoft:
    https://learn.microsoft.com/microsoft-365-apps/deploy/office-deployment-tool-configuration-options

.PARAMETER Bundle
    Generates multi-product XML (M365 Apps + Visio + Project) for the named bundle scenario (same names as before).

.PARAMETER Uninstall
    Removes Microsoft 365 Apps (Click-to-Run) using generated Remove All configuration XML.

.PARAMETER ExcludeApp
    ODT ExcludeApp IDs merged into the suite product (retail or bundle). Ignored when -Uninstall is used.

.EXAMPLE
    .\Deploy-Microsoft365Apps.ps1 -RetailProfile EnterpriseVDI -LanguageId en-us

.EXAMPLE
    .\Deploy-Microsoft365Apps.ps1 -Uninstall

.EXAMPLE
    .\Deploy-Microsoft365Apps.ps1 -RetailProfile EnterprisePhysical -LanguageId en-us -ExcludeApp Teams,OneDrive,Access

.EXAMPLE
    .\Deploy-Microsoft365Apps.ps1 -Bundle O365ProPlusVisioProject-VDI -LanguageId en-us
#>
[CmdletBinding(DefaultParameterSetName = 'Retail')]
param(
    [Parameter(ParameterSetName = 'Uninstall')]
    [switch]$Uninstall,

    [Parameter(ParameterSetName = 'Bundle', Mandatory)]
    [ValidateSet(
        'O365ProPlusVisioProject',
        'O365ProPlusVisioProject-Retail',
        'O365ProPlusVisioProject-2024',
        'O365ProPlusVisioProject-VDI',
        'O365ProPlusVisioProject-Retail-VDI',
        'O365ProPlusVisioProject-2024-VDI'
    )]
    [string]$Bundle,

    [Parameter(ParameterSetName = 'Retail')]
    [ValidateSet('EnterprisePhysical', 'EnterpriseVDI', 'BusinessPhysical', 'BusinessVDI')]
    [string]$RetailProfile = 'EnterprisePhysical',

    [ValidateSet('32', '64')]
    [string]$OfficeClientEdition = '64',

    [string]$Channel,

    [string]$LanguageId = 'en-us',

    [Parameter(ParameterSetName = 'Retail')]
    [Parameter(ParameterSetName = 'Bundle')]
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

if ($Channel) {
    Assert-M365AppsOdtAddChannelValue -Channel $Channel
}

if (-not $SkipAdministratorCheck -and -not (Test-M365AppsAdministrator)) {
    throw 'Run this script from an elevated PowerShell session (Run as Administrator).'
}

if (-not $SkipPrerequisiteTest) {
    Test-M365AppsPrerequisites
}

if ($PSCmdlet.ParameterSetName -in @('Retail', 'Bundle')) {
    $assertPreset = if ($PSCmdlet.ParameterSetName -eq 'Bundle') { $Bundle } else { '' }
    Assert-M365AppsLanguageCompatibleWithDeployment -LanguageId $LanguageId -Preset $assertPreset
}

$excludeNormalized = @()
if ($PSCmdlet.ParameterSetName -in @('Retail', 'Bundle') -and $ExcludeApp -and $ExcludeApp.Count -gt 0) {
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

$destCfg = Join-Path $work 'config.xml'
$ch = if ($Channel) { $Channel } else { $null }
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if ($Uninstall) {
    $xmlUn = New-M365AppsUninstallConfigurationXml -DisplayLevel 'None'
    [System.IO.File]::WriteAllText($destCfg, $xmlUn, $utf8NoBom)
} elseif ($PSCmdlet.ParameterSetName -eq 'Bundle') {
    $xmlBundle = New-M365AppsVisioProjectBundleConfigurationXml -Bundle $Bundle -OfficeClientEdition $OfficeClientEdition `
        -Channel $ch -LanguageId $LanguageId -DisplayLevel 'None'
    Export-M365AppsConfigurationStringToPathWithOverrides -ConfigurationXml $xmlBundle -DestinationPath $destCfg `
        -OfficeClientEdition $OfficeClientEdition -Channel $ch -LanguageId $LanguageId -ExcludeAppIds $excludeNormalized
} else {
    $xml = New-M365AppsO365ConfigurationForRetailProfile -RetailProfile $RetailProfile -OfficeClientEdition $OfficeClientEdition `
        -Channel $ch -LanguageId $LanguageId -DisplayLevel 'None' -AdditionalExcludeAppIds $excludeNormalized
    [System.IO.File]::WriteAllText($destCfg, $xml, $utf8NoBom)
}

$code = Start-M365AppsSetup -SetupExePath $setupExe -ConfigurationPath $destCfg -Wait
if ($code -ne 0) {
    Write-Warning "setup.exe exited with code $code. Review Office logs under %TEMP% or the path set in your configuration."
}
exit $code
