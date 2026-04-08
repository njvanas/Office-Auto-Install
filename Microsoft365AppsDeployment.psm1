# Microsoft365AppsDeployment.psm1
# Office-Auto-Install: automation for Microsoft 365 Apps deployment using official Microsoft tooling (ODT).

# Office Deployment Tool (setup.exe) — retrieved from Microsoft's Office CDN (same delivery channel Microsoft documents for ODT).
$script:OdtSetupExeUrl = 'https://officecdn.microsoft.com/pr/wsus/setup.exe'
$script:ModuleRoot = $PSScriptRoot

function Get-M365AppsOfficialOfficeDeploymentToolUri {
    <#
    .SYNOPSIS
        Returns the official Microsoft CDN URL used to download the Office Deployment Tool (setup.exe).
    #>
    [CmdletBinding()]
    param()
    $script:OdtSetupExeUrl
}

function Get-M365AppsConfigDirectory {
    <#
    .SYNOPSIS
        Returns the path to bundled preset configuration XML files under configs\.
    #>
    [CmdletBinding()]
    param()
    Join-Path -Path $script:ModuleRoot -ChildPath 'configs'
}

function Test-M365AppsAdministrator {
    [CmdletBinding()]
    param()
    $p = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-M365AppsPrerequisites {
    <#
    .SYNOPSIS
        Validates disk space and optional network reachability for ODT / Office setup.
    #>
    [CmdletBinding()]
    param(
        [int]$MinimumFreeSpaceGB = 4,
        [switch]$SkipNetworkTest
    )
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$($env:SystemDrive.TrimEnd('\'))'" -ErrorAction Stop
    $freeGb = [math]::Round($disk.FreeSpace / 1GB, 2)
    if ($freeGb -lt $MinimumFreeSpaceGB) {
        throw "Insufficient disk space: $freeGb GB free; at least $MinimumFreeSpaceGB GB required on $($env:SystemDrive)."
    }
    if (-not $SkipNetworkTest) {
        try {
            $null = Invoke-WebRequest -Uri 'https://www.microsoft.com' -UseBasicParsing -TimeoutSec 15
        } catch {
            throw "Network check failed: $($_.Exception.Message)"
        }
    }
}

function Save-M365AppsOfficeDeploymentTool {
    <#
    .SYNOPSIS
        Downloads setup.exe (Office Deployment Tool) from the official Microsoft Office CDN.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DestinationPath,
        [string]$Uri = $script:OdtSetupExeUrl
    )
    $dir = Split-Path -Parent $DestinationPath
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Invoke-WebRequest -Uri $Uri -OutFile $DestinationPath -UseBasicParsing
    if (-not (Test-Path -LiteralPath $DestinationPath) -or ((Get-Item -LiteralPath $DestinationPath).Length -lt 100KB)) {
        throw 'setup.exe is missing or too small; download from Microsoft may have failed.'
    }
}

function Get-M365AppsUninstallConfigurationPath {
    [CmdletBinding()]
    param()
    $path = Join-Path (Get-M365AppsConfigDirectory) 'Uninstall-Microsoft365Apps.xml'
    if (-not (Test-Path -LiteralPath $path)) { throw "Uninstall configuration not found: $path" }
    (Resolve-Path -LiteralPath $path).Path
}

function Get-M365AppsPresetConfigurationPath {
    <#
    .SYNOPSIS
        Resolves a preset name to a bundled configuration file path (see /configs).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet(
            'O365ProPlus',
            'O365ProPlus-VDI',
            'O365Business',
            'O365Business-VDI',
            'O365ProPlusVisioProject',
            'O365ProPlusVisioProject-VDI'
        )]
        [string]$Preset
    )
    $name = "$Preset.xml"
    $path = Join-Path (Get-M365AppsConfigDirectory) $name
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Preset configuration not found: $path"
    }
    return (Resolve-Path -LiteralPath $path).Path
}

function Set-M365AppsConfigurationOverrides {
    <#
    .SYNOPSIS
        Applies common overrides to an ODT configuration XML (channel, architecture, language).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [ValidateSet('32', '64')]
        [string]$OfficeClientEdition = '64',
        [string]$Channel,
        [string]$LanguageId = 'en-us'
    )
    if ($Channel) {
        $valid = @('Current', 'MonthlyEnterprise', 'SemiAnnualEnterprise', 'SemiAnnualPreview')
        if ($Channel -notin $valid) { throw "Invalid Channel '$Channel'. Use: $($valid -join ', ')." }
    }
    [xml]$doc = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $cfg = $doc.Configuration
    if (-not $cfg) { throw 'Invalid ODT XML: missing Configuration root.' }

    $add = $cfg.Add
    if ($add) {
        $add.SetAttribute('OfficeClientEdition', $OfficeClientEdition)
        if ($Channel) {
            $add.SetAttribute('Channel', $Channel)
        }
    }

    foreach ($product in $doc.SelectNodes('//Product')) {
        foreach ($lang in $product.SelectNodes('Language')) {
            $lang.SetAttribute('ID', $LanguageId)
        }
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $Path).Path, $doc.OuterXml, $utf8NoBom)
}

function Set-M365AppsConfigurationDisplayLevel {
    <#
    .SYNOPSIS
        Sets Display Level (and AcceptEULA) on an ODT configuration file (e.g. after copying a preset).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [ValidateSet('Full', 'None')]
        [string]$Level = 'Full',
        [string]$AcceptEULA = 'TRUE'
    )
    [xml]$doc = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $cfg = $doc.Configuration
    if (-not $cfg) { throw 'Invalid ODT XML: missing Configuration root.' }
    $disp = $cfg.Display
    if (-not $disp) {
        $disp = $doc.CreateElement('Display')
        [void]$cfg.AppendChild($disp)
    }
    $disp.SetAttribute('Level', $Level)
    $disp.SetAttribute('AcceptEULA', $AcceptEULA)
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $Path).Path, $doc.OuterXml, $utf8NoBom)
}

function Copy-M365AppsConfigurationWithOverrides {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,
        [Parameter(Mandatory)]
        [string]$DestinationPath,
        [ValidateSet('32', '64')]
        [string]$OfficeClientEdition = '64',
        [string]$Channel,
        [string]$LanguageId = 'en-us'
    )
    Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    Set-M365AppsConfigurationOverrides -Path $DestinationPath -OfficeClientEdition $OfficeClientEdition -Channel $Channel -LanguageId $LanguageId
}

function Start-M365AppsSetup {
    <#
    .SYNOPSIS
        Runs setup.exe /configure against a configuration XML in the same working directory.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SetupExePath,
        [Parameter(Mandatory)]
        [string]$ConfigurationPath,
        [switch]$Wait
    )
    $wd = Split-Path -Parent $ConfigurationPath
    $name = Split-Path -Leaf $ConfigurationPath
    $arg = "/configure `"$name`""
    if ($Wait) {
        $p = Start-Process -FilePath $SetupExePath -ArgumentList $arg -WorkingDirectory $wd -PassThru -NoNewWindow -Wait
        return $p.ExitCode
    }
    $p = Start-Process -FilePath $SetupExePath -ArgumentList $arg -WorkingDirectory $wd -PassThru -NoNewWindow
    $null = $p.WaitForExit()
    return $p.ExitCode
}

function New-M365AppsInteractiveConfiguration {
    <#
    .SYNOPSIS
        Builds a minimal ODT XML for interactive (GUI/console) flows — retail/volume IDs as selected.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProductId,
        [string]$LanguageId = 'en-us',
        [ValidateSet('32', '64')]
        [string]$OfficeClientEdition = '64',
        [ValidateSet('Current', 'MonthlyEnterprise', 'SemiAnnualEnterprise')]
        [string]$Channel = 'Current',
        [ValidateSet('Full', 'None')]
        [string]$DisplayLevel = 'Full',
        [switch]$IncludeVisio,
        [switch]$IncludeProject,
        [switch]$AutoActivate
    )
    $products = @()
    $products += "<Product ID='$ProductId'>`n  <Language ID='$LanguageId' />`n</Product>"
    if ($IncludeVisio) {
        $products += "<Product ID='VisioPro2021Volume'>`n  <Language ID='$LanguageId' />`n</Product>"
    }
    if ($IncludeProject) {
        $products += "<Product ID='ProjectPro2021Volume'>`n  <Language ID='$LanguageId' />`n</Product>"
    }

    $props = @()
    if ($AutoActivate) {
        $props += '  <Property Name="AUTOACTIVATE" Value="1" />'
    }
    $propsBlock = if ($props.Count) { "`n$($props -join "`n")" } else { '' }

    @"
<Configuration>
  <Add OfficeClientEdition="$OfficeClientEdition" Channel="$Channel">
    $($products -join "`n    ")
  </Add>
  <Display Level="$DisplayLevel" AcceptEULA="TRUE" />$propsBlock
</Configuration>
"@
}

Export-ModuleMember -Function @(
    'Get-M365AppsOfficialOfficeDeploymentToolUri',
    'Get-M365AppsConfigDirectory',
    'Set-M365AppsConfigurationDisplayLevel',
    'Test-M365AppsAdministrator',
    'Test-M365AppsPrerequisites',
    'Save-M365AppsOfficeDeploymentTool',
    'Get-M365AppsUninstallConfigurationPath',
    'Get-M365AppsPresetConfigurationPath',
    'Set-M365AppsConfigurationOverrides',
    'Copy-M365AppsConfigurationWithOverrides',
    'Start-M365AppsSetup',
    'New-M365AppsInteractiveConfiguration'
)
