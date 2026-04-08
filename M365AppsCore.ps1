# M365AppsCore.ps1
# Office-Auto-Install: shared ODT engine (Microsoft 365 Apps languages + helpers).
# Dot-source from installer scripts: . (Join-Path $PSScriptRoot 'M365AppsCore.ps1')
# Not intended as a standalone entry point.

$script:OdtSetupExeUrl = 'https://officecdn.microsoft.com/pr/wsus/setup.exe'
if (-not $PSScriptRoot) {
    throw 'M365AppsCore.ps1 must be saved on disk and dot-sourced so $PSScriptRoot resolves.'
}
$script:ModuleRoot = $PSScriptRoot

# Culture IDs — Microsoft Learn: overview-deploying-languages-microsoft-365-apps
$script:M365AppsLanguageEntries = @(
    @{ Display = 'Afrikaans'; Id = 'af-za' },
    @{ Display = 'Albanian'; Id = 'sq-al' },
    @{ Display = 'Arabic'; Id = 'ar-sa' },
    @{ Display = 'Armenian'; Id = 'hy-am' },
    @{ Display = 'Assamese'; Id = 'as-in' },
    @{ Display = 'Azerbaijani (Latin)'; Id = 'az-latn-az' },
    @{ Display = 'Bangla (Bangladesh)'; Id = 'bn-bd' },
    @{ Display = 'Bangla (India)'; Id = 'bn-in' },
    @{ Display = 'Basque'; Id = 'eu-es' },
    @{ Display = 'Bosnian (Latin)'; Id = 'bs-latn-ba' },
    @{ Display = 'Bulgarian'; Id = 'bg-bg' },
    @{ Display = 'Catalan'; Id = 'ca-es' },
    @{ Display = 'Catalan (Valencia)'; Id = 'ca-es-valencia' },
    @{ Display = 'Chinese (Simplified)'; Id = 'zh-cn' },
    @{ Display = 'Chinese (Traditional)'; Id = 'zh-tw' },
    @{ Display = 'Croatian'; Id = 'hr-hr' },
    @{ Display = 'Czech'; Id = 'cs-cz' },
    @{ Display = 'Danish'; Id = 'da-dk' },
    @{ Display = 'Dutch'; Id = 'nl-nl' },
    @{ Display = 'English (United Kingdom)'; Id = 'en-gb' },
    @{ Display = 'English (United States)'; Id = 'en-us' },
    @{ Display = 'Estonian'; Id = 'et-ee' },
    @{ Display = 'Finnish'; Id = 'fi-fi' },
    @{ Display = 'French (Canada)'; Id = 'fr-ca' },
    @{ Display = 'French (France)'; Id = 'fr-fr' },
    @{ Display = 'Galician'; Id = 'gl-es' },
    @{ Display = 'Georgian'; Id = 'ka-ge' },
    @{ Display = 'German'; Id = 'de-de' },
    @{ Display = 'Greek'; Id = 'el-gr' },
    @{ Display = 'Gujarati'; Id = 'gu-in' },
    @{ Display = 'Hausa (Latin)'; Id = 'ha-latn-ng' },
    @{ Display = 'Hebrew'; Id = 'he-il' },
    @{ Display = 'Hindi'; Id = 'hi-in' },
    @{ Display = 'Hungarian'; Id = 'hu-hu' },
    @{ Display = 'Icelandic'; Id = 'is-is' },
    @{ Display = 'Igbo'; Id = 'ig-ng' },
    @{ Display = 'Indonesian'; Id = 'id-id' },
    @{ Display = 'Irish'; Id = 'ga-ie' },
    @{ Display = 'isiXhosa'; Id = 'xh-za' },
    @{ Display = 'isiZulu'; Id = 'zu-za' },
    @{ Display = 'Italian'; Id = 'it-it' },
    @{ Display = 'Japanese'; Id = 'ja-jp' },
    @{ Display = 'Kannada'; Id = 'kn-in' },
    @{ Display = 'Kazakh'; Id = 'kk-kz' },
    @{ Display = 'Kinyarwanda'; Id = 'rw-rw' },
    @{ Display = 'Kiswahili'; Id = 'sw-ke' },
    @{ Display = 'Konkani'; Id = 'kok-in' },
    @{ Display = 'Korean'; Id = 'ko-kr' },
    @{ Display = 'Kyrgyz'; Id = 'ky-kg' },
    @{ Display = 'Latvian'; Id = 'lv-lv' },
    @{ Display = 'Lithuanian'; Id = 'lt-lt' },
    @{ Display = 'Luxembourgish'; Id = 'lb-lu' },
    @{ Display = 'Macedonian (North Macedonia)'; Id = 'mk-mk' },
    @{ Display = 'Malay (Latin)'; Id = 'ms-my' },
    @{ Display = 'Malayalam'; Id = 'ml-in' },
    @{ Display = 'Maltese'; Id = 'mt-mt' },
    @{ Display = 'Maori'; Id = 'mi-nz' },
    @{ Display = 'Marathi'; Id = 'mr-in' },
    @{ Display = 'Nepali'; Id = 'ne-np' },
    @{ Display = 'Norwegian Bokmal'; Id = 'nb-no' },
    @{ Display = 'Norwegian Nynorsk'; Id = 'nn-no' },
    @{ Display = 'Odia'; Id = 'or-in' },
    @{ Display = 'Pashto'; Id = 'ps-af' },
    @{ Display = 'Persian'; Id = 'fa-ir' },
    @{ Display = 'Polish'; Id = 'pl-pl' },
    @{ Display = 'Portuguese (Brazil)'; Id = 'pt-br' },
    @{ Display = 'Portuguese (Portugal)'; Id = 'pt-pt' },
    @{ Display = 'Punjabi (Gurmukhi)'; Id = 'pa-in' },
    @{ Display = 'Romanian'; Id = 'ro-ro' },
    @{ Display = 'Romansh'; Id = 'rm-ch' },
    @{ Display = 'Russian'; Id = 'ru-ru' },
    @{ Display = 'Scottish Gaelic'; Id = 'gd-gb' },
    @{ Display = 'Serbian (Cyrillic, Bosnia and Herzegovina)'; Id = 'sr-cyrl-ba' },
    @{ Display = 'Serbian (Cyrillic, Serbia)'; Id = 'sr-cyrl-rs' },
    @{ Display = 'Serbian (Latin, Serbia)'; Id = 'sr-latn-rs' },
    @{ Display = 'Sesotho sa Leboa'; Id = 'nso-za' },
    @{ Display = 'Setswana'; Id = 'tn-za' },
    @{ Display = 'Sinhala'; Id = 'si-lk' },
    @{ Display = 'Slovak'; Id = 'sk-sk' },
    @{ Display = 'Slovenian'; Id = 'sl-si' },
    @{ Display = 'Spanish (Mexico)'; Id = 'es-mx' },
    @{ Display = 'Spanish (Spain)'; Id = 'es-es' },
    @{ Display = 'Swedish'; Id = 'sv-se' },
    @{ Display = 'Tamil'; Id = 'ta-in' },
    @{ Display = 'Tatar (Cyrillic)'; Id = 'tt-ru' },
    @{ Display = 'Telugu'; Id = 'te-in' },
    @{ Display = 'Thai'; Id = 'th-th' },
    @{ Display = 'Turkish'; Id = 'tr-tr' },
    @{ Display = 'Ukrainian'; Id = 'uk-ua' },
    @{ Display = 'Urdu'; Id = 'ur-pk' },
    @{ Display = 'Uzbek (Latin)'; Id = 'uz-latn-uz' },
    @{ Display = 'Vietnamese'; Id = 'vi-vn' },
    @{ Display = 'Welsh'; Id = 'cy-gb' },
    @{ Display = 'Wolof'; Id = 'wo-sn' },
    @{ Display = 'Yoruba'; Id = 'yo-ng' }
)
$script:M365AppsLangById = $null
$script:M365AppsLangByDisplayCi = $null

# Same deployment cannot use these as primary when Visio and/or Project (volume) are installed — Microsoft Learn
# "Languages, culture codes, and companion proofing languages" footnote [1] (not in Project or Visio).
$script:M365AppsLanguageIdsExcludedWithVisioOrProjectVolume = @('en-gb', 'fr-ca', 'es-mx')

function Initialize-M365AppsLanguageLookup {
    if ($null -ne $script:M365AppsLangById) { return }
    $script:M365AppsLangById = @{}
    $script:M365AppsLangByDisplayCi = @{}
    foreach ($e in $script:M365AppsLanguageEntries) {
        $id = $e.Id.ToLowerInvariant()
        $script:M365AppsLangById[$id] = $true
        $script:M365AppsLangByDisplayCi[$e.Display.ToLowerInvariant()] = $id
    }
}

function Get-M365AppsSupportedLanguages {
    <#
    .SYNOPSIS
        Lists primary language packs for ODT, sorted by display name.
    .PARAMETER Preset
        When set to a Visio+Project preset, excludes languages that Microsoft does not ship for those apps in the same install.
    .PARAMETER IncludeVisio
    .PARAMETER IncludeProject
        For custom deployments: if either is set, applies the same Visio/Project language exclusions.
    #>
    [CmdletBinding()]
    param(
        [string]$Preset = '',
        [switch]$IncludeVisio,
        [switch]$IncludeProject
    )
    Initialize-M365AppsLanguageLookup
    $applyVpFilter =
        ($Preset -in @('O365ProPlusVisioProject', 'O365ProPlusVisioProject-VDI')) -or
        ($IncludeVisio -or $IncludeProject)
    $rows = $script:M365AppsLanguageEntries | Sort-Object { $_.Display }
    if ($applyVpFilter) {
        $ex = $script:M365AppsLanguageIdsExcludedWithVisioOrProjectVolume
        $rows = $rows | Where-Object { $ex -notcontains ($_.Id.ToLowerInvariant()) }
    }
    $rows | ForEach-Object {
        [pscustomobject]@{ Display = $_.Display; Id = $_.Id.ToLowerInvariant() }
    }
}

function Assert-M365AppsLanguageCompatibleWithDeployment {
    <#
    .SYNOPSIS
        Ensures the primary language is allowed for the chosen preset (Visio/Project rules per Microsoft documentation).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LanguageId,
        [string]$Preset = '',
        [switch]$CustomIncludeVisio,
        [switch]$CustomIncludeProject
    )
    $id = $LanguageId.ToLowerInvariant()
    $useVpRule =
        ($Preset -in @('O365ProPlusVisioProject', 'O365ProPlusVisioProject-VDI')) -or
        ($CustomIncludeVisio -or $CustomIncludeProject)
    if (-not $useVpRule) { return }
    if ($script:M365AppsLanguageIdsExcludedWithVisioOrProjectVolume -contains $id) {
        throw "Primary language '$LanguageId' is not valid together with Visio and/or Project in one deployment (Microsoft does not offer this combination for those apps: en-gb, fr-ca, es-mx). Choose English (United States) or another listed language, or use a profile without Visio/Project."
    }
}

function Resolve-M365AppsLanguageId {
    <#
    .SYNOPSIS
        Resolves a display name or ll-cc culture ID to the canonical lowercase ID used in ODT XML.
    .NOTES
        The bundled list follows Microsoft’s published primary languages. Office setup.exe may still report
        that a language is “not available” for a specific product/channel (e.g. Business vs ProPlus, or Visio/Project).
        Unknown-but-valid-looking culture tags (e.g. new IDs before we update the list) are passed through.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Text
    )
    Initialize-M365AppsLanguageLookup
    $t = $Text.Trim()
    if ([string]::IsNullOrWhiteSpace($t)) { return 'en-us' }
    $tl = $t.ToLowerInvariant()
    if ($script:M365AppsLangById.ContainsKey($tl)) { return $tl }
    if ($script:M365AppsLangByDisplayCi.ContainsKey($tl)) { return $script:M365AppsLangByDisplayCi[$tl] }
    # Pass through ODT-style tags (e.g. sr-latn-rs, ca-es-valencia) so we do not block IDs Microsoft adds later.
    if ($tl -match '^[a-z]{2}(-[a-z0-9]+)+$') {
        Write-Warning "Language '$tl' is not in the built-in catalog; using it as-is. If setup says the language is unavailable, try another language or check that your product/channel supports it."
        return $tl
    }
    throw "Unknown Office language: '$Text'. Enter a culture like en-us or de-de, or pick a name from Get-M365AppsSupportedLanguages."
}

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
