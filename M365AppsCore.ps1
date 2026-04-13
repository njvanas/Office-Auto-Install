# M365AppsCore.ps1
# Office-Auto-Install: shared ODT engine (Microsoft 365 Apps languages + helpers).
# Dot-source from installer scripts: . (Join-Path $PSScriptRoot 'M365AppsCore.ps1')
# Not intended as a standalone entry point.
# Authoritative ODT configuration XML reference: https://learn.microsoft.com/microsoft-365-apps/deploy/office-deployment-tool-configuration-options

$script:OdtSetupExeUrl = 'https://officecdn.microsoft.com/pr/wsus/setup.exe'
if (-not $PSScriptRoot) {
    throw 'M365AppsCore.ps1 must be saved on disk and dot-sourced so $PSScriptRoot resolves.'
}
$script:ModuleRoot = $PSScriptRoot

# Office Deployment Tool — configuration elements, attributes, and allowed values:
# https://learn.microsoft.com/microsoft-365-apps/deploy/office-deployment-tool-configuration-options
$script:M365AppsOdtConfigurationLearnUrl = 'https://learn.microsoft.com/microsoft-365-apps/deploy/office-deployment-tool-configuration-options'

# <Add Channel="..."> — values documented on the page above (plus legacy names Microsoft still accepts).
$script:M365AppsValidOdtAddChannelValues = @(
    'BetaChannel',
    'CurrentPreview',
    'Current',
    'MonthlyEnterprise',
    'SemiAnnualPreview',
    'SemiAnnual',
    'SemiAnnualEnterprise',
    'PerpetualVL2024',
    'PerpetualVL2021',
    'PerpetualVL2019'
)

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

# ODT ExcludeApp ID — Microsoft Learn (Product element). Not used on Visio/Project standalone products here.
$script:M365AppsValidExcludeAppIds = @(
    'Access', 'Excel', 'Groove', 'Lync', 'OneDrive', 'OneNote', 'Outlook', 'OutlookForWindows',
    'PowerPoint', 'Publisher', 'Teams', 'Word'
)
$script:M365AppsExcludeAppSuiteProductIds = @(
    'O365ProPlusRetail', 'O365BusinessRetail', 'ProPlus2024Retail', 'ProPlus2021Volume'
)

# Bundled presets that include Visio + Project — primary-language rules in Get/Assert-M365Apps*Visio*
$script:M365AppsVisioProjectPresetNames = @(
    'O365ProPlusVisioProject',
    'O365ProPlusVisioProject-Retail',
    'O365ProPlusVisioProject-2024',
    'O365ProPlusVisioProject-VDI',
    'O365ProPlusVisioProject-Retail-VDI',
    'O365ProPlusVisioProject-2024-VDI'
)

function Get-M365AppsExcludeAppCatalog {
    <#
    .SYNOPSIS
        Returns display metadata for optional ExcludeApp checkboxes (ODT).
    #>
    [CmdletBinding()]
    param()
    @(
        @{ Id = 'Access'; Label = 'Access' }
        @{ Id = 'Excel'; Label = 'Excel' }
        @{ Id = 'Groove'; Label = 'Groove (legacy sync)' }
        @{ Id = 'Lync'; Label = 'Skype for Business (Lync)' }
        @{ Id = 'OneDrive'; Label = 'OneDrive' }
        @{ Id = 'OneNote'; Label = 'OneNote (Win32)' }
        @{ Id = 'Outlook'; Label = 'Outlook (classic)' }
        @{ Id = 'OutlookForWindows'; Label = 'Outlook (new)' }
        @{ Id = 'PowerPoint'; Label = 'PowerPoint' }
        @{ Id = 'Publisher'; Label = 'Publisher' }
        @{ Id = 'Teams'; Label = 'Microsoft Teams' }
        @{ Id = 'Word'; Label = 'Word' }
    )
}

function Resolve-M365AppsExcludeAppIdList {
    <#
    .SYNOPSIS
        Parses a comma-separated list of ExcludeApp IDs (case-insensitive) and returns canonical IDs.
    #>
    [CmdletBinding()]
    param(
        [string]$CommaSeparatedText
    )
    if ([string]::IsNullOrWhiteSpace($CommaSeparatedText)) { return @() }
    $valid = $script:M365AppsValidExcludeAppIds
    $seen = @{}
    foreach ($part in ($CommaSeparatedText -split ',')) {
        $t = $part.Trim()
        if (-not $t) { continue }
        $m = $valid | Where-Object { $_ -ieq $t } | Select-Object -First 1
        if (-not $m) { throw "Invalid ExcludeApp ID '$t'. Use: $($valid -join ', ')." }
        $seen[$m] = $true
    }
    return @($seen.Keys | Sort-Object)
}

function Get-M365AppsVisioProjectProductIds {
    <#
    .SYNOPSIS
        Maps a Visio/Project line (subscription vs LTSC) to ODT Product IDs. See Microsoft Learn: product IDs for Click-to-Run.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('M365Retail', 'LTSC2021Volume', 'LTSC2024Volume', 'Office2024Retail')]
        [string]$Line
    )
    switch ($Line) {
        'M365Retail' {
            return @{ Visio = 'VisioProRetail'; Project = 'ProjectProRetail' }
        }
        'LTSC2021Volume' {
            return @{ Visio = 'VisioPro2021Volume'; Project = 'ProjectPro2021Volume' }
        }
        'LTSC2024Volume' {
            return @{ Visio = 'VisioPro2024Volume'; Project = 'ProjectPro2024Volume' }
        }
        'Office2024Retail' {
            return @{ Visio = 'VisioPro2024Retail'; Project = 'ProjectPro2024Retail' }
        }
    }
}

function Get-M365AppsDefaultVisioProjectLine {
    <#
    .SYNOPSIS
        Suggested Visio/Project line for the main Office product in custom interactive XML.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProductId
    )
    switch ($ProductId) {
        'O365ProPlusRetail' { return 'M365Retail' }
        'ProPlus2021Volume' { return 'LTSC2021Volume' }
        'ProPlus2024Retail' { return 'Office2024Retail' }
        default { return 'M365Retail' }
    }
}

function Merge-M365AppsExcludeAppsIntoProducts {
    <#
    .SYNOPSIS
        Merges ExcludeApp elements into Microsoft 365 / ProPlus suite products (union with existing excludes).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [string[]]$ExcludeAppIds
    )
    if (-not $ExcludeAppIds -or $ExcludeAppIds.Count -eq 0) { return }
    $valid = $script:M365AppsValidExcludeAppIds
    $norm = @{}
    foreach ($raw in $ExcludeAppIds) {
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $t = $raw.Trim()
        $match = $valid | Where-Object { $_ -ieq $t } | Select-Object -First 1
        if (-not $match) {
            throw "Invalid ExcludeApp ID '$raw'. Use: $($valid -join ', ')."
        }
        $norm[$match] = $true
    }
    if ($norm.Count -eq 0) { return }

    [xml]$doc = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $suite = $script:M365AppsExcludeAppSuiteProductIds
    foreach ($prod in $doc.SelectNodes('//Product')) {
        $pid = $prod.GetAttribute('ID')
        if ($suite -notcontains $pid) { continue }
        $merged = @{}
        foreach ($ex in $prod.SelectNodes('ExcludeApp')) {
            $eid = $ex.GetAttribute('ID')
            if ($eid) { $merged[$eid] = $true }
        }
        foreach ($k in $norm.Keys) { $merged[$k] = $true }
        foreach ($ex in @($prod.SelectNodes('ExcludeApp'))) { [void]$prod.RemoveChild($ex) }
        foreach ($id in ($merged.Keys | Sort-Object)) {
            $el = $doc.CreateElement('ExcludeApp')
            $el.SetAttribute('ID', $id)
            [void]$prod.AppendChild($el)
        }
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $Path).Path, $doc.OuterXml, $utf8NoBom)
}

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
        ($Preset -in $script:M365AppsVisioProjectPresetNames) -or
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
        ($Preset -in $script:M365AppsVisioProjectPresetNames) -or
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
        Returns the path to configs\ next to the module (optional folder for your own XML files).
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

function Get-M365AppsOdtConfigurationDocumentationUri {
    <#
    .SYNOPSIS
        Returns the Microsoft Learn URL for Office Deployment Tool configuration options (valid elements and attributes).
    #>
    [CmdletBinding()]
    param()
    $script:M365AppsOdtConfigurationLearnUrl
}

function Assert-M365AppsOdtAddChannelValue {
    <#
    .SYNOPSIS
        Ensures a Channel value is one of the names documented for the ODT Add element.
    #>
    [CmdletBinding()]
    param(
        [string]$Channel
    )
    if ([string]::IsNullOrWhiteSpace($Channel)) { return }
    if ($script:M365AppsValidOdtAddChannelValues -notcontains $Channel) {
        throw "Invalid ODT Add Channel '$Channel'. See: $(Get-M365AppsOdtConfigurationDocumentationUri). Allowed values: $($script:M365AppsValidOdtAddChannelValues -join ', ')."
    }
}

function Test-M365AppsConfigurationXml {
    <#
    .SYNOPSIS
        Validates that text is well-formed XML with an ODT Configuration root element.
    .NOTES
        Full syntax for Product, Language, Updates, Remove, and other elements is defined by Microsoft:
        Get-M365AppsOdtConfigurationDocumentationUri
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$XmlText
    )
    if ([string]::IsNullOrWhiteSpace($XmlText)) {
        throw 'Configuration XML is empty.'
    }
    try {
        [xml]$doc = $XmlText.Trim()
    } catch {
        throw "Configuration XML is not valid XML: $($_.Exception.Message)"
    }
    if (-not $doc.Configuration) {
        throw 'Invalid ODT configuration: root element must be Configuration.'
    }
}

function Build-M365AppsUpdatesXmlLine {
    <#
    .SYNOPSIS
        Single-line ODT Updates element for Configuration (portal-style update controls).
    #>
    [CmdletBinding()]
    param(
        [bool]$Enabled = $true,
        [string]$TargetVersion,
        [string]$Deadline
    )
    $attrs = @("Enabled=`"$(if ($Enabled) { 'TRUE' } else { 'FALSE' })`"")
    if ($TargetVersion -and $TargetVersion.Trim()) {
        $attrs += "TargetVersion=`"$($TargetVersion.Trim())`""
    }
    if ($Deadline -and $Deadline.Trim()) {
        $attrs += "Deadline=`"$($Deadline.Trim())`""
    }
    "  <Updates $($attrs -join ' ') />"
}

function Merge-M365AppsLanguageIdList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PrimaryLanguageId,
        [string[]]$AdditionalLanguageIds
    )
    $list = New-Object System.Collections.Generic.List[string]
    [void]$list.Add($PrimaryLanguageId.Trim().ToLowerInvariant())
    if ($AdditionalLanguageIds) {
        foreach ($a in $AdditionalLanguageIds) {
            if ([string]::IsNullOrWhiteSpace($a)) { continue }
            $x = $a.Trim().ToLowerInvariant()
            if (-not $list.Contains($x)) { [void]$list.Add($x) }
        }
    }
    return ,$list.ToArray()
}

function New-M365AppsUninstallConfigurationXml {
    <#
    .SYNOPSIS
        Builds ODT XML to remove all Microsoft 365 Apps (Click-to-Run) per Microsoft Remove element guidance.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Full', 'None')]
        [string]$DisplayLevel = 'None'
    )
    @"
<?xml version="1.0" encoding="UTF-8"?>
<Configuration>
  <Remove All="TRUE" />
  <Display Level="$DisplayLevel" AcceptEULA="TRUE" />
</Configuration>
"@
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
        Assert-M365AppsOdtAddChannelValue -Channel $Channel
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

function Export-M365AppsConfigurationStringToPathWithOverrides {
    <#
    .SYNOPSIS
        Writes generated ODT XML to disk, then applies edition/channel/language and optional ExcludeApp merges.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ConfigurationXml,
        [Parameter(Mandatory)]
        [string]$DestinationPath,
        [ValidateSet('32', '64')]
        [string]$OfficeClientEdition = '64',
        [string]$Channel,
        [string]$LanguageId = 'en-us',
        [string[]]$ExcludeAppIds
    )
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($DestinationPath, $ConfigurationXml.Trim(), $utf8NoBom)
    Set-M365AppsConfigurationOverrides -Path $DestinationPath -OfficeClientEdition $OfficeClientEdition -Channel $Channel -LanguageId $LanguageId
    if ($ExcludeAppIds -and $ExcludeAppIds.Count -gt 0) {
        Merge-M365AppsExcludeAppsIntoProducts -Path $DestinationPath -ExcludeAppIds $ExcludeAppIds
    }
}

function Resolve-M365AppsVisioProjectBundleDefinition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet(
            'O365ProPlusVisioProject',
            'O365ProPlusVisioProject-Retail',
            'O365ProPlusVisioProject-2024',
            'O365ProPlusVisioProject-VDI',
            'O365ProPlusVisioProject-Retail-VDI',
            'O365ProPlusVisioProject-2024-VDI'
        )]
        [string]$Bundle
    )
    switch ($Bundle) {
        'O365ProPlusVisioProject' { return @{ Line = 'LTSC2021Volume'; Vdi = $false } }
        'O365ProPlusVisioProject-VDI' { return @{ Line = 'LTSC2021Volume'; Vdi = $true } }
        'O365ProPlusVisioProject-Retail' { return @{ Line = 'M365Retail'; Vdi = $false } }
        'O365ProPlusVisioProject-Retail-VDI' { return @{ Line = 'M365Retail'; Vdi = $true } }
        'O365ProPlusVisioProject-2024' { return @{ Line = 'LTSC2024Volume'; Vdi = $false } }
        'O365ProPlusVisioProject-2024-VDI' { return @{ Line = 'LTSC2024Volume'; Vdi = $true } }
        default { throw "Unknown bundle '$Bundle'." }
    }
}

function New-M365AppsVisioProjectBundleConfigurationXml {
    <#
    .SYNOPSIS
        Builds ODT XML for Microsoft 365 Apps (enterprise retail suite) plus Visio and Project (subscription or LTSC volume).
    .DESCRIPTION
        Replaces former static configs\O365ProPlusVisioProject*.xml files. Baseline suite excludes match physical vs VDI;
        pass user ExcludeApp lists through Export-M365AppsConfigurationStringToPathWithOverrides.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet(
            'O365ProPlusVisioProject',
            'O365ProPlusVisioProject-Retail',
            'O365ProPlusVisioProject-2024',
            'O365ProPlusVisioProject-VDI',
            'O365ProPlusVisioProject-Retail-VDI',
            'O365ProPlusVisioProject-2024-VDI'
        )]
        [string]$Bundle,
        [ValidateSet('32', '64')]
        [string]$OfficeClientEdition = '64',
        [string]$Channel,
        [string]$LanguageId = 'en-us',
        [ValidateSet('Full', 'None')]
        [string]$DisplayLevel = 'None'
    )
    $def = Resolve-M365AppsVisioProjectBundleDefinition -Bundle $Bundle
    $channelResolved = if ($Channel) { $Channel } else { 'MonthlyEnterprise' }
    Assert-M365AppsOdtAddChannelValue -Channel $channelResolved
    $vp = Get-M365AppsVisioProjectProductIds -Line $def.Line
    $suiteEx = @(Get-M365AppsO365RetailBaseExcludeAppIds -Vdi:($def.Vdi))
    $excLines = @()
    foreach ($id in ($suiteEx | Sort-Object)) {
        $excLines += "      <ExcludeApp ID=`"$id`" />"
    }
    $excBlock = if ($excLines.Count) { "`n$($excLines -join "`n")" } else { '' }
    $vdiProps = if ($def.Vdi) { "`n  <Property Name=`"SharedComputerLicensing`" Value=`"1`" />" } else { '' }
    @"
<?xml version="1.0" encoding="UTF-8"?>
<Configuration ID="$Bundle">
  <Add OfficeClientEdition="$OfficeClientEdition" Channel="$channelResolved">
    <Product ID="O365ProPlusRetail">
      <Language ID="$LanguageId" />$excBlock
    </Product>
    <Product ID="$($vp.Visio)">
      <Language ID="$LanguageId" />
    </Product>
    <Product ID="$($vp.Project)">
      <Language ID="$LanguageId" />
    </Product>
  </Add>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />$vdiProps
  <Display Level="$DisplayLevel" AcceptEULA="TRUE" />
</Configuration>
"@
}

function Add-M365AppsOptionalVisioProjectProducts {
    <#
    .SYNOPSIS
        Appends Visio and/or Project Product elements to an existing ODT configuration (e.g. after copying a base preset).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$LanguageId,
        [switch]$IncludeVisio,
        [switch]$IncludeProject,
        [Parameter(Mandatory)]
        [ValidateSet('M365Retail', 'LTSC2021Volume', 'LTSC2024Volume', 'Office2024Retail')]
        [string]$VisioProjectLine,
        [string[]]$AdditionalLanguageIds = @()
    )
    if (-not $IncludeVisio -and -not $IncludeProject) { return }
    $vp = Get-M365AppsVisioProjectProductIds -Line $VisioProjectLine
    $allLangs = Merge-M365AppsLanguageIdList -PrimaryLanguageId $LanguageId -AdditionalLanguageIds $AdditionalLanguageIds
    [xml]$doc = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $add = $doc.Configuration.Add
    if (-not $add) { throw 'Invalid ODT XML: Configuration/Add missing; cannot append Visio/Project.' }
    $existing = @{}
    foreach ($p in $add.SelectNodes('Product')) {
        $id = $p.GetAttribute('ID')
        if ($id) { $existing[$id.ToLowerInvariant()] = $true }
    }
    function New-ProductNode {
        param([string]$ProductId)
        $prod = $doc.CreateElement('Product')
        $prod.SetAttribute('ID', $ProductId)
        foreach ($lid in $allLangs) {
            $lang = $doc.CreateElement('Language')
            $lang.SetAttribute('ID', $lid)
            [void]$prod.AppendChild($lang)
        }
        return $prod
    }
    if ($IncludeVisio -and -not $existing.ContainsKey($vp.Visio.ToLowerInvariant())) {
        [void]$add.AppendChild((New-ProductNode -ProductId $vp.Visio))
    }
    if ($IncludeProject -and -not $existing.ContainsKey($vp.Project.ToLowerInvariant())) {
        [void]$add.AppendChild((New-ProductNode -ProductId $vp.Project))
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $Path).Path, $doc.OuterXml, $utf8NoBom)
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
    .PARAMETER AddOnsOnly
        Omit the primary Office suite; install only Visio and/or Project (at least one required).
    #>
    [CmdletBinding(DefaultParameterSetName = 'Suite')]
    param(
        [Parameter(ParameterSetName = 'Suite', Mandatory)]
        [string]$ProductId,
        [Parameter(ParameterSetName = 'AddOnsOnly', Mandatory)]
        [switch]$AddOnsOnly,
        [string]$LanguageId = 'en-us',
        [ValidateSet('32', '64')]
        [string]$OfficeClientEdition = '64',
        [string]$Channel = 'Current',
        [ValidateSet('Full', 'None')]
        [string]$DisplayLevel = 'Full',
        [switch]$IncludeVisio,
        [switch]$IncludeProject,
        [Parameter()]
        [ValidateSet('M365Retail', 'LTSC2021Volume', 'LTSC2024Volume', 'Office2024Retail')]
        [string]$VisioProjectLine,
        [switch]$AutoActivate,
        [string[]]$ExcludeAppIds,
        [string[]]$AdditionalLanguageIds = @(),
        [bool]$UpdatesEnabled = $true,
        [string]$UpdatesTargetVersion,
        [string]$UpdatesDeadline,
        [switch]$SharedComputerLicensing
    )
    if ($PSCmdlet.ParameterSetName -eq 'AddOnsOnly') {
        if (-not $IncludeVisio -and -not $IncludeProject) {
            throw 'AddOnsOnly requires -IncludeVisio and/or -IncludeProject.'
        }
    }
    if ($SharedComputerLicensing -and $PSCmdlet.ParameterSetName -eq 'Suite' -and $ProductId -ne 'O365ProPlusRetail') {
        throw 'Shared computer licensing applies to Microsoft 365 Apps (Click-to-Run) suite (O365ProPlusRetail) in this tool.'
    }
    Assert-M365AppsOdtAddChannelValue -Channel $Channel
    $validEx = $script:M365AppsValidExcludeAppIds
    $mergedEx = @{}
    if ($PSCmdlet.ParameterSetName -eq 'Suite' -and $ProductId -eq 'O365ProPlusRetail') {
        $baseEx = if ($SharedComputerLicensing) {
            @(Get-M365AppsO365RetailBaseExcludeAppIds -Vdi)
        } else {
            @(Get-M365AppsO365RetailBaseExcludeAppIds)
        }
        foreach ($b in $baseEx) { $mergedEx[$b] = $true }
    }
    if ($PSCmdlet.ParameterSetName -eq 'Suite' -and $ExcludeAppIds) {
        foreach ($x in $ExcludeAppIds) {
            if ([string]::IsNullOrWhiteSpace($x)) { continue }
            $m = $validEx | Where-Object { $_ -ieq $x.Trim() } | Select-Object -First 1
            if (-not $m) { throw "Invalid ExcludeApp ID '$x'. Use: $($validEx -join ', ')." }
            $mergedEx[$m] = $true
        }
    }
    $excLines = @()
    foreach ($id in ($mergedEx.Keys | Sort-Object)) {
        $excLines += "    <ExcludeApp ID=`"$id`" />"
    }
    $excBlock = if ($excLines.Count) { "`n$($excLines -join "`n")" } else { '' }
    $langIds = Merge-M365AppsLanguageIdList -PrimaryLanguageId $LanguageId -AdditionalLanguageIds $AdditionalLanguageIds
    $langInner = ($langIds | ForEach-Object { "  <Language ID=`"$_`" />" }) -join "`n"
    $products = @()
    if ($PSCmdlet.ParameterSetName -eq 'Suite') {
        $products += "<Product ID='$ProductId'>`n$langInner$excBlock`n</Product>"
    }
    $vpIds = $null
    if ($IncludeVisio -or $IncludeProject) {
        $line = if ($PSCmdlet.ParameterSetName -eq 'AddOnsOnly') {
            if ($PSBoundParameters.ContainsKey('VisioProjectLine')) { $VisioProjectLine } else { 'M365Retail' }
        } elseif ($PSBoundParameters.ContainsKey('VisioProjectLine')) {
            $VisioProjectLine
        } else {
            Get-M365AppsDefaultVisioProjectLine -ProductId $ProductId
        }
        if (-not $line) { $line = 'M365Retail' }
        $vpIds = Get-M365AppsVisioProjectProductIds -Line $line
    }
    if ($IncludeVisio) {
        $products += "<Product ID='$($vpIds.Visio)'>`n$langInner`n</Product>"
    }
    if ($IncludeProject) {
        $products += "<Product ID='$($vpIds.Project)'>`n$langInner`n</Product>"
    }

    $props = @()
    if ($PSCmdlet.ParameterSetName -eq 'Suite') {
        $props += '  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />'
        if ($SharedComputerLicensing) {
            $props += '  <Property Name="SharedComputerLicensing" Value="1" />'
        }
    }
    if ($AutoActivate -and $PSCmdlet.ParameterSetName -eq 'Suite') {
        $props += '  <Property Name="AUTOACTIVATE" Value="1" />'
    }
    $propsBlock = if ($props.Count) { "`n$($props -join "`n")" } else { '' }
    $updatesLine = Build-M365AppsUpdatesXmlLine -Enabled:$UpdatesEnabled -TargetVersion $UpdatesTargetVersion -Deadline $UpdatesDeadline

    @"
<Configuration>
  <Add OfficeClientEdition="$OfficeClientEdition" Channel="$Channel">
    $($products -join "`n    ")
  </Add>
$updatesLine
  <Display Level="$DisplayLevel" AcceptEULA="TRUE" />$propsBlock
</Configuration>
"@
}

function Get-M365AppsO365RetailBaseExcludeAppIds {
    <#
    .SYNOPSIS
        Built-in ExcludeApp IDs for Microsoft 365 Apps (retail) profiles before user-selected excludes are merged.
    #>
    [CmdletBinding()]
    param(
        [switch]$Vdi
    )
    if ($Vdi) {
        return @('Groove', 'Lync', 'OneDrive', 'OutlookForWindows', 'Teams')
    }
    return @('OutlookForWindows')
}

function New-M365AppsO365Configuration {
    <#
    .SYNOPSIS
        Builds ODT XML for Microsoft 365 Apps retail suite (O365ProPlusRetail or O365BusinessRetail).
    .DESCRIPTION
        Enterprise vs business differ only by product ID and default update channel. VDI/shared PC adds
        SharedComputerLicensing and standard VDI excludes; physical profiles add a minimal base exclude.
        AdditionalExcludeAppIds are unioned with the base list and validated against ODT ExcludeApp IDs.
        AdditionalLanguageIds add extra Language elements (same as Microsoft 365 admin center / deployment settings).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Enterprise', 'Business')]
        [string]$O365Sku,
        [switch]$Vdi,
        [ValidateSet('32', '64')]
        [string]$OfficeClientEdition = '64',
        [string]$Channel,
        [string]$LanguageId = 'en-us',
        [string[]]$AdditionalLanguageIds = @(),
        [ValidateSet('Full', 'None')]
        [string]$DisplayLevel = 'None',
        [string[]]$AdditionalExcludeAppIds,
        [bool]$UpdatesEnabled = $true,
        [string]$UpdatesTargetVersion,
        [string]$UpdatesDeadline
    )
    $productId = if ($O365Sku -eq 'Enterprise') { 'O365ProPlusRetail' } else { 'O365BusinessRetail' }
    $resolvedChannel = $Channel
    if (-not $resolvedChannel) {
        $resolvedChannel = if ($O365Sku -eq 'Enterprise') { 'MonthlyEnterprise' } else { 'Current' }
    } else {
        Assert-M365AppsOdtAddChannelValue -Channel $resolvedChannel
    }
    $configId = if ($O365Sku -eq 'Enterprise') {
        if ($Vdi) { 'EnterpriseVDI' } else { 'EnterprisePhysical' }
    } else {
        if ($Vdi) { 'BusinessVDI' } else { 'BusinessPhysical' }
    }
    $baseEx = @(Get-M365AppsO365RetailBaseExcludeAppIds -Vdi:$Vdi)
    $validEx = $script:M365AppsValidExcludeAppIds
    $merged = @{}
    foreach ($x in $baseEx) { $merged[$x] = $true }
    if ($AdditionalExcludeAppIds) {
        foreach ($raw in $AdditionalExcludeAppIds) {
            if ([string]::IsNullOrWhiteSpace($raw)) { continue }
            $m = $validEx | Where-Object { $_ -ieq $raw.Trim() } | Select-Object -First 1
            if (-not $m) { throw "Invalid ExcludeApp ID '$raw'. Use: $($validEx -join ', ')." }
            $merged[$m] = $true
        }
    }
    $excLines = @()
    foreach ($id in ($merged.Keys | Sort-Object)) {
        $excLines += "      <ExcludeApp ID=`"$id`" />"
    }
    $excBlock = if ($excLines.Count) { "`n$($excLines -join "`n")" } else { '' }
    $vdiProps = if ($Vdi) { "`n  <Property Name=`"SharedComputerLicensing`" Value=`"1`" />" } else { '' }
    $langIds = Merge-M365AppsLanguageIdList -PrimaryLanguageId $LanguageId -AdditionalLanguageIds $AdditionalLanguageIds
    $langLines = @()
    foreach ($lid in $langIds) {
        $langLines += "      <Language ID=`"$lid`" />"
    }
    $langBlock = $langLines -join "`n"
    $updatesLine = Build-M365AppsUpdatesXmlLine -Enabled:$UpdatesEnabled -TargetVersion $UpdatesTargetVersion -Deadline $UpdatesDeadline
    @"
<?xml version="1.0" encoding="UTF-8"?>
<Configuration ID="$configId">
  <Add OfficeClientEdition="$OfficeClientEdition" Channel="$resolvedChannel">
    <Product ID="$productId">
$langBlock$excBlock
    </Product>
  </Add>
$updatesLine
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />$vdiProps
  <Display Level="$DisplayLevel" AcceptEULA="TRUE" />
</Configuration>
"@
}

function New-M365AppsO365ConfigurationForRetailProfile {
    <#
    .SYNOPSIS
        Same as New-M365AppsO365Configuration using the four retail deployment profiles (enterprise/business x physical/VDI).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('EnterprisePhysical', 'EnterpriseVDI', 'BusinessPhysical', 'BusinessVDI')]
        [string]$RetailProfile,
        [ValidateSet('32', '64')]
        [string]$OfficeClientEdition = '64',
        [string]$Channel,
        [string]$LanguageId = 'en-us',
        [string[]]$AdditionalLanguageIds = @(),
        [ValidateSet('Full', 'None')]
        [string]$DisplayLevel = 'None',
        [string[]]$AdditionalExcludeAppIds,
        [bool]$UpdatesEnabled = $true,
        [string]$UpdatesTargetVersion,
        [string]$UpdatesDeadline
    )
    $sku = 'Enterprise'
    $vdi = $false
    switch ($RetailProfile) {
        'EnterprisePhysical' { $sku = 'Enterprise'; $vdi = $false }
        'EnterpriseVDI' { $sku = 'Enterprise'; $vdi = $true }
        'BusinessPhysical' { $sku = 'Business'; $vdi = $false }
        'BusinessVDI' { $sku = 'Business'; $vdi = $true }
    }
    New-M365AppsO365Configuration -O365Sku $sku -Vdi:$vdi -OfficeClientEdition $OfficeClientEdition `
        -Channel $Channel -LanguageId $LanguageId -AdditionalLanguageIds $AdditionalLanguageIds -DisplayLevel $DisplayLevel `
        -AdditionalExcludeAppIds $AdditionalExcludeAppIds -UpdatesEnabled:$UpdatesEnabled `
        -UpdatesTargetVersion $UpdatesTargetVersion -UpdatesDeadline $UpdatesDeadline
}

function Resolve-M365AppsLegacyO365PresetNameToRetailProfile {
    <#
    .SYNOPSIS
        Maps deprecated O365*.xml preset names to RetailProfile values (for older scripts and env vars).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)
    $t = $Name.Trim()
    switch ($t) {
        'O365ProPlus' { return 'EnterprisePhysical' }
        'O365ProPlus-VDI' { return 'EnterpriseVDI' }
        'O365Business' { return 'BusinessPhysical' }
        'O365Business-VDI' { return 'BusinessVDI' }
        default { return $null }
    }
}
