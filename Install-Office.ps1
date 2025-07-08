# Office Auto Installer - Enhanced UI Version
# Downloads and installs Microsoft Office through official channels
# No licensing modifications - uses Microsoft's official deployment tools

# Pause at start for user readiness
Pause

# Setup Paths
$installerFolder = "$PSScriptRoot\OfficeInstaller"

# Auto-elevate if not running as admin
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Host "⚠️  Script is not running as Administrator. Attempting to relaunch..." -ForegroundColor Yellow
    try {
        Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    } catch {
        Write-Host "❌ Failed to relaunch script with admin rights. Error: $_" -ForegroundColor Red
        Pause
    }
    Exit
}

# Clean the folder at the start to avoid old/corrupt files
if (Test-Path $installerFolder) {
    Remove-Item -Path "$installerFolder\*" -Recurse -Force -ErrorAction SilentlyContinue
} else {
    New-Item -ItemType Directory -Path $installerFolder | Out-Null
}
Set-Location -Path $installerFolder

# Logging setup
$logFile = "$installerFolder\installer.log"

function Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -Append -FilePath $logFile
}

function Show-Header {
    Clear-Host
    $width = 80
    $border = "═" * $width
    $title = "MICROSOFT OFFICE AUTO INSTALLER"
    $subtitle = "Official Microsoft Office Deployment Tool Interface"
    $version = "v2.0 - Enhanced UI"
    
    Write-Host "╔$border╗" -ForegroundColor Cyan
    Write-Host "║" -ForegroundColor Cyan -NoNewline
    Write-Host $title.PadLeft(($width + $title.Length) / 2).PadRight($width) -ForegroundColor White -NoNewline
    Write-Host "║" -ForegroundColor Cyan
    Write-Host "║" -ForegroundColor Cyan -NoNewline
    Write-Host $subtitle.PadLeft(($width + $subtitle.Length) / 2).PadRight($width) -ForegroundColor Gray -NoNewline
    Write-Host "║" -ForegroundColor Cyan
    Write-Host "║" -ForegroundColor Cyan -NoNewline
    Write-Host $version.PadLeft(($width + $version.Length) / 2).PadRight($width) -ForegroundColor DarkGray -NoNewline
    Write-Host "║" -ForegroundColor Cyan
    Write-Host "╚$border╝" -ForegroundColor Cyan
    Write-Host ""
}

function Show-Progress {
    param(
        [string]$Activity,
        [int]$PercentComplete = 0,
        [string]$Status = "Processing..."
    )
    
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
}

function Show-MenuOption {
    param(
        [string]$Number,
        [string]$Title,
        [string]$Description = "",
        [string]$Color = "White"
    )
    
    Write-Host "  " -NoNewline
    Write-Host "[$Number]" -ForegroundColor Cyan -NoNewline
    Write-Host " $Title" -ForegroundColor $Color
    if ($Description) {
        Write-Host "      $Description" -ForegroundColor DarkGray
    }
}

function Get-UserChoice {
    param(
        [string]$Prompt,
        [string[]]$ValidChoices,
        [string]$DefaultChoice = $ValidChoices[0]
    )
    
    do {
        Write-Host ""
        Write-Host "➤ " -ForegroundColor Green -NoNewline
        Write-Host $Prompt -ForegroundColor Yellow -NoNewline
        Write-Host " [Default: $DefaultChoice]: " -ForegroundColor DarkGray -NoNewline
        $choice = Read-Host
        
        if ([string]::IsNullOrWhiteSpace($choice)) {
            $choice = $DefaultChoice
        }
        
        if ($ValidChoices -contains $choice) {
            return $choice
        } else {
            Write-Host "❌ Invalid choice. Please select from: $($ValidChoices -join ', ')" -ForegroundColor Red
        }
    } while ($true)
}

function Fix-SystemPath {
    Log "Checking and fixing essential system PATH entries..."
    Show-Progress -Activity "System Check" -PercentComplete 25 -Status "Validating system PATH..."

    $requiredPaths = @(
        "C:\Windows\System32",
        "C:\Windows",
        "C:\Windows\System32\Wbem",
        "C:\Windows\System32\WindowsPowerShell\v1.0\"
    )

    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    $currentPath = (Get-ItemProperty -Path $regPath -Name Path).Path
    $pathArray = $currentPath -split ";"

    $added = 0
    foreach ($p in $requiredPaths) {
        if (-not ($pathArray -contains $p)) {
            $pathArray += $p
            Log "Added missing path: $p"
            $added++
        } else {
            Log "Path already present: $p"
        }
    }

    if ($added -gt 0) {
        $newPath = ($pathArray | Where-Object { $_ -ne "" }) -join ";"
        Set-ItemProperty -Path $regPath -Name Path -Value $newPath
        Log "System PATH updated successfully."
        Write-Host "✅ System PATH updated with $added missing entries." -ForegroundColor Green
    } else {
        Log "No changes made to PATH."
        Write-Host "✅ System PATH is properly configured." -ForegroundColor Green
    }
}

function Check-Internet {
    Show-Progress -Activity "System Check" -PercentComplete 50 -Status "Testing internet connectivity..."
    Write-Host "🌐 Checking internet connection..." -ForegroundColor Blue
    
    Try {
        $null = Invoke-WebRequest -Uri "https://www.microsoft.com" -UseBasicParsing -TimeoutSec 10
        Log "Internet connection verified."
        Write-Host "✅ Internet connection verified." -ForegroundColor Green
    } Catch {
        Log "Internet connection failed. Exiting script."
        Write-Host "❌ No internet connection detected. Please connect and try again." -ForegroundColor Red
        Pause
        Exit 1
    }
}

function Show-ConfigurationMenu {
    Show-Header
    Write-Host "📋 OFFICE CONFIGURATION SETUP" -ForegroundColor Yellow
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    
    Log "Prompting user for configuration selections."

    # Architecture Selection
    Write-Host "🏗️  SELECT ARCHITECTURE:" -ForegroundColor Magenta
    Show-MenuOption "1" "64-bit (Recommended)" "Best performance for modern systems"
    Show-MenuOption "2" "32-bit" "For older systems or specific compatibility needs"
    $arch = Get-UserChoice "Choose architecture" @("1", "2") "1"
    $bit = if ($arch -eq "2") { "32" } else { "64" }

    Write-Host ""
    Write-Host "📦 SELECT OFFICE EDITION:" -ForegroundColor Magenta
    Show-MenuOption "1" "Office 2024 Pro Plus" "Latest version with all features"
    Show-MenuOption "2" "Office LTSC 2021" "Long-term support channel"
    Show-MenuOption "3" "Microsoft 365 Apps" "Cloud-connected productivity suite"
    $editionChoice = Get-UserChoice "Choose edition" @("1", "2", "3") "1"
    
    $editionMap = @{ 
        "1" = @{ID = "ProPlus2024Retail"; Name = "Office 2024 Pro Plus"}
        "2" = @{ID = "ProPlus2021Volume"; Name = "Office LTSC 2021"}
        "3" = @{ID = "O365ProPlusRetail"; Name = "Microsoft 365 Apps"}
    }
    $edition = $editionMap[$editionChoice]

    Write-Host ""
    Write-Host "🎨 ADDITIONAL COMPONENTS:" -ForegroundColor Magenta
    Show-MenuOption "1" "Include Visio" "Diagramming and vector graphics"
    Show-MenuOption "2" "Skip Visio"
    $visio = Get-UserChoice "Include Visio?" @("1", "2") "2"

    Show-MenuOption "1" "Include Project" "Project management tools"
    Show-MenuOption "2" "Skip Project"
    $project = Get-UserChoice "Include Project?" @("1", "2") "2"

    Write-Host ""
    Write-Host "🔄 UPDATE CHANNEL:" -ForegroundColor Magenta
    Show-MenuOption "1" "Monthly Channel" "Latest features and updates"
    Show-MenuOption "2" "Semi-Annual Channel" "Stable, tested updates"
    $channelChoice = Get-UserChoice "Choose update channel" @("1", "2") "1"
    $channel = if ($channelChoice -eq "2") { "Broad" } else { "Current" }

    Write-Host ""
    Write-Host "🌍 LANGUAGE SELECTION:" -ForegroundColor Magenta
    Show-MenuOption "1" "English (United States)" "en-us"
    Show-MenuOption "2" "English (United Kingdom)" "en-gb"
    Show-MenuOption "3" "French (France)" "fr-fr"
    Show-MenuOption "4" "German (Germany)" "de-de"
    Show-MenuOption "5" "Dutch (Netherlands)" "nl-nl"
    Show-MenuOption "6" "Spanish (Spain)" "es-es"
    Show-MenuOption "7" "Portuguese (Brazil)" "pt-br"
    $langChoice = Get-UserChoice "Choose language" @("1", "2", "3", "4", "5", "6", "7") "1"

    $languageMap = @{
        "1" = @{Code = "en-us"; Name = "English (United States)"}
        "2" = @{Code = "en-gb"; Name = "English (United Kingdom)"}
        "3" = @{Code = "fr-fr"; Name = "French (France)"}
        "4" = @{Code = "de-de"; Name = "German (Germany)"}
        "5" = @{Code = "nl-nl"; Name = "Dutch (Netherlands)"}
        "6" = @{Code = "es-es"; Name = "Spanish (Spain)"}
        "7" = @{Code = "pt-br"; Name = "Portuguese (Brazil)"}
    }
    $language = $languageMap[$langChoice]

    Write-Host ""
    Write-Host "🖥️  INSTALLATION INTERFACE:" -ForegroundColor Magenta
    Show-MenuOption "1" "Show installation progress" "Display Office setup window"
    Show-MenuOption "2" "Silent installation" "Install in background"
    $uiChoice = Get-UserChoice "Display installation UI?" @("1", "2") "1"
    $uiLevel = if ($uiChoice -eq "1") { "Full" } else { "None" }

    return @{ 
        bit = $bit
        visio = $visio
        project = $project
        channel = $channel
        language = $language.Code
        languageName = $language.Name
        ui = $uiLevel
        edition = $edition.ID
        editionName = $edition.Name
    }
}

function Show-ConfigurationSummary($options) {
    Show-Header
    Write-Host "📋 INSTALLATION SUMMARY" -ForegroundColor Yellow
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    
    Write-Host "Edition:      " -ForegroundColor Cyan -NoNewline
    Write-Host $options.editionName -ForegroundColor White
    
    Write-Host "Architecture: " -ForegroundColor Cyan -NoNewline
    Write-Host "$($options.bit)-bit" -ForegroundColor White
    
    Write-Host "Language:     " -ForegroundColor Cyan -NoNewline
    Write-Host $options.languageName -ForegroundColor White
    
    Write-Host "Channel:      " -ForegroundColor Cyan -NoNewline
    Write-Host $options.channel -ForegroundColor White
    
    Write-Host "Visio:        " -ForegroundColor Cyan -NoNewline
    Write-Host $(if ($options.visio -eq "1") { "Yes" } else { "No" }) -ForegroundColor White
    
    Write-Host "Project:      " -ForegroundColor Cyan -NoNewline
    Write-Host $(if ($options.project -eq "1") { "Yes" } else { "No" }) -ForegroundColor White
    
    Write-Host "UI Mode:      " -ForegroundColor Cyan -NoNewline
    Write-Host $options.ui -ForegroundColor White
    
    Write-Host ""
    Write-Host "Press any key to continue with installation..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Download-ODT {
    $url = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    $output = "$installerFolder\setup.exe"

    Show-Header
    Write-Host "📥 DOWNLOADING OFFICE DEPLOYMENT TOOL" -ForegroundColor Yellow
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    
    Log "Downloading Office Deployment Tool from $url..."
    Write-Host "🔗 Source: " -ForegroundColor Cyan -NoNewline
    Write-Host $url -ForegroundColor White
    Write-Host "📁 Destination: " -ForegroundColor Cyan -NoNewline
    Write-Host $output -ForegroundColor White
    Write-Host ""
    
    Show-Progress -Activity "Download" -PercentComplete 0 -Status "Starting download..."
    
    try {
        Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing
        Show-Progress -Activity "Download" -PercentComplete 100 -Status "Download complete"
    } catch {
        Log "Download failed: $_"
        Write-Host "❌ Download failed: $_" -ForegroundColor Red
        Pause
        Exit 1
    }

    if (-Not (Test-Path $output) -or ((Get-Item $output).Length -lt 100000)) {
        Log "Downloaded file appears to be corrupted or incomplete."
        Write-Host "❌ Downloaded file is corrupted or incomplete. Please try again." -ForegroundColor Red
        Pause
        Exit 1
    }

    Log "Office Deployment Tool downloaded successfully."
    Write-Host "✅ Office Deployment Tool downloaded successfully!" -ForegroundColor Green
    Start-Sleep -Seconds 2
}

function Generate-Config($options) {
    Show-Header
    Write-Host "⚙️  GENERATING CONFIGURATION" -ForegroundColor Yellow
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    
    Log "Generating config.xml with selected options..."
    Show-Progress -Activity "Configuration" -PercentComplete 50 -Status "Creating XML configuration..."

    $products = @()
    $products += "<Product ID='" + $options.edition + "'>`n  <Language ID='" + $options.language + "' />`n</Product>"

    if ($options.visio -eq "1") {
        $products += "<Product ID='VisioPro2021Volume'>`n  <Language ID='" + $options.language + "' />`n</Product>"
        Write-Host "📊 Adding Visio Professional..." -ForegroundColor Green
    }
    if ($options.project -eq "1") {
        $products += "<Product ID='ProjectPro2021Volume'>`n  <Language ID='" + $options.language + "' />`n</Product>"
        Write-Host "📋 Adding Project Professional..." -ForegroundColor Green
    }

    $xmlContent = @"
<Configuration>
  <Add OfficeClientEdition="${($options.bit)}" Channel="${($options.channel)}">
    $($products -join "`n    ")
  </Add>
  <Display Level="${($options.ui)}" AcceptEULA="TRUE" />
</Configuration>
"@

    $configPath = "$installerFolder\config.xml"
    $xmlContent | Out-File -FilePath $configPath -Encoding UTF8
    Log "config.xml generated at $configPath"
    
    Show-Progress -Activity "Configuration" -PercentComplete 100 -Status "Configuration complete"
    Write-Host "✅ Configuration file created successfully!" -ForegroundColor Green
    Start-Sleep -Seconds 2
}

function Install-Office {
    Show-Header
    Write-Host "🚀 INSTALLING MICROSOFT OFFICE" -ForegroundColor Yellow
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    
    Log "Starting Office installation..."
    $setupExe = "$installerFolder\setup.exe"

    if (-Not (Test-Path $setupExe)) {
        Log "ERROR: setup.exe not found."
        Write-Host "❌ Error: setup.exe not found in $installerFolder." -ForegroundColor Red
        Pause
        Exit 1
    }

    Write-Host "🔧 Starting Office installation process..." -ForegroundColor Blue
    Write-Host "⏳ This may take several minutes depending on your internet speed..." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "📝 Note: Do not close this window during installation!" -ForegroundColor Red
    Write-Host ""

    Show-Progress -Activity "Installation" -PercentComplete 0 -Status "Launching Office installer..."
    
    try {
        Start-Process -FilePath $setupExe -ArgumentList "/configure config.xml" -Wait
        Log "Office installation process finished."
        Write-Host "✅ Office installation completed successfully!" -ForegroundColor Green
    } catch {
        Log "Installation failed: $_"
        Write-Host "❌ Installation failed: $_" -ForegroundColor Red
        Pause
        Exit 1
    }
}

function Show-CompletionSummary($options) {
    Show-Header
    Write-Host "🎉 INSTALLATION COMPLETE!" -ForegroundColor Green
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    
    Write-Host "📦 Installed Configuration:" -ForegroundColor Cyan
    Write-Host "  • Edition: " -ForegroundColor White -NoNewline
    Write-Host $options.editionName -ForegroundColor Yellow
    Write-Host "  • Architecture: " -ForegroundColor White -NoNewline
    Write-Host "$($options.bit)-bit" -ForegroundColor Yellow
    Write-Host "  • Language: " -ForegroundColor White -NoNewline
    Write-Host $options.languageName -ForegroundColor Yellow
    Write-Host "  • Update Channel: " -ForegroundColor White -NoNewline
    Write-Host $options.channel -ForegroundColor Yellow
    Write-Host "  • Visio: " -ForegroundColor White -NoNewline
    Write-Host $(if ($options.visio -eq "1") { "Included" } else { "Not included" }) -ForegroundColor Yellow
    Write-Host "  • Project: " -ForegroundColor White -NoNewline
    Write-Host $(if ($options.project -eq "1") { "Included" } else { "Not included" }) -ForegroundColor Yellow
    Write-Host "  • UI Level: " -ForegroundColor White -NoNewline
    Write-Host $options.ui -ForegroundColor Yellow
    
    Write-Host ""
    Write-Host "📁 Installation Files:" -ForegroundColor Cyan
    Write-Host "  • Installer folder: " -ForegroundColor White -NoNewline
    Write-Host $installerFolder -ForegroundColor Yellow
    Write-Host "  • Installation log: " -ForegroundColor White -NoNewline
    Write-Host $logFile -ForegroundColor Yellow
    
    Write-Host ""
    Write-Host "🔍 Next Steps:" -ForegroundColor Cyan
    Write-Host "  • Look for Office applications in your Start Menu" -ForegroundColor White
    Write-Host "  • First launch may require Microsoft account sign-in" -ForegroundColor White
    Write-Host "  • Check Windows Updates for the latest Office patches" -ForegroundColor White
    
    Write-Host ""
    Write-Host "⚠️  Important: This installer uses Microsoft's official deployment tools." -ForegroundColor Yellow
    Write-Host "    Ensure you have proper licensing for the installed Office edition." -ForegroundColor Yellow
    
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ==== Main Execution Flow ====

try {
    Log "=== Enhanced Office Installer Started ==="
    
    Show-Header
    Write-Host "🔧 SYSTEM PREPARATION" -ForegroundColor Yellow
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    
    Fix-SystemPath
    Check-Internet
    
    Write-Host ""
    Write-Host "✅ System checks completed successfully!" -ForegroundColor Green
    Write-Host "Press any key to continue to configuration..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    
    $options = Show-ConfigurationMenu
    Show-ConfigurationSummary -options $options
    Download-ODT
    Generate-Config -options $options
    Install-Office
    Show-CompletionSummary -options $options
    
    Log "=== Enhanced Office Installer Completed Successfully ==="
    
} catch {
    Log "FATAL ERROR: $_"
    Write-Host ""
    Write-Host "❌ FATAL ERROR OCCURRED" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Yellow
    Write-Host "Check the log file for details: $logFile" -ForegroundColor Gray
    Pause
    Exit 1
}