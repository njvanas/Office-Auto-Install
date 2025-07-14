# Office Auto Installer - Enhanced UI Version
# Downloads and installs Microsoft Office through official channels
# No licensing modifications - uses Microsoft's official deployment tools

# ==== UNIVERSAL PAUSE PROTECTION ====
# This ensures the window NEVER closes automatically, regardless of how it's run
param()

# Set window title immediately
$Host.UI.RawUI.WindowTitle = "Microsoft Office Auto Installer - Loading..."

# Create a global flag to track if we should pause
$global:ShouldPauseOnExit = $true

# Override the exit function globally to always pause
function global:Exit-WithPause {
    param([int]$ExitCode = 0)
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor DarkGray
    Write-Host "Press any key to close this window..." -ForegroundColor Yellow -BackgroundColor DarkBlue
    Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor DarkGray
    try {
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {
        # Fallback if ReadKey fails
        Read-Host "Press Enter to continue"
    }
    [Environment]::Exit($ExitCode)
}

# Replace the built-in exit with our pause version
function global:Exit { 
    param([int]$ExitCode = 0)
    Exit-WithPause -ExitCode $ExitCode 
}

# Also handle script termination
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    if ($global:ShouldPauseOnExit) {
        Write-Host ""
        Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor DarkGray
        Write-Host "Script is closing - Press any key to close this window..." -ForegroundColor Yellow -BackgroundColor DarkBlue
        Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor DarkGray
        try {
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        } catch {
            Start-Sleep -Seconds 3
        }
    }
}

# ==== EXECUTION POLICY FIX ====
# This section ensures the script can run regardless of PowerShell execution policy

# Check if we need to bypass execution policy
try {
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
    if ($currentPolicy -eq 'Restricted' -or $currentPolicy -eq 'AllSigned') {
        Write-Host "🔧 Fixing PowerShell execution policy..." -ForegroundColor Yellow
        try {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
            Write-Host "✅ Execution policy updated successfully!" -ForegroundColor Green
        } catch {
            Write-Host "⚠️  Could not update execution policy automatically." -ForegroundColor Yellow
            Write-Host "   This script will still work, but you may see warnings." -ForegroundColor Gray
        }
    }
} catch {
    # If execution policy check fails, continue anyway
    Write-Host "⚠️  Could not check execution policy, continuing anyway..." -ForegroundColor Yellow
}

# Update window title
$Host.UI.RawUI.WindowTitle = "Microsoft Office Auto Installer - Ready"

# Welcome message and admin check
function Show-WelcomeScreen {
    Clear-Host
    $width = 80
    $border = "═" * $width
    $title = "MICROSOFT OFFICE AUTO INSTALLER"
    $subtitle = "Easy Office Installation for Everyone"
    $version = "v3.2 - Universal Pause Protection"
    
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
    
    Write-Host "👋 Welcome! This tool will help you install Microsoft Office easily." -ForegroundColor Green
    Write-Host "   No technical knowledge required - just follow the simple prompts!" -ForegroundColor Gray
    Write-Host ""
    
    # Show execution method
    Write-Host "🖥️  Execution Method: " -ForegroundColor Blue -NoNewline
    Write-Host "PowerShell Script" -ForegroundColor White
    Write-Host "🔒 Window Protection: " -ForegroundColor Blue -NoNewline
    Write-Host "Enabled (window will not auto-close)" -ForegroundColor Green
    Write-Host ""
    
    # Check if running as admin
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    
    if (-not $isAdmin) {
        Write-Host "⚠️  ADMINISTRATOR PRIVILEGES REQUIRED" -ForegroundColor Yellow
        Write-Host "═" * 50 -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "📋 Why do we need admin rights?" -ForegroundColor Cyan
        Write-Host "   • Office installation requires system-level access" -ForegroundColor White
        Write-Host "   • Ensures proper integration with Windows" -ForegroundColor White
        Write-Host "   • Prevents installation errors and conflicts" -ForegroundColor White
        Write-Host ""
        Write-Host "🔒 What happens next?" -ForegroundColor Cyan
        Write-Host "   • Windows will ask for permission (UAC prompt)" -ForegroundColor White
        Write-Host "   • Click 'Yes' to continue with installation" -ForegroundColor White
        Write-Host "   • The script will restart with proper privileges" -ForegroundColor White
        Write-Host ""
        Write-Host "✅ This is completely safe and standard for software installation!" -ForegroundColor Green
        Write-Host ""
        
        Write-Host "Press any key to request administrator privileges..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        try {
            Write-Host "🔄 Requesting administrator privileges..." -ForegroundColor Blue
            
            # Get the current script path
            $scriptPath = $MyInvocation.MyCommand.Path
            if (-not $scriptPath) {
                $scriptPath = $PSCommandPath
            }
            
            if ($scriptPath) {
                # Method 1: Try with the script file path
                Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
            } else {
                # Method 2: Try with encoded command
                $bytes = [System.Text.Encoding]::Unicode.GetBytes($MyInvocation.MyCommand.Definition)
                $encodedCommand = [Convert]::ToBase64String($bytes)
                Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCommand" -Verb RunAs
            }
            
            Write-Host "✅ New window should open with admin rights. You can close this one." -ForegroundColor Green
            Write-Host ""
            Write-Host "⏳ Waiting 5 seconds before closing this window..." -ForegroundColor Gray
            Start-Sleep -Seconds 5
            
        } catch {
            Write-Host ""
            Write-Host "❌ Failed to request admin privileges!" -ForegroundColor Red
            Write-Host "   Error: $_" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "🔧 Manual solutions (try these in order):" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "   METHOD 1 - Right-Click as Admin:" -ForegroundColor Yellow
            Write-Host "   1. Right-click on this script file" -ForegroundColor White
            Write-Host "   2. Select 'Run as administrator'" -ForegroundColor White
            Write-Host "   3. Click 'Yes' when Windows asks for permission" -ForegroundColor White
            Write-Host ""
            Write-Host "   METHOD 2 - Copy & Paste Method:" -ForegroundColor Yellow
            Write-Host "   1. Press Win+X and select 'Windows PowerShell (Admin)'" -ForegroundColor White
            Write-Host "   2. Copy this entire script content" -ForegroundColor White
            Write-Host "   3. Paste it into the admin PowerShell window" -ForegroundColor White
            Write-Host "   4. Press Enter to run" -ForegroundColor White
            Write-Host ""
            Write-Host "   METHOD 3 - Command Line:" -ForegroundColor Yellow
            Write-Host "   1. Open Command Prompt as Administrator" -ForegroundColor White
            Write-Host "   2. Type: powershell -ExecutionPolicy Bypass -File `"[path to this script]`"" -ForegroundColor White
            Write-Host ""
        }
        Exit 1
    } else {
        Write-Host "✅ Running with administrator privileges - Ready to install!" -ForegroundColor Green
        Write-Host ""
        Write-Host "Press any key to continue..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# Setup Paths
$installerFolder = if ($PSScriptRoot) { "$PSScriptRoot\OfficeInstaller" } else { "$env:TEMP\OfficeInstaller" }

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
    $version = "v3.2 - Universal Pause Protection"
    
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
        [string]$Color = "White",
        [string]$Recommendation = ""
    )
    
    Write-Host "  " -NoNewline
    Write-Host "[$Number]" -ForegroundColor Cyan -NoNewline
    Write-Host " $Title" -ForegroundColor $Color
    if ($Description) {
        Write-Host "      $Description" -ForegroundColor DarkGray
    }
    if ($Recommendation) {
        Write-Host "      " -NoNewline
        Write-Host $Recommendation -ForegroundColor Green
    }
}

function Get-UserChoice {
    param(
        [string]$Prompt,
        [string[]]$ValidChoices,
        [string]$DefaultChoice = $ValidChoices[0],
        [string]$HelpText = ""
    )
    
    do {
        Write-Host ""
        if ($HelpText) {
            Write-Host "💡 $HelpText" -ForegroundColor Blue
        }
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
            Write-Host "   Just type the number and press Enter!" -ForegroundColor Gray
        }
    } while ($true)
}

function Test-SystemRequirements {
    Show-Header
    Write-Host "🔍 CHECKING SYSTEM REQUIREMENTS" -ForegroundColor Yellow
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    
    Log "Starting system requirements check..."
    
    # Check Windows version
    Show-Progress -Activity "System Check" -PercentComplete 20 -Status "Checking Windows version..."
    $osVersion = [System.Environment]::OSVersion.Version
    Write-Host "🖥️  Windows Version: " -ForegroundColor Cyan -NoNewline
    Write-Host "$($osVersion.Major).$($osVersion.Minor)" -ForegroundColor White
    
    if ($osVersion.Major -lt 10) {
        Write-Host "⚠️  Warning: Windows 10 or later is recommended for best compatibility" -ForegroundColor Yellow
    } else {
        Write-Host "✅ Windows version is compatible" -ForegroundColor Green
    }
    
    # Check available disk space
    Show-Progress -Activity "System Check" -PercentComplete 40 -Status "Checking disk space..."
    $systemDrive = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq $env:SystemDrive }
    $freeSpaceGB = [math]::Round($systemDrive.FreeSpace / 1GB, 2)
    Write-Host "💾 Available Space: " -ForegroundColor Cyan -NoNewline
    Write-Host "$freeSpaceGB GB" -ForegroundColor White
    
    if ($freeSpaceGB -lt 4) {
        Write-Host "❌ Error: At least 4GB of free space is required" -ForegroundColor Red
        Write-Host "   Please free up some disk space and try again" -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        Exit 1
    } else {
        Write-Host "✅ Sufficient disk space available" -ForegroundColor Green
    }
    
    # Check RAM
    Show-Progress -Activity "System Check" -PercentComplete 60 -Status "Checking system memory..."
    $totalRAM = [math]::Round((Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
    Write-Host "🧠 System RAM: " -ForegroundColor Cyan -NoNewline
    Write-Host "$totalRAM GB" -ForegroundColor White
    
    if ($totalRAM -lt 2) {
        Write-Host "⚠️  Warning: 4GB RAM or more is recommended for optimal performance" -ForegroundColor Yellow
    } else {
        Write-Host "✅ RAM meets requirements" -ForegroundColor Green
    }
    
    # Check internet connection
    Show-Progress -Activity "System Check" -PercentComplete 80 -Status "Testing internet connectivity..."
    Write-Host "🌐 Internet Connection: " -ForegroundColor Cyan -NoNewline
    
    Try {
        $null = Invoke-WebRequest -Uri "https://www.microsoft.com" -UseBasicParsing -TimeoutSec 10
        Write-Host "Connected" -ForegroundColor Green
        Log "Internet connection verified."
    } Catch {
        Write-Host "Failed" -ForegroundColor Red
        Write-Host "❌ No internet connection detected!" -ForegroundColor Red
        Write-Host "   Please check your internet connection and try again" -ForegroundColor Yellow
        Log "Internet connection failed. Exiting script."
        Read-Host "Press Enter to exit"
        Exit 1
    }
    
    Show-Progress -Activity "System Check" -PercentComplete 100 -Status "System check complete"
    Write-Host ""
    Write-Host "✅ All system requirements met!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Press any key to continue..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Show-BeginnerFriendlyMenu {
    Show-Header
    Write-Host "📋 OFFICE SETUP - MADE SIMPLE!" -ForegroundColor Yellow
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "🎯 Don't worry - we'll guide you through each step!" -ForegroundColor Green
    Write-Host "   Just pick the options that sound right for you." -ForegroundColor Gray
    Write-Host ""
    
    Log "Starting user-friendly configuration setup."

    # Simple architecture selection
    Write-Host "🏗️  STEP 1: CHOOSE YOUR SYSTEM TYPE" -ForegroundColor Magenta
    Write-Host "   (Don't know? Choose option 1 - it works for most computers)" -ForegroundColor Gray
    Write-Host ""
    Show-MenuOption "1" "64-bit (Recommended)" "For most modern computers (2010 and newer)" "White" "✨ Best choice for most users"
    Show-MenuOption "2" "32-bit" "For older computers or specific compatibility needs"
    $arch = Get-UserChoice "What type of computer do you have?" @("1", "2") "1" "If unsure, choose 1 - it works on most computers!"
    $bit = if ($arch -eq "2") { "32" } else { "64" }

    Write-Host ""
    Write-Host "📦 STEP 2: CHOOSE YOUR OFFICE VERSION" -ForegroundColor Magenta
    Write-Host "   (Each version has the same core apps: Word, Excel, PowerPoint, Outlook)" -ForegroundColor Gray
    Write-Host ""
    Show-MenuOption "1" "Office 2024 Pro Plus" "Latest version with newest features" "White" "✨ Most popular choice"
    Show-MenuOption "2" "Office LTSC 2021" "Stable version, less frequent updates"
    Show-MenuOption "3" "Microsoft 365 Apps" "Cloud-connected with online features"
    $editionChoice = Get-UserChoice "Which Office version would you like?" @("1", "2", "3") "1" "Option 1 gives you the latest features and is most commonly used"
    
    $editionMap = @{ 
        "1" = @{ID = "ProPlus2024Retail"; Name = "Office 2024 Pro Plus"}
        "2" = @{ID = "ProPlus2021Volume"; Name = "Office LTSC 2021"}
        "3" = @{ID = "O365ProPlusRetail"; Name = "Microsoft 365 Apps"}
    }
    $edition = $editionMap[$editionChoice]

    Write-Host ""
    Write-Host "🎨 STEP 3: EXTRA PROGRAMS (OPTIONAL)" -ForegroundColor Magenta
    Write-Host "   (These are bonus programs - you can skip them if you don't need them)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "📊 Visio (for creating diagrams and flowcharts):" -ForegroundColor Cyan
    Show-MenuOption "1" "Yes, include Visio" "Adds diagram and flowchart creation tools"
    Show-MenuOption "2" "No, skip Visio" "Just install the main Office programs" "White" "✨ Most users choose this"
    $visio = Get-UserChoice "Do you want Visio?" @("1", "2") "2" "Most people don't need Visio - it's for making diagrams"

    Write-Host ""
    Write-Host "📋 Project (for project management):" -ForegroundColor Cyan
    Show-MenuOption "1" "Yes, include Project" "Adds project management tools"
    Show-MenuOption "2" "No, skip Project" "Just install the main Office programs" "White" "✨ Most users choose this"
    $project = Get-UserChoice "Do you want Project?" @("1", "2") "2" "Most people don't need Project - it's for managing big projects"

    Write-Host ""
    Write-Host "🔄 STEP 4: HOW OFTEN TO UPDATE" -ForegroundColor Magenta
    Write-Host "   (This controls how often Office gets new features)" -ForegroundColor Gray
    Write-Host ""
    Show-MenuOption "1" "Monthly updates" "Get new features as soon as they're ready" "White" "✨ Recommended for most users"
    Show-MenuOption "2" "Less frequent updates" "Get updates after they've been tested more"
    $channelChoice = Get-UserChoice "How often do you want updates?" @("1", "2") "1" "Monthly updates give you the latest features and security fixes"
    $channel = if ($channelChoice -eq "2") { "Broad" } else { "Current" }

    Write-Host ""
    Write-Host "🌍 STEP 5: CHOOSE YOUR LANGUAGE" -ForegroundColor Magenta
    Write-Host ""
    Show-MenuOption "1" "English (United States)" "en-us" "White" "✨ Most common choice"
    Show-MenuOption "2" "English (United Kingdom)" "en-gb"
    Show-MenuOption "3" "French (France)" "fr-fr"
    Show-MenuOption "4" "German (Germany)" "de-de"
    Show-MenuOption "5" "Dutch (Netherlands)" "nl-nl"
    Show-MenuOption "6" "Spanish (Spain)" "es-es"
    Show-MenuOption "7" "Portuguese (Brazil)" "pt-br"
    $langChoice = Get-UserChoice "What language do you want?" @("1", "2", "3", "4", "5", "6", "7") "1" "Choose the language you're most comfortable with"

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
    Write-Host "🖥️  STEP 6: INSTALLATION STYLE" -ForegroundColor Magenta
    Write-Host "   (This is just about what you see during installation)" -ForegroundColor Gray
    Write-Host ""
    Show-MenuOption "1" "Show me the installation progress" "You'll see what's happening during install" "White" "✨ Recommended - lets you see progress"
    Show-MenuOption "2" "Install quietly in background" "Install without showing progress windows"
    $uiChoice = Get-UserChoice "How do you want to install?" @("1", "2") "1" "Option 1 lets you see what's happening - it's more reassuring!"
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

function Show-FriendlyConfigurationSummary($options) {
    Show-Header
    Write-Host "📋 READY TO INSTALL!" -ForegroundColor Yellow
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "🎉 Great! Here's what we're going to install for you:" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "📦 Office Version:  " -ForegroundColor Cyan -NoNewline
    Write-Host $options.editionName -ForegroundColor White
    
    Write-Host "🏗️  System Type:     " -ForegroundColor Cyan -NoNewline
    Write-Host "$($options.bit)-bit" -ForegroundColor White
    
    Write-Host "🌍 Language:        " -ForegroundColor Cyan -NoNewline
    Write-Host $options.languageName -ForegroundColor White
    
    Write-Host "🔄 Updates:         " -ForegroundColor Cyan -NoNewline
    Write-Host $(if ($options.channel -eq "Current") { "Monthly (recommended)" } else { "Less frequent" }) -ForegroundColor White
    
    Write-Host "📊 Visio:           " -ForegroundColor Cyan -NoNewline
    Write-Host $(if ($options.visio -eq "1") { "Yes, included" } else { "No, not included" }) -ForegroundColor White
    
    Write-Host "📋 Project:         " -ForegroundColor Cyan -NoNewline
    Write-Host $(if ($options.project -eq "1") { "Yes, included" } else { "No, not included" }) -ForegroundColor White
    
    Write-Host "🖥️  Installation:    " -ForegroundColor Cyan -NoNewline
    Write-Host $(if ($options.ui -eq "Full") { "Show progress" } else { "Quiet background" }) -ForegroundColor White
    
    Write-Host ""
    Write-Host "⏱️  Installation will take about 10-30 minutes depending on your internet speed." -ForegroundColor Blue
    Write-Host "☕ Perfect time to grab a coffee!" -ForegroundColor Gray
    Write-Host ""
    Write-Host "✅ Everything look good? Press any key to start installing..." -ForegroundColor Green
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Download-ODT {
    $url = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    $output = "$installerFolder\setup.exe"

    Show-Header
    Write-Host "📥 DOWNLOADING OFFICE INSTALLER" -ForegroundColor Yellow
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "🔗 Getting the official Microsoft Office installer..." -ForegroundColor Blue
    Write-Host "   This is completely safe - we're downloading directly from Microsoft!" -ForegroundColor Gray
    Write-Host ""
    
    Log "Downloading Office Deployment Tool from $url..."
    
    Show-Progress -Activity "Download" -PercentComplete 0 -Status "Connecting to Microsoft servers..."
    
    try {
        # Create a WebClient for progress tracking
        $webClient = New-Object System.Net.WebClient
        
        # Register progress event
        Register-ObjectEvent -InputObject $webClient -EventName DownloadProgressChanged -Action {
            $percent = $Event.SourceEventArgs.ProgressPercentage
            Show-Progress -Activity "Download" -PercentComplete $percent -Status "Downloading Office installer... $percent%"
        } | Out-Null
        
        # Download the file
        $webClient.DownloadFile($url, $output)
        $webClient.Dispose()
        
        Show-Progress -Activity "Download" -PercentComplete 100 -Status "Download complete!"
        
    } catch {
        Log "Download failed: $_"
        Write-Host "❌ Download failed!" -ForegroundColor Red
        Write-Host "   Error: $_" -ForegroundColor Yellow
        Write-Host "   Please check your internet connection and try again." -ForegroundColor Gray
        Read-Host "Press Enter to exit"
        Exit 1
    }

    if (-Not (Test-Path $output) -or ((Get-Item $output).Length -lt 100000)) {
        Log "Downloaded file appears to be corrupted or incomplete."
        Write-Host "❌ Download seems incomplete. Please try again." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        Exit 1
    }

    Log "Office Deployment Tool downloaded successfully."
    Write-Host "✅ Download completed successfully!" -ForegroundColor Green
    Start-Sleep -Seconds 2
}

function Generate-Config($options) {
    Show-Header
    Write-Host "⚙️  PREPARING INSTALLATION" -ForegroundColor Yellow
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "📝 Creating your personalized Office configuration..." -ForegroundColor Blue
    Write-Host ""
    
    Log "Generating config.xml with selected options..."
    Show-Progress -Activity "Configuration" -PercentComplete 50 -Status "Creating installation configuration..."

    $products = @()
    $products += "<Product ID='" + $options.edition + "'>`n  <Language ID='" + $options.language + "' />`n</Product>"

    if ($options.visio -eq "1") {
        $products += "<Product ID='VisioPro2021Volume'>`n  <Language ID='" + $options.language + "' />`n</Product>"
        Write-Host "📊 Adding Visio Professional to your installation..." -ForegroundColor Green
    }
    if ($options.project -eq "1") {
        $products += "<Product ID='ProjectPro2021Volume'>`n  <Language ID='" + $options.language + "' />`n</Product>"
        Write-Host "📋 Adding Project Professional to your installation..." -ForegroundColor Green
    }

    $xmlContent = @"
<Configuration>
  <Add OfficeClientEdition="${($options.bit)}" Channel="${($options.channel)}">
    $($products -join "`n    ")
  </Add>
  <Display Level="${($options.ui)}" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
"@

    $configPath = "$installerFolder\config.xml"
    $xmlContent | Out-File -FilePath $configPath -Encoding UTF8
    Log "config.xml generated at $configPath"
    
    Show-Progress -Activity "Configuration" -PercentComplete 100 -Status "Configuration ready!"
    Write-Host "✅ Configuration created successfully!" -ForegroundColor Green
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
        Write-Host "❌ Installation file missing!" -ForegroundColor Red
        Write-Host "   Something went wrong with the download. Please restart the script." -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        Exit 1
    }

    Write-Host "🎯 Starting your Office installation now!" -ForegroundColor Green
    Write-Host ""
    Write-Host "⏳ This will take 10-30 minutes depending on:" -ForegroundColor Blue
    Write-Host "   • Your internet speed (Office downloads during installation)" -ForegroundColor Gray
    Write-Host "   • Your computer's performance" -ForegroundColor Gray
    Write-Host "   • Which components you selected" -ForegroundColor Gray
    Write-Host ""
    Write-Host "☕ Perfect time for a coffee break!" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "⚠️  IMPORTANT: Don't close this window or turn off your computer!" -ForegroundColor Red
    Write-Host "   Doing so could corrupt the installation." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "🔄 Starting installation in 3 seconds..." -ForegroundColor Green
    Start-Sleep -Seconds 3

    Show-Progress -Activity "Installation" -PercentComplete 0 -Status "Launching Office installer..."
    
    try {
        Log "Executing: $setupExe /configure config.xml"
        $process = Start-Process -FilePath $setupExe -ArgumentList "/configure config.xml" -PassThru -NoNewWindow
        
        # Monitor the installation process
        $counter = 0
        while (-not $process.HasExited) {
            $counter++
            $percent = [Math]::Min(90, $counter * 2)  # Cap at 90% until we know it's done
            Show-Progress -Activity "Installation" -PercentComplete $percent -Status "Installing Office... Please wait..."
            Start-Sleep -Seconds 5
        }
        
        $process.WaitForExit()
        $exitCode = $process.ExitCode
        
        if ($exitCode -eq 0) {
            Show-Progress -Activity "Installation" -PercentComplete 100 -Status "Installation completed successfully!"
            Log "Office installation completed successfully with exit code: $exitCode"
            Write-Host "✅ Office installation completed successfully!" -ForegroundColor Green
        } else {
            Log "Office installation failed with exit code: $exitCode"
            Write-Host "⚠️  Installation completed with warnings (Exit code: $exitCode)" -ForegroundColor Yellow
            Write-Host "   Office should still work normally. Check the programs in your Start Menu." -ForegroundColor Gray
        }
        
    } catch {
        Log "Installation failed: $_"
        Write-Host "❌ Installation encountered an error!" -ForegroundColor Red
        Write-Host "   Error: $_" -ForegroundColor Yellow
        Write-Host "   You can try running the script again, or contact support." -ForegroundColor Gray
        Read-Host "Press Enter to exit"
        Exit 1
    }
}

function Show-FriendlyCompletionSummary($options) {
    Show-Header
    Write-Host "🎉 CONGRATULATIONS! OFFICE IS INSTALLED!" -ForegroundColor Green
    Write-Host "═" * 50 -ForegroundColor DarkGray
    Write-Host ""
    
    Write-Host "✅ Your Microsoft Office installation is complete!" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "📦 What was installed:" -ForegroundColor Cyan
    Write-Host "   • " -NoNewline -ForegroundColor White
    Write-Host $options.editionName -ForegroundColor Yellow
    Write-Host "   • Word, Excel, PowerPoint, Outlook, and more!" -ForegroundColor White
    if ($options.visio -eq "1") {
        Write-Host "   • Visio Professional (for diagrams)" -ForegroundColor White
    }
    if ($options.project -eq "1") {
        Write-Host "   • Project Professional (for project management)" -ForegroundColor White
    }
    Write-Host "   • Language: " -NoNewline -ForegroundColor White
    Write-Host $options.languageName -ForegroundColor Yellow
    Write-Host "   • Architecture: " -NoNewline -ForegroundColor White
    Write-Host "$($options.bit)-bit" -ForegroundColor Yellow
    
    Write-Host ""
    Write-Host "🚀 How to start using Office:" -ForegroundColor Cyan
    Write-Host "   1. Click the Windows Start button" -ForegroundColor White
    Write-Host "   2. Look for 'Word', 'Excel', 'PowerPoint', etc." -ForegroundColor White
    Write-Host "   3. Click on any Office app to start using it!" -ForegroundColor White
    
    Write-Host ""
    Write-Host "🔐 First time setup:" -ForegroundColor Cyan
    Write-Host "   • Office may ask you to sign in with a Microsoft account" -ForegroundColor White
    Write-Host "   • This is normal and helps sync your settings" -ForegroundColor White
    Write-Host "   • You can skip this if you prefer to use Office offline" -ForegroundColor White
    
    Write-Host ""
    Write-Host "🔄 Keeping Office updated:" -ForegroundColor Cyan
    Write-Host "   • Office will automatically check for updates" -ForegroundColor White
    Write-Host "   • You chose: " -NoNewline -ForegroundColor White
    Write-Host $(if ($options.channel -eq "Current") { "Monthly updates (recommended)" } else { "Less frequent updates" }) -ForegroundColor Yellow
    
    Write-Host ""
    Write-Host "📁 Installation files:" -ForegroundColor Cyan
    Write-Host "   • Saved in: " -NoNewline -ForegroundColor White
    Write-Host $installerFolder -ForegroundColor Yellow
    Write-Host "   • You can safely delete this folder after confirming Office works" -ForegroundColor White
    Write-Host "   • Installation log: " -NoNewline -ForegroundColor White
    Write-Host $logFile -ForegroundColor Yellow
    
    Write-Host ""
    Write-Host "❓ Need help?" -ForegroundColor Cyan
    Write-Host "   • Office has built-in help - just press F1 in any Office app" -ForegroundColor White
    Write-Host "   • Visit support.microsoft.com for online help" -ForegroundColor White
    Write-Host "   • Check Windows Updates for the latest Office patches" -ForegroundColor White
    
    Write-Host ""
    Write-Host "⚖️  Important reminder:" -ForegroundColor Yellow
    Write-Host "   This installer uses Microsoft's official tools and doesn't modify licensing." -ForegroundColor Gray
    Write-Host "   Make sure you have proper licensing for your Office installation." -ForegroundColor Gray
    
    Write-Host ""
    Write-Host "🎊 Enjoy your new Microsoft Office installation!" -ForegroundColor Green
    Write-Host ""
    
    # Final pause message
    Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor DarkGray
    Write-Host "Installation complete! Press any key to close this window..." -ForegroundColor Yellow -BackgroundColor DarkBlue
    Write-Host "═══════════════════════════════════════════════════════════════════════════════" -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ==== Main Execution Flow ====

try {
    Log "=== Enhanced Office Installer Started (Universal Pause Protection Version) ==="
    
    # Show welcome screen and handle admin elevation
    Show-WelcomeScreen
    
    # Run system requirements check
    Test-SystemRequirements
    
    # Get user configuration with beginner-friendly interface
    $options = Show-BeginnerFriendlyMenu
    
    # Show friendly summary
    Show-FriendlyConfigurationSummary -options $options
    
    # Download, configure, and install
    Download-ODT
    Generate-Config -options $options
    Install-Office
    
    # Show completion summary
    Show-FriendlyCompletionSummary -options $options
    
    Log "=== Enhanced Office Installer Completed Successfully ==="
    
} catch {
    Log "FATAL ERROR: $_"
    Write-Host ""
    Write-Host "❌ SOMETHING WENT WRONG!" -ForegroundColor Red
    Write-Host "═" * 30 -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Don't worry - this happens sometimes. Here's what you can try:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "🔧 Quick fixes:" -ForegroundColor Cyan
    Write-Host "   1. Make sure you're connected to the internet" -ForegroundColor White
    Write-Host "   2. Try running the script again" -ForegroundColor White
    Write-Host "   3. Restart your computer and try again" -ForegroundColor White
    Write-Host "   4. Temporarily disable antivirus and try again" -ForegroundColor White
    Write-Host ""
    Write-Host "📋 Error details:" -ForegroundColor Cyan
    Write-Host "   $($_)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "📁 Check the log file for more details:" -ForegroundColor Cyan
    Write-Host "   $logFile" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "🔧 Copy & Paste Method:" -ForegroundColor Yellow
    Write-Host "   If this keeps failing, try copying this entire script" -ForegroundColor White
    Write-Host "   and pasting it into an Administrator PowerShell window" -ForegroundColor White
    Write-Host ""
    Exit 1
}

# Disable the pause flag at the very end for clean exit
$global:ShouldPauseOnExit = $false