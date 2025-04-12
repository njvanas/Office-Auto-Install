Pause

# Setup Paths
$installerFolder = "$PSScriptRoot\OfficeInstaller"

# Auto-elevate if not running as admin
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Host "⚠️  Script is not running as Administrator. Attempting to relaunch..."
    try {
        Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    } catch {
        Write-Host "❌ Failed to relaunch script with admin rights. Error: $_"
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

# Now that we’re elevated, we can log
$logFile = "$installerFolder\installer.log"

function Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -Append -FilePath $logFile
}

function Fix-SystemPath {
    Log "Checking and fixing essential system PATH entries..."

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
        Write-Host "✅ System PATH updated with $added missing entries. You may need to log out or restart."
    } else {
        Log "No changes made to PATH."
    }
}

function Check-Internet {
    Write-Host "Checking internet connection..."
    Try {
        $null = Invoke-WebRequest -Uri "https://www.google.com" -UseBasicParsing -TimeoutSec 5
        Log "Internet connection verified."
    } Catch {
        Log "Internet connection failed. Exiting script."
        Write-Host "No internet connection. Please connect and try again."
        Exit 1
    }
}

function Show-Menu {
    Clear-Host
    Log "Prompting user for selections."

    Write-Host "Select Office Architecture:"
    Write-Host "1: 64-bit"
    Write-Host "2: 32-bit"
    $arch = Read-Host "Enter choice (1 or 2)"
    $bit = if ($arch -eq "2") { "32" } else { "64" }

    Write-Host "`nInclude Visio?"
    Write-Host "1: Yes"
    Write-Host "2: No"
    $visio = Read-Host "Enter choice (1 or 2)"

    Write-Host "`nInclude Project?"
    Write-Host "1: Yes"
    Write-Host "2: No"
    $project = Read-Host "Enter choice (1 or 2)"

    Write-Host "`nSelect Update Channel:"
    Write-Host "1: Monthly"
    Write-Host "2: Semi-Annual"
    $channelChoice = Read-Host "Enter choice (1 or 2)"
    $channel = if ($channelChoice -eq "2") { "Broad" } else { "Current" }

    Write-Host "`nSelect Language:"
    Write-Host "1: English (United States) [en-us]"
    Write-Host "2: English (United Kingdom) [en-gb]"
    Write-Host "3: French (France) [fr-fr]"
    Write-Host "4: German (Germany) [de-de]"
    Write-Host "5: Dutch (Netherlands) [nl-nl]"
    Write-Host "6: Spanish (Spain) [es-es]"
    Write-Host "7: Portuguese (Brazil) [pt-br]"
    $langChoice = Read-Host "Enter choice (1-7)"

    $languageMap = @{
        "1" = "en-us"
        "2" = "en-gb"
        "3" = "fr-fr"
        "4" = "de-de"
        "5" = "nl-nl"
        "6" = "es-es"
        "7" = "pt-br"
    }

    $language = $languageMap[$langChoice]
    if (-not $language) {
        Write-Host "Invalid selection. Defaulting to en-us"
        $language = "en-us"
    }

    Write-Host "`nDisplay installation UI?"
    Write-Host "1: Yes (Show Office setup window)"
    Write-Host "2: No (Silent install)"
    $uiChoice = Read-Host "Enter choice (1 or 2)"
    $uiLevel = if ($uiChoice -eq "1") { "Full" } else { "None" }

    return @{ bit=$bit; visio=$visio; project=$project; channel=$channel; language=$language; ui=$uiLevel }
}

function Download-ODT {
    $url = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    $output = "$installerFolder\setup.exe"

    Log "Downloading Office Deployment Tool from $url..."
    Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing

    if (-Not (Test-Path $output) -or ((Get-Item $output).Length -lt 100000)) {
        Log "Downloaded file appears to be corrupted or incomplete."
        Write-Host "`n❌ The downloaded file is too small or unreadable. Please check your connection and try again."
        Exit 1
    }

    Log "Office Deployment Tool downloaded successfully."
}

function Generate-Config($options) {
    Log "Generating config.xml with selected options..."

    $products = @()
    $products += "<Product ID='ProPlus2024Retail'>`n  <Language ID='" + $options.language + "' />`n</Product>"

    if ($options.visio -eq "1") {
        $products += "<Product ID='VisioPro2024Retail'>`n  <Language ID='" + $options.language + "' />`n</Product>"
    }
    if ($options.project -eq "1") {
        $products += "<Product ID='ProjectPro2024Retail'>`n  <Language ID='" + $options.language + "' />`n</Product>"
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
}

function Install-Office {
    Log "Starting Office installation..."
    $setupExe = "$installerFolder\setup.exe"

    if (-Not (Test-Path $setupExe)) {
        Log "ERROR: setup.exe not found."
        Write-Host "`n❌ Error: setup.exe not found in $installerFolder."
        Exit 1
    }

    Start-Process -FilePath $setupExe -ArgumentList "/configure config.xml" -Wait
    Log "Office installation process finished."
}

# ==== Main Execution Flow ====

Log "=== Script Started ==="
Fix-SystemPath
Check-Internet
$options = Show-Menu
Download-ODT
Generate-Config -options $options
Install-Office
Log "=== Script Completed ==="

Write-Host "`nInstallation Complete!"
Write-Host "Office has been installed using the following configuration:"
Write-Host " - Architecture: $($options.bit)-bit"
Write-Host " - Language: $($options.language)"
Write-Host " - Channel: $($options.channel)"
Write-Host " - Visio: $([string]::Join('', @('No', 'Yes')[$options.visio -eq '1']))"
Write-Host " - Project: $([string]::Join('', @('No', 'Yes')[$options.project -eq '1']))"
Write-Host " - UI Level: $($options.ui)"
Write-Host " - Installer folder: $installerFolder"
Write-Host "`nInstallation log saved at: $logFile"
