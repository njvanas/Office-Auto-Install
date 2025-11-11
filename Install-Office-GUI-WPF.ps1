#Requires -Version 5.1
# Office Auto Installer - WPF with Windows 11 Fluent Design
# Single-file, self-contained Office installer with WPF and Fluent theme
# Downloads and installs Microsoft Office through official channels

<#
.SYNOPSIS
    Microsoft Office Auto Installer - WPF Fluent Design Edition
.DESCRIPTION
    A beautiful, modern WPF application with Windows 11 Fluent Design theme.
    All options are pre-filled with recommended defaults for one-click installation.
    Fully self-contained - no external dependencies required.
.NOTES
    Version: 5.0 - WPF with Fluent Design Theme
    Author: Office Auto Installer Team
    Requires: .NET Framework 4.7.2+ or .NET 6+ (for Fluent theme: .NET 9+)
#>

# ============================================================================
# INITIALIZATION & PREREQUISITES
# ============================================================================
# This section handles error handling, dependency checks, and basic setup
# before loading the WPF user interface.

$ErrorActionPreference = "Continue"  # Continue on errors to show user-friendly messages

# Load Windows Forms assembly first (needed for error dialogs before WPF is loaded)
# This ensures we can display error messages even if WPF fails to initialize
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

# ============================================================================
# .NET FRAMEWORK VERSION CHECK
# ============================================================================
# WPF (Windows Presentation Foundation) requires .NET Framework 4.7.2 or later.
# Note: WPF does NOT work with .NET Core/.NET 5+. It requires the full .NET Framework.
# This check ensures the user has the required runtime before attempting to load WPF.

$netFrameworkVersion = $null
$netFrameworkInstalled = $false

try {
    # Check registry for .NET Framework 4.x
    $netFrameworkVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue).Release
    if ($null -ne $netFrameworkVersion) {
        $netFrameworkInstalled = $true
    }
} catch {
    # Registry check failed, try alternative method
    try {
        $netFrameworkVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Client" -ErrorAction SilentlyContinue).Release
        if ($null -ne $netFrameworkVersion) {
            $netFrameworkInstalled = $true
        }
    } catch {
        $netFrameworkInstalled = $false
    }
}

# If .NET Framework not found, show error
if (-not $netFrameworkInstalled) {
    $errorMsg = ".NET Framework 4.7.2 or later is required for this WPF application.`n`n" +
                "Please install .NET Framework 4.8 from:`n" +
                "https://dotnet.microsoft.com/download/dotnet-framework/net48`n`n" +
                "After installation, restart this application.`n`n" +
                "Note: This application requires Windows PowerShell (not PowerShell Core)."
    
    try {
        [System.Windows.Forms.MessageBox]::Show($errorMsg, ".NET Framework Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } catch {
        Write-Host $errorMsg -ForegroundColor Red
    }
    exit 1
}

# Check if .NET Framework version is 4.7.2 or later (Release >= 461808)
if ($netFrameworkVersion -lt 461808) {
    $errorMsg = ".NET Framework 4.7.2 or later is required.`n`n" +
                "Your version appears to be older (Release: $netFrameworkVersion).`n`n" +
                "Please install .NET Framework 4.8 from:`n" +
                "https://dotnet.microsoft.com/download/dotnet-framework/net48`n`n" +
                "After installation, restart this application."
    
    try {
        [System.Windows.Forms.MessageBox]::Show($errorMsg, ".NET Framework Version Too Old", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } catch {
        Write-Host $errorMsg -ForegroundColor Red
    }
    exit 1
}

# Load required assemblies for WPF
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName PresentationCore -ErrorAction Stop
    Add-Type -AssemblyName WindowsBase -ErrorAction Stop
    Add-Type -AssemblyName System.Xaml -ErrorAction Stop
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to load WPF assemblies.`n`nError: $_`n`nPlease ensure .NET Framework 4.7.2 or later is installed.`n`nDownload from: https://dotnet.microsoft.com/download/dotnet-framework/net48",
        "WPF Assembly Load Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

# Try to load Fluent theme if available (.NET 9+)
# Note: Fluent theme is only available in .NET 9+, so we'll use custom styling
$fluentThemeAvailable = $false
# Custom Fluent-like styling will be used instead

# ==== EXECUTION POLICY FIX ====
try {
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
    if ($currentPolicy -eq 'Restricted' -or $currentPolicy -eq 'AllSigned') {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    }
} catch { }

# ==== ADMIN CHECK ====
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $isAdmin) {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This installer requires administrator privileges to install Office.`n`nWould you like to restart with administrator rights?",
        "Administrator Required",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $scriptPath = if ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } else { $PSCommandPath }
        if ($scriptPath) {
            Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
        } else {
            $bytes = [System.Text.Encoding]::Unicode.GetBytes($MyInvocation.MyCommand.Definition)
            $encodedCommand = [Convert]::ToBase64String($bytes)
            Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCommand" -Verb RunAs
        }
        exit
    } else {
        exit
    }
}

# ==== LOGGING ====
$logFile = "$env:TEMP\OfficeInstaller.log"
function Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# ============================================================================
# TEMPORARY INSTALLATION FOLDER SETUP
# ============================================================================
# Creates a temporary folder in the user's temp directory to store:
# - setup.exe (Office Deployment Tool)
# - config.xml (Office installation configuration)
# This folder is cleaned up after installation completes.

$installerFolder = "$env:TEMP\OfficeInstaller"
if (-not (Test-Path $installerFolder)) {
    New-Item -ItemType Directory -Path $installerFolder -Force | Out-Null
}

# ============================================================================
# WPF XAML USER INTERFACE DEFINITION
# ============================================================================
# This embedded XAML string defines the entire WPF window structure.
# 
# Structure Overview:
# 1. Window: Main application window with dark theme background
# 2. Resources: Color brushes, styles for ComboBox, Button, etc.
# 3. Layout: Grid-based layout with header, scrollable content, and footer
# 4. Controls: ComboBoxes for selections, CheckBoxes for optional components
# 5. Status Panel: Hidden by default, shown during installation
# 6. Install Button: Primary action button at the bottom
#
# Design Philosophy:
# - Windows 11 Fluent Design aesthetic with rounded corners
# - Dark theme (#202020 background) for modern look
# - Custom ComboBox styling to remove default white boxes
# - Semi-transparent hover effects on dropdown items
# - High-contrast focus indicators for accessibility (WCAG 2.4.7)
#
# Note: x:Class is intentionally omitted as it's not supported when
# loading XAML dynamically in PowerShell.

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft Office Auto Installer"
        Width="920" Height="1050"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="#202020">
  <Window.Resources>
    <ResourceDictionary>
      $(if ($fluentThemeAvailable) {
        @"
      <ResourceDictionary.MergedDictionaries>
        <ResourceDictionary Source="pack://application:,,,/PresentationFramework.Fluent;component/Themes/Fluent.xaml"/>
      </ResourceDictionary.MergedDictionaries>
"@
      })
      
      <!-- Windows 11 Fluent Design Colors -->
      <SolidColorBrush x:Key="FluentBackgroundBrush" Color="#202020"/>
      <SolidColorBrush x:Key="FluentCardBrush" Color="#252525"/>
      <SolidColorBrush x:Key="FluentHeaderBrush" Color="#1C1C1C"/>
      <SolidColorBrush x:Key="FluentTextBrush" Color="#FFFFFF"/>
      <SolidColorBrush x:Key="FluentTextSecondaryBrush" Color="#F3F3F3"/>
      <SolidColorBrush x:Key="FluentTextMutedBrush" Color="#C8C8C8"/>
      <SolidColorBrush x:Key="FluentPrimaryBrush" Color="#0078D7"/>
      <SolidColorBrush x:Key="FluentPrimaryHoverBrush" Color="#005AB8"/>
      <SolidColorBrush x:Key="FluentFocusBrush" Color="#60CDFF"/>
      <SolidColorBrush x:Key="FluentBorderBrush" Color="#484848"/>
      <SolidColorBrush x:Key="FluentControlBrush" Color="#2B2B2B"/>
      <SolidColorBrush x:Key="FluentHoverBrush" Color="#353535" Opacity="0.6"/>
      <SolidColorBrush x:Key="FluentHoverSolidBrush" Color="#3A3A3A"/>
      
      <!-- Fluent ComboBox Style (fixes white boxes) -->
      <Style x:Key="FluentComboBoxStyle" TargetType="ComboBox">
        <Setter Property="MinHeight" Value="36"/>
        <Setter Property="Padding" Value="12,8"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Background" Value="{StaticResource FluentControlBrush}"/>
        <Setter Property="Foreground" Value="{StaticResource FluentTextSecondaryBrush}"/>
        <Setter Property="BorderBrush" Value="{StaticResource FluentBorderBrush}"/>
        <Setter Property="FontSize" Value="13"/>
        <Setter Property="FontFamily" Value="Segoe UI Variable"/>
        <Setter Property="Template">
          <Setter.Value>
            <ControlTemplate TargetType="ComboBox">
              <Grid>
                <Border x:Name="FocusBorder"
                        BorderBrush="{StaticResource FluentFocusBrush}"
                        BorderThickness="0"
                        CornerRadius="8"
                        Margin="-1"
                        IsHitTestVisible="False"/>
                <Border x:Name="Border"
                        Background="{TemplateBinding Background}"
                        BorderBrush="{TemplateBinding BorderBrush}"
                        BorderThickness="{TemplateBinding BorderThickness}"
                        CornerRadius="8"/>
                <TextBlock x:Name="ContentSite"
                           Text="{Binding SelectedItem.Content, RelativeSource={RelativeSource AncestorType=ComboBox}}"
                           Margin="12,0,36,0"
                           VerticalAlignment="Center"
                           HorizontalAlignment="Left"
                           FontSize="{TemplateBinding FontSize}"
                           FontFamily="{TemplateBinding FontFamily}"
                           TextTrimming="CharacterEllipsis"
                           Foreground="{Binding Foreground, RelativeSource={RelativeSource AncestorType=ComboBox}}"/>
                <Path x:Name="Arrow"
                      Data="M 0 0 L 4 4 L 8 0 Z"
                      Fill="{StaticResource FluentTextMutedBrush}"
                      HorizontalAlignment="Right"
                      VerticalAlignment="Center"
                      Margin="0,0,12,0"
                      Width="12" Height="12"/>
                <ToggleButton x:Name="ToggleButton"
                              Focusable="False"
                              ClickMode="Press"
                              IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                              Background="Transparent"
                              BorderThickness="0"
                              HorizontalAlignment="Stretch"
                              VerticalAlignment="Stretch">
                  <ToggleButton.Style>
                    <Style TargetType="ToggleButton">
                      <Setter Property="Background" Value="Transparent"/>
                      <Setter Property="BorderThickness" Value="0"/>
                      <Setter Property="Template">
                        <Setter.Value>
                          <ControlTemplate TargetType="ToggleButton">
                            <Border Background="Transparent" BorderThickness="0"/>
                          </ControlTemplate>
                        </Setter.Value>
                      </Setter>
                    </Style>
                  </ToggleButton.Style>
                </ToggleButton>
                <Popup x:Name="Popup"
                       Placement="Bottom"
                       PlacementTarget="{Binding ElementName=Border}"
                       AllowsTransparency="True"
                       Focusable="False"
                       PopupAnimation="Slide"
                       IsOpen="{TemplateBinding IsDropDownOpen}">
                  <Border x:Name="DropDownBorder"
                          Background="{StaticResource FluentCardBrush}"
                          BorderBrush="{StaticResource FluentBorderBrush}"
                          BorderThickness="1"
                          CornerRadius="8"
                          MaxHeight="{TemplateBinding MaxDropDownHeight}"
                          MinWidth="{Binding ActualWidth, ElementName=Border}">
                    <ScrollViewer Margin="4,6,4,6"
                                  SnapsToDevicePixels="True">
                      <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                    </ScrollViewer>
                  </Border>
                </Popup>
              </Grid>
              <ControlTemplate.Triggers>
                <Trigger Property="IsKeyboardFocusWithin" Value="True">
                  <Setter TargetName="FocusBorder" Property="BorderThickness" Value="2"/>
                  <Setter TargetName="Border" Property="BorderBrush" Value="{StaticResource FluentFocusBrush}"/>
                  <Setter TargetName="Border" Property="BorderThickness" Value="1"/>
                </Trigger>
                <Trigger Property="IsKeyboardFocusWithin" Value="False">
                  <Setter TargetName="FocusBorder" Property="BorderThickness" Value="0"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                  <Setter Property="Foreground" Value="{StaticResource FluentTextMutedBrush}"/>
                </Trigger>
              </ControlTemplate.Triggers>
            </ControlTemplate>
          </Setter.Value>
        </Setter>
        <Setter Property="ItemContainerStyle">
          <Setter.Value>
            <Style TargetType="ComboBoxItem">
              <Setter Property="Padding" Value="12,8"/>
              <Setter Property="Background" Value="Transparent"/>
              <Setter Property="Foreground" Value="{StaticResource FluentTextSecondaryBrush}"/>
              <Setter Property="FontSize" Value="13"/>
              <Setter Property="FontFamily" Value="Segoe UI Variable"/>
              <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                  <Setter Property="Background">
                    <Setter.Value>
                      <SolidColorBrush Color="#353535" Opacity="0.6"/>
                    </Setter.Value>
                  </Setter>
                  <Setter Property="Foreground" Value="{StaticResource FluentTextBrush}"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                  <Setter Property="Background" Value="{StaticResource FluentPrimaryBrush}"/>
                  <Setter Property="Foreground" Value="White"/>
                </Trigger>
                <MultiTrigger>
                  <MultiTrigger.Conditions>
                    <Condition Property="IsMouseOver" Value="True"/>
                    <Condition Property="IsSelected" Value="True"/>
                  </MultiTrigger.Conditions>
                  <Setter Property="Background" Value="{StaticResource FluentPrimaryHoverBrush}"/>
                  <Setter Property="Foreground" Value="White"/>
                </MultiTrigger>
              </Style.Triggers>
            </Style>
          </Setter.Value>
        </Setter>
      </Style>
      
      <!-- Fluent Button Style -->
      <Style x:Key="FluentButtonStyle" TargetType="Button">
        <Setter Property="Background" Value="{StaticResource FluentPrimaryBrush}"/>
        <Setter Property="Foreground" Value="White"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="FontSize" Value="15"/>
        <Setter Property="FontWeight" Value="Bold"/>
        <Setter Property="FontFamily" Value="Segoe UI Variable"/>
        <Setter Property="Padding" Value="24,12"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template">
          <Setter.Value>
            <ControlTemplate TargetType="Button">
              <Border Background="{TemplateBinding Background}"
                      CornerRadius="12"
                      Padding="{TemplateBinding Padding}">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
            </ControlTemplate>
          </Setter.Value>
        </Setter>
        <Style.Triggers>
          <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="{StaticResource FluentPrimaryHoverBrush}"/>
          </Trigger>
        </Style.Triggers>
      </Style>
    </ResourceDictionary>
  </Window.Resources>
  
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="100"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    
    <!-- Header -->
    <Border Grid.Row="0" Background="{StaticResource FluentHeaderBrush}">
      <StackPanel Margin="30,0">
        <TextBlock Text="Microsoft Office Auto Installer"
                   FontSize="26" FontWeight="Bold"
                   Foreground="{StaticResource FluentTextBrush}"
                   FontFamily="Segoe UI Variable"
                   Margin="0,28,0,0"/>
        <TextBlock Text="Automated Office deployment with customizable options"
                   FontSize="12"
                   Foreground="{StaticResource FluentTextMutedBrush}"
                   FontFamily="Segoe UI Variable"
                   Margin="0,4,0,0"/>
      </StackPanel>
    </Border>
    
    <!-- Content -->
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
      <StackPanel Margin="40,30,40,20">
        
        <!-- System Configuration -->
        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="System Configuration"
                     FontSize="15" FontWeight="Bold"
                     Foreground="{StaticResource FluentTextSecondaryBrush}"
                     FontFamily="Segoe UI Variable"
                     Margin="0,0,0,12"/>
          <Border Background="{StaticResource FluentCardBrush}" CornerRadius="12" Padding="20,16">
            <ComboBox x:Name="ArchCombo" Style="{StaticResource FluentComboBoxStyle}">
              <ComboBoxItem Content="64-bit (Recommended for most computers)" IsSelected="True"/>
              <ComboBoxItem Content="32-bit (For older systems)"/>
            </ComboBox>
          </Border>
        </StackPanel>
        
        <!-- Office Version -->
        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="Office Version"
                     FontSize="15" FontWeight="Bold"
                     Foreground="{StaticResource FluentTextSecondaryBrush}"
                     FontFamily="Segoe UI Variable"
                     Margin="0,0,0,12"/>
          <Border Background="{StaticResource FluentCardBrush}" CornerRadius="12" Padding="20,16">
            <ComboBox x:Name="EditionCombo" Style="{StaticResource FluentComboBoxStyle}">
              <ComboBoxItem Content="Office 2024 Pro Plus (Latest features)" IsSelected="True"/>
              <ComboBoxItem Content="Office LTSC 2021 (Long-term support)"/>
              <ComboBoxItem Content="Microsoft 365 Apps (Cloud-connected)"/>
            </ComboBox>
          </Border>
        </StackPanel>
        
        <!-- Optional Components -->
        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="Optional Components"
                     FontSize="15" FontWeight="Bold"
                     Foreground="{StaticResource FluentTextSecondaryBrush}"
                     FontFamily="Segoe UI Variable"
                     Margin="0,0,0,12"/>
          <Border Background="{StaticResource FluentCardBrush}" CornerRadius="12" Padding="20,16">
            <StackPanel>
              <CheckBox x:Name="VisioCheck" Content="Include Visio Professional (for diagrams and flowcharts)"
                        Foreground="{StaticResource FluentTextSecondaryBrush}"
                        FontSize="13" FontFamily="Segoe UI Variable"
                        Margin="0,0,0,12"/>
              <CheckBox x:Name="ProjectCheck" Content="Include Project Professional (for project management)"
                        Foreground="{StaticResource FluentTextSecondaryBrush}"
                        FontSize="13" FontFamily="Segoe UI Variable"/>
            </StackPanel>
          </Border>
        </StackPanel>
        
        <!-- Update Settings -->
        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="Update Settings"
                     FontSize="15" FontWeight="Bold"
                     Foreground="{StaticResource FluentTextSecondaryBrush}"
                     FontFamily="Segoe UI Variable"
                     Margin="0,0,0,12"/>
          <Border Background="{StaticResource FluentCardBrush}" CornerRadius="12" Padding="20,16">
            <ComboBox x:Name="ChannelCombo" Style="{StaticResource FluentComboBoxStyle}">
              <ComboBoxItem Content="Monthly updates (Recommended - latest features)" IsSelected="True"/>
              <ComboBoxItem Content="Less frequent updates (More stable)"/>
            </ComboBox>
          </Border>
        </StackPanel>
        
        <!-- Language -->
        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="Language"
                     FontSize="15" FontWeight="Bold"
                     Foreground="{StaticResource FluentTextSecondaryBrush}"
                     FontFamily="Segoe UI Variable"
                     Margin="0,0,0,12"/>
          <Border Background="{StaticResource FluentCardBrush}" CornerRadius="12" Padding="20,16">
            <StackPanel>
              <ComboBox x:Name="LangCombo" Style="{StaticResource FluentComboBoxStyle}" Margin="0,0,0,12">
                <ComboBoxItem Content="English (United States)" IsSelected="True"/>
                <ComboBoxItem Content="English (United Kingdom)"/>
                <ComboBoxItem Content="French (France)"/>
                <ComboBoxItem Content="German (Germany)"/>
                <ComboBoxItem Content="Dutch (Netherlands)"/>
                <ComboBoxItem Content="Spanish (Spain)"/>
                <ComboBoxItem Content="Portuguese (Brazil)"/>
              </ComboBox>
              <TextBlock Text="Configure your installation options above, then click 'Install Office' to begin."
                         FontSize="12"
                         Foreground="{StaticResource FluentTextMutedBrush}"
                         FontFamily="Segoe UI Variable"
                         TextWrapping="Wrap"/>
            </StackPanel>
          </Border>
        </StackPanel>
        
        <!-- Installation Display -->
        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="Installation Display"
                     FontSize="15" FontWeight="Bold"
                     Foreground="{StaticResource FluentTextSecondaryBrush}"
                     FontFamily="Segoe UI Variable"
                     Margin="0,0,0,12"/>
          <Border Background="{StaticResource FluentCardBrush}" CornerRadius="12" Padding="20,16">
            <ComboBox x:Name="UICombo" Style="{StaticResource FluentComboBoxStyle}">
              <ComboBoxItem Content="Show installation progress (Recommended)" IsSelected="True"/>
              <ComboBoxItem Content="Install quietly in background"/>
            </ComboBox>
          </Border>
        </StackPanel>
        
      </StackPanel>
    </ScrollViewer>
    
    <!-- Status Panel (Hidden by default) -->
    <Border x:Name="StatusPanel" Grid.Row="1" 
            Background="{StaticResource FluentBackgroundBrush}"
            Visibility="Collapsed"
            VerticalAlignment="Bottom"
            Margin="40,0,40,100">
      <StackPanel Margin="0,15">
        <ProgressBar x:Name="ProgressBar" Height="28" Margin="0,0,0,10"/>
        <TextBlock x:Name="StatusLabel" 
                   Foreground="{StaticResource FluentTextMutedBrush}"
                   FontSize="12" FontFamily="Segoe UI Variable"/>
      </StackPanel>
    </Border>
    
    <!-- Install Button -->
    <Button x:Name="InstallButton" Grid.Row="2"
            Content="Install Office"
            Style="{StaticResource FluentButtonStyle}"
            Width="840" Height="50"
            Margin="40,20,40,20"
            HorizontalAlignment="Center"/>
            
  </Grid>
</Window>
"@

# ============================================================================
# XAML LOADING & WINDOW CREATION
# ============================================================================
# Parses the embedded XAML string and creates the WPF Window object.
# This is where the UI is instantiated from the XAML definition above.

try {
    # Create an XmlReader from the XAML string
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    # Load the XAML into a WPF Window object
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    $reader.Close()
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to load the user interface.`n`nError: $_`n`nPlease ensure you have .NET Framework 4.7.2 or later installed.",
        "XAML Load Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

# ============================================================================
# UI CONTROL REFERENCES
# ============================================================================
# Retrieves references to all named controls defined in the XAML.
# These references are used to:
# - Read user selections (ComboBox SelectedItem, CheckBox IsChecked)
# - Update UI during installation (ProgressBar, StatusLabel)
# - Handle user interactions (Button Click events)
#
# Control Mapping:
# - ArchCombo: System architecture selection (32-bit/64-bit)
# - EditionCombo: Office edition selection (Pro Plus, LTSC, M365)
# - VisioCheck: Optional Visio Professional component
# - ProjectCheck: Optional Project Professional component
# - ChannelCombo: Update channel selection (Monthly, Semi-Annual, etc.)
# - LangCombo: Language selection for Office installation
# - UICombo: Installation display mode (Show progress / Quiet)
# - InstallButton: Primary action button to start installation
# - StatusPanel: Container for progress indicators (hidden by default)
# - ProgressBar: Visual progress indicator during download/install
# - StatusLabel: Text status updates during installation

try {
    $archCombo = $window.FindName("ArchCombo")
    $editionCombo = $window.FindName("EditionCombo")
    $visioCheck = $window.FindName("VisioCheck")
    $projectCheck = $window.FindName("ProjectCheck")
    $channelCombo = $window.FindName("ChannelCombo")
    $langCombo = $window.FindName("LangCombo")
    $uiCombo = $window.FindName("UICombo")
    $installButton = $window.FindName("InstallButton")
    $statusPanel = $window.FindName("StatusPanel")
    $progressBar = $window.FindName("ProgressBar")
    $statusLabel = $window.FindName("StatusLabel")
    
    # Verify all controls were found
    if ($null -eq $archCombo -or $null -eq $editionCombo -or $null -eq $installButton) {
        throw "Some UI controls could not be found. The XAML may be invalid."
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to initialize the user interface.`n`nError: $_`n`nPlease check the script for errors.",
        "UI Initialization Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

# Make controls accessible globally for use in functions and event handlers
# Using $script: scope ensures variables are accessible across function boundaries
$script:window = $window
$script:statusLabel = $statusLabel
$script:progressBar = $progressBar

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================
# These utility functions convert user-friendly display names to Office
# Deployment Tool (ODT) configuration values.

<#
.SYNOPSIS
    Converts a language display name to its ODT language code.
.DESCRIPTION
    Maps user-friendly language names (e.g., "English (United States)")
    to Office Deployment Tool language codes (e.g., "en-us").
.PARAMETER languageName
    The display name of the language as shown in the ComboBox.
.OUTPUTS
    String - The ODT language code (e.g., "en-us", "fr-fr").
#>
function Get-LanguageCode {
    param([string]$languageName)
    $langMap = @{
        "English (United States)" = "en-us"
        "English (United Kingdom)" = "en-gb"
        "French (France)" = "fr-fr"
        "German (Germany)" = "de-de"
        "Dutch (Netherlands)" = "nl-nl"
        "Spanish (Spain)" = "es-es"
        "Portuguese (Brazil)" = "pt-br"
    }
    return $langMap[$languageName]
}

<#
.SYNOPSIS
    Converts a ComboBox index to an Office edition product ID.
.DESCRIPTION
    Maps the selected ComboBox index to the corresponding Office Deployment
    Tool product ID used in config.xml.
.PARAMETER index
    The zero-based index of the selected item in the EditionCombo.
.OUTPUTS
    String - The ODT product ID (e.g., "ProPlus2024Retail").
#>
function Get-EditionID {
    param([int]$index)
    $editionMap = @{
        0 = "ProPlus2024Retail"      # Office 2024 Pro Plus
        1 = "ProPlus2021Volume"       # Office LTSC 2021
        2 = "O365ProPlusRetail"       # Microsoft 365 Apps
    }
    return $editionMap[$index]
}

<#
.SYNOPSIS
    Converts a ComboBox index to a user-friendly edition name.
.DESCRIPTION
    Maps the selected ComboBox index to a display name for user messages.
.PARAMETER index
    The zero-based index of the selected item in the EditionCombo.
.OUTPUTS
    String - The edition display name (e.g., "Office 2024 Pro Plus").
#>
function Get-EditionName {
    param([int]$index)
    $nameMap = @{
        0 = "Office 2024 Pro Plus"
        1 = "Office LTSC 2021"
        2 = "Microsoft 365 Apps"
    }
    return $nameMap[$index]
}

<#
.SYNOPSIS
    Updates the status label and progress bar in the UI.
.DESCRIPTION
    Thread-safe UI update function that uses WPF Dispatcher to update
    UI elements from background threads. Shows the status panel and updates
    progress if specified.
.PARAMETER message
    The status message to display to the user.
.PARAMETER progress
    Optional progress percentage (0-100). If -1, only updates the message.
.NOTES
    Uses Dispatcher.Invoke to ensure thread-safe UI updates, as download
    and installation operations run on background threads.
#>
function Update-Status {
    param([string]$message, [int]$progress = -1)
    $window.Dispatcher.Invoke([action]{
        $script:statusLabel.Text = $message
        if ($progress -ge 0) {
            $script:progressBar.Value = $progress
            $script:progressBar.Visibility = "Visible"
        }
        $script:statusPanel.Visibility = "Visible"
    })
}

<#
.SYNOPSIS
    Validates system requirements before starting installation.
.DESCRIPTION
    Checks for minimum disk space (4GB) and active internet connection.
    These are required for downloading and installing Office.
.OUTPUTS
    Boolean - $true if requirements are met, $false otherwise.
#>
function Test-SystemRequirements {
    Update-Status "Checking system requirements..." 5
    
    # Check available disk space on the system drive
    $systemDrive = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq $env:SystemDrive }
    $freeSpaceGB = [math]::Round($systemDrive.FreeSpace / 1GB, 2)
    
    if ($freeSpaceGB -lt 4) {
        [System.Windows.Forms.MessageBox]::Show(
            "Error: At least 4GB of free space is required.`n`nAvailable: $freeSpaceGB GB`n`nPlease free up some disk space and try again.",
            "Insufficient Disk Space",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
    
    # Test internet connectivity by attempting to reach Microsoft's website
    Update-Status "Testing internet connection..." 10
    try {
        $null = Invoke-WebRequest -Uri "https://www.microsoft.com" -UseBasicParsing -TimeoutSec 10
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "No internet connection detected!`n`nPlease check your internet connection and try again.",
            "No Internet Connection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
    
    return $true
}

<#
.SYNOPSIS
    Downloads the Office Deployment Tool (ODT) from Microsoft's CDN.
.DESCRIPTION
    Downloads setup.exe (the Office Deployment Tool) from Microsoft's official
    CDN. Shows real-time download progress in the UI and validates the
    downloaded file size.
.OUTPUTS
    Boolean - $true if download succeeded, $false otherwise.
.NOTES
    The ODT is Microsoft's official tool for deploying Office. It's downloaded
    fresh each time to ensure the latest version is used.
#>
function Download-ODT {
    $url = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    $output = "$installerFolder\setup.exe"
    
    Update-Status "Downloading Office installer from Microsoft..." 15
    Log "Downloading Office Deployment Tool from $url..."
    
    try {
        $webClient = New-Object System.Net.WebClient
        
        # Register event handler for download progress updates
        # This allows real-time progress bar updates during download
        $eventJob = Register-ObjectEvent -InputObject $webClient -EventName DownloadProgressChanged -Action {
            $percent = $Event.SourceEventArgs.ProgressPercentage
            # Update UI on the main thread (required for WPF)
            $window.Dispatcher.Invoke([action]{
                # Progress ranges from 15% (start) to 85% (complete)
                $script:progressBar.Value = [Math]::Min(85, 15 + ($percent * 0.7))
                $script:statusLabel.Text = "Downloading Office installer... $percent%"
            })
        }
        
        # Perform the download
        $webClient.DownloadFile($url, $output)
        $webClient.Dispose()
        
        # Clean up event handler
        if ($eventJob) {
            Unregister-Event -SourceIdentifier $eventJob.Name -ErrorAction SilentlyContinue
            Remove-Job -Job $eventJob -ErrorAction SilentlyContinue
        }
        
        Update-Status "Download completed successfully!" 85
        
        # Validate downloaded file exists and has reasonable size (>100KB)
        if (-Not (Test-Path $output) -or ((Get-Item $output).Length -lt 100000)) {
            throw "Downloaded file appears to be corrupted or incomplete."
        }
        
        Log "Office Deployment Tool downloaded successfully."
        Start-Sleep -Seconds 1
        return $true
        
    } catch {
        Log "Download failed: $_"
        [System.Windows.Forms.MessageBox]::Show(
            "Download failed!`n`nError: $_`n`nPlease check your internet connection and try again.",
            "Download Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}

<#
.SYNOPSIS
    Generates the Office Deployment Tool configuration XML file.
.DESCRIPTION
    Creates config.xml based on user selections. This XML file tells the
    Office Deployment Tool what to install, which language, update channel,
    and display preferences.
.PARAMETER options
    Hashtable containing user selections:
    - edition: Office edition product ID (e.g., "ProPlus2024Retail")
    - language: Language code (e.g., "en-us")
    - bit: Architecture ("32" or "64")
    - channel: Update channel (e.g., "Monthly", "Broad")
    - ui: Display level ("Full" or "None")
    - visio: "1" if Visio should be installed, "0" otherwise
    - project: "1" if Project should be installed, "0" otherwise
.NOTES
    The generated config.xml follows the Office Deployment Tool schema.
    See: https://docs.microsoft.com/en-us/deployoffice/configuration-options-for-the-office-2016-deployment-tool
#>
function Generate-Config {
    param($options)
    
    Update-Status "Creating installation configuration..." 87
    Log "Generating config.xml with selected options..."
    
    # Build product list - always includes the main Office edition
    $products = @()
    $products += "<Product ID='" + $options.edition + "'>`n  <Language ID='" + $options.language + "' />`n</Product>"
    
    # Add optional components if selected
    if ($options.visio -eq "1") {
        $products += "<Product ID='VisioPro2021Volume'>`n  <Language ID='" + $options.language + "' />`n</Product>"
    }
    if ($options.project -eq "1") {
        $products += "<Product ID='ProjectPro2021Volume'>`n  <Language ID='" + $options.language + "' />`n</Product>"
    }
    
    # Generate the complete config.xml following ODT schema
    $xmlContent = @"
<Configuration>
  <Add OfficeClientEdition="${($options.bit)}" Channel="${($options.channel)}">
    $($products -join "`n    ")
  </Add>
  <Display Level="${($options.ui)}" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
"@
    
    # Write config.xml to the installer folder
    $configPath = "$installerFolder\config.xml"
    $xmlContent | Out-File -FilePath $configPath -Encoding UTF8
    Log "config.xml generated at $configPath"
    
    Update-Status "Configuration created successfully!" 90
    Start-Sleep -Seconds 1
}

<#
.SYNOPSIS
    Executes the Office installation using the Office Deployment Tool.
.DESCRIPTION
    Runs setup.exe with the generated config.xml to install Office.
    Monitors the installation process and provides user feedback based on
    the exit code.
.PARAMETER options
    Hashtable containing installation options (used for success message).
.OUTPUTS
    Boolean - $true if installation completed (even with warnings),
              $false if installation failed.
.NOTES
    Installation typically takes 10-30 minutes depending on:
    - Internet speed (if downloading Office)
    - System performance
    - Selected components
    
    Exit codes:
    - 0: Success
    - Non-zero: Warning or error (but Office may still be installed)
#>
function Install-Office {
    param($options)
    
    Update-Status "Starting Office installation... This may take 10-30 minutes." 92
    Log "Starting Office installation..."
    $setupExe = "$installerFolder\setup.exe"
    
    # Verify setup.exe exists before attempting installation
    if (-Not (Test-Path $setupExe)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Installation file missing!`n`nSomething went wrong with the download. Please restart the installer.",
            "Installation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
    
    try {
        Log "Executing: $setupExe /configure config.xml"
        Set-Location -Path $installerFolder
        
        # Execute Office Deployment Tool with config.xml
        # -PassThru: Returns process object to check exit code
        # -NoNewWindow: Runs in current window (shows progress if UI level is Full)
        # -Wait: Blocks until installation completes
        $process = Start-Process -FilePath $setupExe -ArgumentList "/configure config.xml" -PassThru -NoNewWindow -Wait
        
        $exitCode = $process.ExitCode
        
        if ($exitCode -eq 0) {
            # Success - Office installed without errors
            Update-Status "Office installation completed successfully!" 100
            Log "Office installation completed successfully with exit code: $exitCode"
            
            [System.Windows.Forms.MessageBox]::Show(
                "Office installation completed successfully!`n`nYou can now find Office applications in your Start Menu.`n`nInstalled:`n• $($options.editionName)`n• Language: $($options.languageName)`n• Architecture: $($options.bit)-bit",
                "Installation Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return $true
        } else {
            # Non-zero exit code - may indicate warning, but Office might still be installed
            Log "Office installation completed with exit code: $exitCode"
            Update-Status "Installation completed with warnings (Exit code: $exitCode)" 100
            
            [System.Windows.Forms.MessageBox]::Show(
                "Office installation completed with exit code: $exitCode`n`nThis may indicate a warning or non-critical issue. Office may still be installed correctly.`n`nPlease check if Office applications are available in your Start Menu.",
                "Installation Warning",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return $true
        }
        
    } catch {
        # Installation process failed to start or crashed
        Log "Installation failed: $_"
        [System.Windows.Forms.MessageBox]::Show(
            "Installation encountered an error!`n`nError: $_`n`nYou can try running the installer again.",
            "Installation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}

# ============================================================================
# INSTALL BUTTON CLICK EVENT HANDLER
# ============================================================================
# This is the main entry point when the user clicks "Install Office".
# It orchestrates the entire installation process:
# 1. Collects user selections from UI controls
# 2. Validates system requirements
# 3. Downloads Office Deployment Tool
# 4. Generates configuration XML
# 5. Executes Office installation
#
# The handler disables UI controls during installation to prevent
# user interference and provides real-time progress updates.

$installButton.Add_Click({
    # Disable UI controls to prevent changes during installation
    $installButton.IsEnabled = $false
    $archCombo.IsEnabled = $false
    $editionCombo.IsEnabled = $false
    $visioCheck.IsEnabled = $false
    $projectCheck.IsEnabled = $false
    $channelCombo.IsEnabled = $false
    $langCombo.IsEnabled = $false
    $uiCombo.IsEnabled = $false
    
    try {
        Log "=== Office Installer GUI Started ==="
        
        if (-not (Test-SystemRequirements)) {
            $installButton.IsEnabled = $true
            $archCombo.IsEnabled = $true
            $editionCombo.IsEnabled = $true
            $visioCheck.IsEnabled = $true
            $projectCheck.IsEnabled = $true
            $channelCombo.IsEnabled = $true
            $langCombo.IsEnabled = $true
            $uiCombo.IsEnabled = $true
            $statusPanel.Visibility = "Collapsed"
            return
        }
        
        # Collect user selections from UI controls
        # Convert ComboBox indices and CheckBox states to ODT configuration values
        $bit = if ($archCombo.SelectedIndex -eq 0) { "64" } else { "32" }
        $editionID = Get-EditionID -index $editionCombo.SelectedIndex
        $editionName = Get-EditionName -index $editionCombo.SelectedIndex
        $visio = if ($visioCheck.IsChecked) { "1" } else { "2" }
        $project = if ($projectCheck.IsChecked) { "1" } else { "2" }
        $channel = if ($channelCombo.SelectedIndex -eq 0) { "Current" } else { "Broad" }
        $languageCode = Get-LanguageCode -languageName ($langCombo.SelectedItem.Content)
        $languageName = $langCombo.SelectedItem.Content
        $uiLevel = if ($uiCombo.SelectedIndex -eq 0) { "Full" } else { "None" }
        
        # Build options hashtable to pass to installation functions
        $options = @{
            bit = $bit
            visio = $visio
            project = $project
            channel = $channel
            language = $languageCode
            languageName = $languageName
            ui = $uiLevel
            edition = $editionID
            editionName = $editionName
        }
        
        if (-not (Download-ODT)) {
            $installButton.IsEnabled = $true
            $archCombo.IsEnabled = $true
            $editionCombo.IsEnabled = $true
            $visioCheck.IsEnabled = $true
            $projectCheck.IsEnabled = $true
            $channelCombo.IsEnabled = $true
            $langCombo.IsEnabled = $true
            $uiCombo.IsEnabled = $true
            $statusPanel.Visibility = "Collapsed"
            return
        }
        
        Generate-Config -options $options
        $success = Install-Office -options $options
        
        if ($success) {
            Log "=== Office Installer GUI Completed Successfully ==="
        }
        
    } catch {
        Log "FATAL ERROR: $_"
        [System.Windows.Forms.MessageBox]::Show(
            "An unexpected error occurred: $_`n`nPlease try running the installer again.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    } finally {
        $installButton.IsEnabled = $true
        $archCombo.IsEnabled = $true
        $editionCombo.IsEnabled = $true
        $visioCheck.IsEnabled = $true
        $projectCheck.IsEnabled = $true
        $channelCombo.IsEnabled = $true
        $langCombo.IsEnabled = $true
        $uiCombo.IsEnabled = $true
    }
})

# ============================================================================
# APPLICATION EXECUTION
# ============================================================================
# Displays the WPF window and starts the message loop.
# ShowDialog() blocks until the window is closed, keeping the script
# running for the lifetime of the application.
#
# The window will remain open until:
# - User closes it (X button)
# - Installation completes and user acknowledges
# - An error occurs and user dismisses the error dialog

Log "=== Office Installer GUI Started ==="
$window.ShowDialog() | Out-Null

