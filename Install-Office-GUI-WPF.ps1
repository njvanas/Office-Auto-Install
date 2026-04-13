#Requires -Version 5.1
# Office Auto Installer - WPF UI aligned with the GitHub Pages site (slate / blue)
# Single-file, self-contained Office installer with WPF
# Downloads and installs Microsoft Office through official channels
# Updated release workflow to include both GUI and Console versions

<#
.SYNOPSIS
    Microsoft Office Auto Installer - WPF (site-themed UI)
.DESCRIPTION
    WPF installer window styled to match the public site: slate gradient, blue accents, Inter font when installed.
    All options are pre-filled with recommended defaults for one-click installation.
    Fully self-contained - no external dependencies required.
.NOTES
    Version: 3.8 - Standard M365 retail profiles (generated XML) + custom advanced path
    Author: Office Auto Installer Team
    Requires: .NET Framework 4.7.2+ (Windows PowerShell; WPF is not available in PowerShell 7+)
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

# Optional PresentationFramework.Fluent merge (.NET 9+ only; unused on .NET Framework)
$fluentThemeAvailable = $false
# UI uses embedded site-themed brushes/styles (matches index.html)

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

$corePath = Join-Path $PSScriptRoot 'M365AppsCore.ps1'
if (-not (Test-Path -LiteralPath $corePath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Missing M365AppsCore.ps1 (expected next to this script):`n$corePath",
        "Office Auto Install",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}
. $corePath

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
# 1. Window: gradient background matching site (slate-900 / blue-900)
# 2. Resources: Tailwind-aligned brushes; ComboBox / Button / ProgressBar styles
# 3. Layout: header (nav-style branding), scrollable content, footer bar with primary action
# 4. Controls: ComboBoxes, CheckBoxes; status strip during install
#
# Design Philosophy (matches index.html):
# - Inter, Segoe UI fallback (same stack as the site; Inter if installed on the PC)
# - slate-950 cards, slate-600 borders, blue-600 primary, blue-400 accents / focus
# - Custom ComboBox template (no default light chrome)
#
# Note: x:Class is intentionally omitted (dynamic XAML load in PowerShell).

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Office Auto Installer"
        Width="920" Height="1120"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
  <Window.Background>
    <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
      <GradientStop Color="#0F172A" Offset="0"/>
      <GradientStop Color="#1E3A8A" Offset="0.5"/>
      <GradientStop Color="#0F172A" Offset="1"/>
    </LinearGradientBrush>
  </Window.Background>
  <Window.Resources>
    <ResourceDictionary>
      $(if ($fluentThemeAvailable) {
        @"
      <ResourceDictionary.MergedDictionaries>
        <ResourceDictionary Source="pack://application:,,,/PresentationFramework.Fluent;component/Themes/Fluent.xaml"/>
      </ResourceDictionary.MergedDictionaries>
"@
      })

      <!-- GitHub Pages / Tailwind palette (see index.html) -->
      <SolidColorBrush x:Key="SitePanelBrush" Color="#0F172A"/>
      <SolidColorBrush x:Key="SiteCardBrush" Color="#020617"/>
      <SolidColorBrush x:Key="SiteHeaderBrush" Color="#0F172A"/>
      <SolidColorBrush x:Key="SiteTextBrush" Color="#FFFFFF"/>
      <SolidColorBrush x:Key="SiteTextBodyBrush" Color="#D1D5DB"/>
      <SolidColorBrush x:Key="SiteTextMutedBrush" Color="#9CA3AF"/>
      <SolidColorBrush x:Key="SitePrimaryBrush" Color="#2563EB"/>
      <SolidColorBrush x:Key="SitePrimaryHoverBrush" Color="#1D4ED8"/>
      <SolidColorBrush x:Key="SiteAccentBrush" Color="#60A5FA"/>
      <SolidColorBrush x:Key="SiteBorderBrush" Color="#475569"/>
      <SolidColorBrush x:Key="SiteControlBrush" Color="#1E293B"/>
      <SolidColorBrush x:Key="SiteItemHoverBrush" Color="#334155"/>

      <Style x:Key="SiteComboBoxStyle" TargetType="ComboBox">
        <Setter Property="MinHeight" Value="36"/>
        <Setter Property="Padding" Value="12,8"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Background" Value="{StaticResource SiteControlBrush}"/>
        <Setter Property="Foreground" Value="{StaticResource SiteTextBodyBrush}"/>
        <Setter Property="BorderBrush" Value="{StaticResource SiteBorderBrush}"/>
        <Setter Property="FontSize" Value="13"/>
        <Setter Property="FontFamily" Value="Inter, Segoe UI"/>
        <Setter Property="Template">
          <Setter.Value>
            <ControlTemplate TargetType="ComboBox">
              <Grid>
                <Border x:Name="FocusBorder"
                        BorderBrush="{StaticResource SiteAccentBrush}"
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
                      Fill="{StaticResource SiteTextMutedBrush}"
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
                          Background="{StaticResource SiteCardBrush}"
                          BorderBrush="{StaticResource SiteBorderBrush}"
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
                  <Setter TargetName="Border" Property="BorderBrush" Value="{StaticResource SiteAccentBrush}"/>
                  <Setter TargetName="Border" Property="BorderThickness" Value="1"/>
                </Trigger>
                <Trigger Property="IsKeyboardFocusWithin" Value="False">
                  <Setter TargetName="FocusBorder" Property="BorderThickness" Value="0"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                  <Setter Property="Foreground" Value="{StaticResource SiteTextMutedBrush}"/>
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
              <Setter Property="Foreground" Value="{StaticResource SiteTextBodyBrush}"/>
              <Setter Property="FontSize" Value="13"/>
              <Setter Property="FontFamily" Value="Inter, Segoe UI"/>
              <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                  <Setter Property="Background" Value="{StaticResource SiteItemHoverBrush}"/>
                  <Setter Property="Foreground" Value="{StaticResource SiteTextBrush}"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                  <Setter Property="Background" Value="{StaticResource SitePrimaryBrush}"/>
                  <Setter Property="Foreground" Value="White"/>
                </Trigger>
                <MultiTrigger>
                  <MultiTrigger.Conditions>
                    <Condition Property="IsMouseOver" Value="True"/>
                    <Condition Property="IsSelected" Value="True"/>
                  </MultiTrigger.Conditions>
                  <Setter Property="Background" Value="{StaticResource SitePrimaryHoverBrush}"/>
                  <Setter Property="Foreground" Value="White"/>
                </MultiTrigger>
              </Style.Triggers>
            </Style>
          </Setter.Value>
        </Setter>
      </Style>

      <Style x:Key="SiteButtonStyle" TargetType="Button">
        <Setter Property="Background" Value="{StaticResource SitePrimaryBrush}"/>
        <Setter Property="Foreground" Value="White"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="FontSize" Value="15"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="FontFamily" Value="Inter, Segoe UI"/>
        <Setter Property="Padding" Value="24,12"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template">
          <Setter.Value>
            <ControlTemplate TargetType="Button">
              <Border Background="{TemplateBinding Background}"
                      CornerRadius="8"
                      Padding="{TemplateBinding Padding}">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
            </ControlTemplate>
          </Setter.Value>
        </Setter>
        <Style.Triggers>
          <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="{StaticResource SitePrimaryHoverBrush}"/>
          </Trigger>
        </Style.Triggers>
      </Style>

      <Style x:Key="SiteCheckBoxStyle" TargetType="CheckBox">
        <Setter Property="Foreground" Value="{StaticResource SiteTextBodyBrush}"/>
        <Setter Property="FontSize" Value="13"/>
        <Setter Property="FontFamily" Value="Inter, Segoe UI"/>
      </Style>

      <Style x:Key="SiteProgressBarStyle" TargetType="ProgressBar">
        <Setter Property="Foreground" Value="{StaticResource SitePrimaryBrush}"/>
        <Setter Property="Background" Value="{StaticResource SiteControlBrush}"/>
        <Setter Property="BorderBrush" Value="{StaticResource SiteBorderBrush}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Height" Value="10"/>
      </Style>
    </ResourceDictionary>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header (site nav: logo tile + title) -->
    <Border Grid.Row="0"
            Background="{StaticResource SiteHeaderBrush}"
            BorderBrush="{StaticResource SiteBorderBrush}"
            BorderThickness="0,0,0,1"
            Padding="24,16,24,16">
      <StackPanel Orientation="Horizontal">
        <Border Width="32" Height="32" CornerRadius="8" Margin="0,2,12,0" VerticalAlignment="Top">
          <Border.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
              <GradientStop Color="#2563EB" Offset="0"/>
              <GradientStop Color="#60A5FA" Offset="1"/>
            </LinearGradientBrush>
          </Border.Background>
          <TextBlock Text="O"
                     Foreground="White"
                     FontWeight="Bold"
                     FontSize="16"
                     FontFamily="Inter, Segoe UI"
                     HorizontalAlignment="Center"
                     VerticalAlignment="Center"/>
        </Border>
        <StackPanel VerticalAlignment="Center">
          <TextBlock Text="Office Auto Installer"
                     FontSize="24"
                     FontWeight="Bold"
                     Foreground="{StaticResource SiteAccentBrush}"
                     FontFamily="Inter, Segoe UI"/>
          <TextBlock Text="Offline deployment settings (like Microsoft 365 admin center) — generates ODT XML and installs with setup.exe"
                     FontSize="12"
                     Foreground="{StaticResource SiteTextMutedBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,4,0,0"
                     TextWrapping="Wrap"/>
        </StackPanel>
      </StackPanel>
    </Border>

    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Disabled" Background="Transparent">
      <TabControl x:Name="MainTabControl" Margin="28,20,28,16" Background="Transparent" BorderThickness="0">
        <TabControl.Resources>
          <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="#E2E8F0"/>
            <Setter Property="FontFamily" Value="Inter, Segoe UI"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="14,10"/>
            <Setter Property="Margin" Value="0,0,4,0"/>
          </Style>
        </TabControl.Resources>
        <TabItem Header="Products &amp; apps">
          <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="0,12,8,0">
            <StackPanel Margin="12,4,12,20">

        <StackPanel Margin="0,0,0,20">
          <TextBlock Text="Products"
                     FontSize="15" FontWeight="SemiBold"
                     Foreground="{StaticResource SiteTextBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,0,0,12"/>
          <TextBlock Text="Choose the Microsoft 365 Apps offering, or Other products for Office LTSC / Office 2024 / Visio-only scenarios (same choices as the online deployment settings, running locally)."
                     FontSize="12"
                     Foreground="{StaticResource SiteTextMutedBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,0,0,10"
                     TextWrapping="Wrap"/>
          <Border Background="{StaticResource SiteCardBrush}"
                  BorderBrush="{StaticResource SiteBorderBrush}"
                  BorderThickness="1"
                  CornerRadius="12"
                  Padding="20,16">
            <StackPanel>
              <ComboBox x:Name="ProductSuiteCombo" Style="{StaticResource SiteComboBoxStyle}">
                <ComboBoxItem Content="Microsoft 365 Apps for enterprise" IsSelected="True"/>
                <ComboBoxItem Content="Microsoft 365 Apps for business"/>
                <ComboBoxItem Content="Other products (Office 2024, LTSC, Visio/Project only, …)"/>
              </ComboBox>
              <StackPanel x:Name="DeploymentTargetPanel" Margin="0,16,0,0">
                <TextBlock Text="How will this installation be used?"
                           FontSize="12"
                           Foreground="{StaticResource SiteTextMutedBrush}"
                           FontFamily="Inter, Segoe UI"
                           Margin="0,0,0,8"
                           TextWrapping="Wrap"/>
                <ComboBox x:Name="DeploymentTargetCombo" Style="{StaticResource SiteComboBoxStyle}">
                  <ComboBoxItem Content="This device (desktop or laptop)" IsSelected="True"/>
                  <ComboBoxItem Content="Shared computer / virtual desktop (VDI, Azure Virtual Desktop, Windows 365, RDS)"/>
                </ComboBox>
              </StackPanel>
            </StackPanel>
          </Border>
        </StackPanel>

        <StackPanel x:Name="EditionSection" Margin="0,0,0,24" Visibility="Collapsed">
          <TextBlock Text="Office Version (custom only)"
                     FontSize="15" FontWeight="SemiBold"
                     Foreground="{StaticResource SiteTextBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,0,0,12"/>
          <Border Background="{StaticResource SiteCardBrush}"
                  BorderBrush="{StaticResource SiteBorderBrush}"
                  BorderThickness="1"
                  CornerRadius="12"
                  Padding="20,16">
            <ComboBox x:Name="EditionCombo" Style="{StaticResource SiteComboBoxStyle}">
              <ComboBoxItem Content="Office 2024 Pro Plus (Latest features)" IsSelected="True"/>
              <ComboBoxItem Content="Office LTSC 2021 (Long-term support)"/>
              <ComboBoxItem Content="Microsoft 365 Apps (Cloud-connected)"/>
              <ComboBoxItem Content="Visio and/or Project only (no Word/Excel suite)"/>
            </ComboBox>
          </Border>
        </StackPanel>

        <StackPanel x:Name="OptionalSection" Margin="0,0,0,24" Visibility="Visible">
          <TextBlock Text="Optional: Visio / Project"
                     FontSize="15" FontWeight="SemiBold"
                     Foreground="{StaticResource SiteTextBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,0,0,12"/>
          <TextBlock Text="Add either, both, or neither. Edition below applies to any Visio/Project you include (not to the Office suite)."
                     FontSize="12"
                     Foreground="{StaticResource SiteTextMutedBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,0,0,10"
                     TextWrapping="Wrap"/>
          <Border Background="{StaticResource SiteCardBrush}"
                  BorderBrush="{StaticResource SiteBorderBrush}"
                  BorderThickness="1"
                  CornerRadius="12"
                  Padding="20,16">
            <StackPanel>
              <CheckBox x:Name="VisioCheck" Style="{StaticResource SiteCheckBoxStyle}"
                        Content="Include Visio Professional (for diagrams and flowcharts)"
                        Margin="0,0,0,12"/>
              <CheckBox x:Name="ProjectCheck" Style="{StaticResource SiteCheckBoxStyle}"
                        Content="Include Project Professional (for project management)"/>
              <StackPanel x:Name="VisioProjectLinePanel" Margin="0,16,0,0" Visibility="Collapsed">
                <TextBlock Text="Visio / Project product line (when either is checked)"
                           FontSize="12"
                           Foreground="{StaticResource SiteTextMutedBrush}"
                           FontFamily="Inter, Segoe UI"
                           Margin="0,0,0,8"
                           TextWrapping="Wrap"/>
                <ComboBox x:Name="VisioProjectLineCombo" Style="{StaticResource SiteComboBoxStyle}">
                  <ComboBoxItem Content="Microsoft 365 subscription (Visio Plan 2 / Project)" Tag="M365Retail" IsSelected="True"/>
                  <ComboBoxItem Content="Office LTSC 2021 (volume license)" Tag="LTSC2021Volume"/>
                  <ComboBoxItem Content="Office LTSC 2024 (volume license)" Tag="LTSC2024Volume"/>
                  <ComboBoxItem Content="Office 2024 (retail perpetual)" Tag="Office2024Retail"/>
                </ComboBox>
              </StackPanel>
            </StackPanel>
          </Border>
        </StackPanel>

        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="Exclude apps (optional)"
                     FontSize="15" FontWeight="SemiBold"
                     Foreground="{StaticResource SiteTextBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,0,0,12"/>
          <TextBlock Text="Uncheck apps you do not want in the Microsoft 365 Apps suite (ODT ExcludeApp). Enterprise/business packages include baseline excludes; shared-computer (VDI) adds stricter defaults. Your choices are merged into the XML."
                     FontSize="12"
                     Foreground="{StaticResource SiteTextMutedBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,0,0,10"
                     TextWrapping="Wrap"/>
          <Border Background="{StaticResource SiteCardBrush}"
                  BorderBrush="{StaticResource SiteBorderBrush}"
                  BorderThickness="1"
                  CornerRadius="12"
                  Padding="20,16">
            <WrapPanel x:Name="ExcludeAppsPanel" />
          </Border>
        </StackPanel>

            </StackPanel>
          </ScrollViewer>
        </TabItem>

        <TabItem Header="Languages">
          <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="0,12,8,0">
            <StackPanel Margin="12,4,12,20">
              <TextBlock Text="Primary language"
                         FontSize="15" FontWeight="SemiBold"
                         Foreground="{StaticResource SiteTextBrush}"
                         FontFamily="Inter, Segoe UI"
                         Margin="0,0,0,12"/>
              <Border Background="{StaticResource SiteCardBrush}"
                      BorderBrush="{StaticResource SiteBorderBrush}"
                      BorderThickness="1"
                      CornerRadius="12"
                      Padding="20,16">
                <StackPanel>
                  <ComboBox x:Name="LangCombo" Style="{StaticResource SiteComboBoxStyle}" Margin="0,0,0,12" IsEditable="False"/>
                  <TextBlock Text="Filtered when Visio or Project are included (Microsoft-supported combinations only)."
                             FontSize="12"
                             Foreground="{StaticResource SiteTextMutedBrush}"
                             FontFamily="Inter, Segoe UI"
                             TextWrapping="Wrap"/>
                </StackPanel>
              </Border>
              <TextBlock Text="Additional languages (optional)"
                         FontSize="15" FontWeight="SemiBold"
                         Foreground="{StaticResource SiteTextBrush}"
                         FontFamily="Inter, Segoe UI"
                         Margin="0,20,0,12"/>
              <TextBlock Text="Ctrl+click to multi-select. Added as extra Language elements in the ODT configuration."
                         FontSize="12"
                         Foreground="{StaticResource SiteTextMutedBrush}"
                         FontFamily="Inter, Segoe UI"
                         Margin="0,0,0,10"
                         TextWrapping="Wrap"/>
              <Border Background="{StaticResource SiteCardBrush}"
                      BorderBrush="{StaticResource SiteBorderBrush}"
                      BorderThickness="1"
                      CornerRadius="12"
                      Padding="12,12">
                <ListBox x:Name="AdditionalLangList"
                         SelectionMode="Extended"
                         MinHeight="260"
                         MaxHeight="360"
                         Background="{StaticResource SiteControlBrush}"
                         BorderBrush="{StaticResource SiteBorderBrush}"
                         Foreground="{StaticResource SiteTextBodyBrush}"
                         FontFamily="Inter, Segoe UI"
                         FontSize="13"/>
              </Border>
            </StackPanel>
          </ScrollViewer>
        </TabItem>

        <TabItem Header="Updates">
          <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="0,12,8,0">
            <StackPanel Margin="12,4,12,20">
              <TextBlock Text="Update channel"
                         FontSize="15" FontWeight="SemiBold"
                         Foreground="{StaticResource SiteTextBrush}"
                         FontFamily="Inter, Segoe UI"
                         Margin="0,0,0,12"/>
              <TextBlock Text="Microsoft 365 Apps for enterprise/business can follow the recommended channel for your selection, or use an override."
                         FontSize="12"
                         Foreground="{StaticResource SiteTextMutedBrush}"
                         FontFamily="Inter, Segoe UI"
                         Margin="0,0,0,10"
                         TextWrapping="Wrap"/>
              <Border Background="{StaticResource SiteCardBrush}"
                      BorderBrush="{StaticResource SiteBorderBrush}"
                      BorderThickness="1"
                      CornerRadius="12"
                      Padding="20,16">
                <ComboBox x:Name="ChannelCombo" Style="{StaticResource SiteComboBoxStyle}">
                  <ComboBoxItem x:Name="ChannelProfileDefaultItem" Content="Use recommended default channel for this product" IsSelected="True"/>
                  <ComboBoxItem Content="Monthly / Current (override)"/>
                  <ComboBoxItem Content="Semi-annual Enterprise (override)"/>
                </ComboBox>
              </Border>
              <TextBlock Text="Microsoft apps updates"
                         FontSize="15" FontWeight="SemiBold"
                         Foreground="{StaticResource SiteTextBrush}"
                         FontFamily="Inter, Segoe UI"
                         Margin="0,24,0,12"/>
              <Border Background="{StaticResource SiteCardBrush}"
                      BorderBrush="{StaticResource SiteBorderBrush}"
                      BorderThickness="1"
                      CornerRadius="12"
                      Padding="20,16">
                <StackPanel>
                  <CheckBox x:Name="UpdatesEnabledCheck" Style="{StaticResource SiteCheckBoxStyle}"
                            Content="Enable updates (ODT Updates Enabled=TRUE)"
                            IsChecked="True"
                            Margin="0,0,0,16"/>
                  <TextBlock Text="Target version (optional)"
                             FontSize="12"
                             Foreground="{StaticResource SiteTextMutedBrush}"
                             FontFamily="Inter, Segoe UI"
                             Margin="0,0,0,6"/>
                  <TextBox x:Name="UpdatesTargetVersionBox"
                           MinHeight="36"
                           Padding="10,8"
                           FontFamily="Consolas, Segoe UI"
                           FontSize="12"
                           Background="{StaticResource SiteControlBrush}"
                           Foreground="{StaticResource SiteTextBodyBrush}"
                           BorderBrush="{StaticResource SiteBorderBrush}"
                           BorderThickness="1"
                           Margin="0,0,0,14"/>
                  <TextBlock Text="Update deadline (optional, ODT Deadline format)"
                             FontSize="12"
                             Foreground="{StaticResource SiteTextMutedBrush}"
                             FontFamily="Inter, Segoe UI"
                             Margin="0,0,0,6"/>
                  <TextBox x:Name="UpdatesDeadlineBox"
                           MinHeight="36"
                           Padding="10,8"
                           FontFamily="Consolas, Segoe UI"
                           FontSize="12"
                           Background="{StaticResource SiteControlBrush}"
                           Foreground="{StaticResource SiteTextBodyBrush}"
                           BorderBrush="{StaticResource SiteBorderBrush}"
                           BorderThickness="1"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </ScrollViewer>
        </TabItem>

        <TabItem Header="Installation">
          <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="0,12,8,0">
            <StackPanel Margin="12,4,12,20">
              <TextBlock Text="Architecture"
                         FontSize="15" FontWeight="SemiBold"
                         Foreground="{StaticResource SiteTextBrush}"
                         FontFamily="Inter, Segoe UI"
                         Margin="0,0,0,12"/>
              <Border Background="{StaticResource SiteCardBrush}"
                      BorderBrush="{StaticResource SiteBorderBrush}"
                      BorderThickness="1"
                      CornerRadius="12"
                      Padding="20,16">
                <ComboBox x:Name="ArchCombo" Style="{StaticResource SiteComboBoxStyle}">
                  <ComboBoxItem Content="64-bit (recommended)" IsSelected="True"/>
                  <ComboBoxItem Content="32-bit (older systems)"/>
                </ComboBox>
              </Border>
              <StackPanel x:Name="SharedComputerCustomPanel" Margin="0,20,0,0" Visibility="Collapsed">
                <TextBlock Text="Licensing (custom path: Microsoft 365 Apps only)"
                           FontSize="15" FontWeight="SemiBold"
                           Foreground="{StaticResource SiteTextBrush}"
                           FontFamily="Inter, Segoe UI"
                           Margin="0,0,0,12"/>
                <Border Background="{StaticResource SiteCardBrush}"
                        BorderBrush="{StaticResource SiteBorderBrush}"
                        BorderThickness="1"
                        CornerRadius="12"
                        Padding="20,16">
                  <CheckBox x:Name="SharedComputerCustomCheck" Style="{StaticResource SiteCheckBoxStyle}">
                    <TextBlock Text="Shared computer activation (VDI, Azure Virtual Desktop, Windows 365) — SharedComputerLicensing and VDI-style suite excludes"
                               TextWrapping="Wrap"
                               Foreground="{StaticResource SiteTextBodyBrush}"
                               FontFamily="Inter, Segoe UI"
                               FontSize="13"/>
                  </CheckBox>
                </Border>
              </StackPanel>
              <TextBlock Text="Installation display"
                         FontSize="15" FontWeight="SemiBold"
                         Foreground="{StaticResource SiteTextBrush}"
                         FontFamily="Inter, Segoe UI"
                         Margin="0,24,0,12"/>
              <Border Background="{StaticResource SiteCardBrush}"
                      BorderBrush="{StaticResource SiteBorderBrush}"
                      BorderThickness="1"
                      CornerRadius="12"
                      Padding="20,16">
                <ComboBox x:Name="UICombo" Style="{StaticResource SiteComboBoxStyle}">
                  <ComboBoxItem Content="Show installation progress (recommended)" IsSelected="True"/>
                  <ComboBoxItem Content="Install quietly in the background"/>
                </ComboBox>
              </Border>
            </StackPanel>
          </ScrollViewer>
        </TabItem>
      </TabControl>
    </ScrollViewer>

    <Border x:Name="StatusPanel" Grid.Row="1"
            Background="{StaticResource SitePanelBrush}"
            Visibility="Collapsed"
            VerticalAlignment="Bottom"
            Margin="40,0,40,100"
            BorderBrush="{StaticResource SiteBorderBrush}"
            BorderThickness="0,1,0,0"
            Padding="0,12,0,0">
      <StackPanel Margin="0,4">
        <ProgressBar x:Name="ProgressBar" Style="{StaticResource SiteProgressBarStyle}" Margin="0,0,0,10"/>
        <TextBlock x:Name="StatusLabel"
                   Foreground="{StaticResource SiteTextMutedBrush}"
                   FontSize="12"
                   FontFamily="Inter, Segoe UI"/>
      </StackPanel>
    </Border>

    <Border Grid.Row="2"
            Background="{StaticResource SiteHeaderBrush}"
            BorderBrush="{StaticResource SiteBorderBrush}"
            BorderThickness="0,1,0,0"
            Padding="0,20,0,20">
      <Button x:Name="InstallButton"
              Content="Install Office"
              Style="{StaticResource SiteButtonStyle}"
              Width="840"
              Height="48"
              HorizontalAlignment="Center"/>
    </Border>

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
# - MainTabControl: Portal-style sections (Products & apps, Languages, Updates, Installation)
# - ProductSuiteCombo / DeploymentTargetCombo: M365 enterprise/business + device vs shared VDI
# - ArchCombo: 32/64-bit; EditionCombo: custom products only
# - VisioCheck / ProjectCheck, VisioProjectLineCombo: optional add-ons
# - ChannelCombo, Updates* controls: update channel and ODT Updates element
# - LangCombo + AdditionalLangList: primary and additional languages
# - SharedComputerCustomCheck: VDI licensing on custom M365 Apps path only
# - UICombo: setup display level
# - InstallButton: Primary action button to start installation
# - StatusPanel: Container for progress indicators (hidden by default)
# - ProgressBar: Visual progress indicator during download/install
# - StatusLabel: Text status updates during installation

try {
    $mainTabControl = $window.FindName("MainTabControl")
    $productSuiteCombo = $window.FindName("ProductSuiteCombo")
    $deploymentTargetPanel = $window.FindName("DeploymentTargetPanel")
    $deploymentTargetCombo = $window.FindName("DeploymentTargetCombo")
    $editionSection = $window.FindName("EditionSection")
    $optionalSection = $window.FindName("OptionalSection")
    $archCombo = $window.FindName("ArchCombo")
    $editionCombo = $window.FindName("EditionCombo")
    $visioCheck = $window.FindName("VisioCheck")
    $projectCheck = $window.FindName("ProjectCheck")
    $visioProjectLinePanel = $window.FindName("VisioProjectLinePanel")
    $visioProjectLineCombo = $window.FindName("VisioProjectLineCombo")
    $channelCombo = $window.FindName("ChannelCombo")
    $channelProfileDefaultItem = $window.FindName("ChannelProfileDefaultItem")
    $langCombo = $window.FindName("LangCombo")
    $additionalLangList = $window.FindName("AdditionalLangList")
    $updatesEnabledCheck = $window.FindName("UpdatesEnabledCheck")
    $updatesTargetVersionBox = $window.FindName("UpdatesTargetVersionBox")
    $updatesDeadlineBox = $window.FindName("UpdatesDeadlineBox")
    $sharedComputerCustomPanel = $window.FindName("SharedComputerCustomPanel")
    $sharedComputerCustomCheck = $window.FindName("SharedComputerCustomCheck")
    $uiCombo = $window.FindName("UICombo")
    $installButton = $window.FindName("InstallButton")
    $statusPanel = $window.FindName("StatusPanel")
    $progressBar = $window.FindName("ProgressBar")
    $statusLabel = $window.FindName("StatusLabel")
    $excludeAppsPanel = $window.FindName("ExcludeAppsPanel")

    if ($null -eq $mainTabControl -or $null -eq $productSuiteCombo -or $null -eq $archCombo -or $null -eq $editionCombo -or $null -eq $installButton -or $null -eq $excludeAppsPanel) {
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
$script:statusPanel = $statusPanel
$script:statusLabel = $statusLabel
$script:progressBar = $progressBar
$script:excludeAppsPanel = $excludeAppsPanel

$cbStyle = $window.FindResource('SiteCheckBoxStyle')
foreach ($item in Get-M365AppsExcludeAppCatalog) {
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.Content = $item.Label
    $cb.Tag = $item.Id
    $cb.Margin = '0,4,16,4'
    if ($cbStyle) { $cb.Style = $cbStyle }
    [void]$excludeAppsPanel.Children.Add($cb)
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================
# These utility functions convert user selections to Office
# Deployment Tool (ODT) configuration values.

function Get-EditionID {
    param([int]$index)
    if ($index -eq 3) { return 'ADDONS_ONLY' }
    $editionMap = @{ 0 = "ProPlus2024Retail"; 1 = "ProPlus2021Volume"; 2 = "O365ProPlusRetail" }
    return $editionMap[$index]
}

function Get-EditionName {
    param([int]$index)
    if ($index -eq 3) { return 'Visio/Project only' }
    $nameMap = @{ 0 = "Office 2024 Pro Plus"; 1 = "Office LTSC 2021"; 2 = "Microsoft 365 Apps" }
    return $nameMap[$index]
}

function Get-M365RetailProfileFromPortalSelectors {
    $s = $productSuiteCombo.SelectedIndex
    if ($s -lt 0) { return $null }
    if ($s -eq 2) { return $null }
    $d = 0
    if ($null -ne $deploymentTargetCombo -and $deploymentTargetCombo.SelectedIndex -ge 0) {
        $d = $deploymentTargetCombo.SelectedIndex
    }
    if ($s -eq 0) {
        return $(if ($d -eq 0) { 'EnterprisePhysical' } else { 'EnterpriseVDI' })
    }
    if ($s -eq 1) {
        return $(if ($d -eq 0) { 'BusinessPhysical' } else { 'BusinessVDI' })
    }
    return $null
}

function Get-PortalDeploymentSummaryLabel {
    param(
        $RetailProfile,
        [bool]$IsCustom,
        [string]$EditionName
    )
    if ($RetailProfile) {
        switch ($RetailProfile) {
            'EnterprisePhysical' { return 'Microsoft 365 Apps for enterprise · this device' }
            'EnterpriseVDI' { return 'Microsoft 365 Apps for enterprise · shared computer / VDI' }
            'BusinessPhysical' { return 'Microsoft 365 Apps for business · this device' }
            'BusinessVDI' { return 'Microsoft 365 Apps for business · shared computer / VDI' }
        }
    }
    return "Other products: $EditionName"
}

function Get-SelectedExcludeAppIds {
    $ids = @()
    foreach ($c in $script:excludeAppsPanel.Children) {
        if ($c -is [System.Windows.Controls.CheckBox] -and $c.IsChecked -eq $true -and $c.Tag) {
            $ids += [string]$c.Tag
        }
    }
    return ,$ids
}

function Set-ExcludeAppsPanelEnabled {
    param([bool]$Enabled)
    foreach ($c in $script:excludeAppsPanel.Children) {
        if ($c -is [System.Windows.Controls.CheckBox]) { $c.IsEnabled = $Enabled }
    }
}

function Sync-LanguageComboFromProfile {
    try {
        $prevId = $null
        if ($langCombo.SelectedItem -and $null -ne $langCombo.SelectedItem.Tag) {
            $prevId = [string]$langCombo.SelectedItem.Tag
        }
        $incV = [bool]$visioCheck.IsChecked
        $incP = [bool]$projectCheck.IsChecked
        $langs = Get-M365AppsSupportedLanguages -IncludeVisio:$incV -IncludeProject:$incP
        $langCombo.Items.Clear()
        foreach ($lang in $langs) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $lang.Display
            $item.Tag = $lang.Id
            [void]$langCombo.Items.Add($item)
        }
        $pick = 'en-us'
        if ($prevId -and ($langs.Id -contains $prevId)) {
            $pick = $prevId
        }
        for ($i = 0; $i -lt $langCombo.Items.Count; $i++) {
            if ([string]$langCombo.Items[$i].Tag -eq $pick) {
                $langCombo.SelectedIndex = $i
                return
            }
        }
        if ($langCombo.Items.Count -gt 0) { $langCombo.SelectedIndex = 0 }
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to build language list.`n`n$_",
            "Office Auto Install",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit 1
    }
    Initialize-AdditionalLanguagesList
}

function Initialize-AdditionalLanguagesList {
    if ($null -eq $additionalLangList) { return }
    try {
        $additionalLangList.Items.Clear()
        $incV = [bool]$visioCheck.IsChecked
        $incP = [bool]$projectCheck.IsChecked
        foreach ($lang in Get-M365AppsSupportedLanguages -IncludeVisio:$incV -IncludeProject:$incP) {
            $item = New-Object System.Windows.Controls.ListBoxItem
            $item.Content = $lang.Display
            $item.Tag = $lang.Id
            [void]$additionalLangList.Items.Add($item)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to build additional language list.`n`n$_",
            "Office Auto Install",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Get-SelectedAdditionalLanguageIds {
    param([string]$PrimaryId)
    if ($null -eq $additionalLangList) { return @() }
    $primary = if ($PrimaryId) { $PrimaryId.ToLowerInvariant() } else { '' }
    $ids = New-Object System.Collections.Generic.List[string]
    foreach ($o in $additionalLangList.SelectedItems) {
        if ($null -eq $o) { continue }
        $id = [string]$o.Tag
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        $lc = $id.ToLowerInvariant()
        if ($lc -eq $primary) { continue }
        if (-not $ids.Contains($lc)) { [void]$ids.Add($lc) }
    }
    return ,$ids.ToArray()
}

function Sync-VisioProjectLineComboDefault {
    if ($null -eq $visioProjectLineCombo -or $null -eq $editionCombo) { return }
    $edIdx = $editionCombo.SelectedIndex
    $want = if ($edIdx -eq 3) {
        'M365Retail'
    } else {
        Get-M365AppsDefaultVisioProjectLine -ProductId (Get-EditionID -index $edIdx)
    }
    for ($i = 0; $i -lt $visioProjectLineCombo.Items.Count; $i++) {
        $it = $visioProjectLineCombo.Items[$i]
        if ([string]$it.Tag -eq $want) {
            $visioProjectLineCombo.SelectedIndex = $i
            return
        }
    }
}

function Update-ProfileDependentUI {
    $custom = ($productSuiteCombo.SelectedIndex -eq 2)
    $editionSection.Visibility = if ($custom) { 'Visible' } else { 'Collapsed' }
    $optionalSection.Visibility = 'Visible'
    $editionCombo.IsEnabled = $custom
    $visioCheck.IsEnabled = $true
    $projectCheck.IsEnabled = $true
    $addonsOnly = $custom -and ($editionCombo.SelectedIndex -eq 3)
    Set-ExcludeAppsPanelEnabled -Enabled (-not $addonsOnly)
    if ($null -ne $deploymentTargetPanel) {
        $deploymentTargetPanel.Visibility = if (-not $custom) { 'Visible' } else { 'Collapsed' }
    }
    if ($null -ne $sharedComputerCustomPanel) {
        $showSc = $custom -and ($editionCombo.SelectedIndex -eq 2)
        $sharedComputerCustomPanel.Visibility = if ($showSc) { 'Visible' } else { 'Collapsed' }
        if (-not $showSc -and $null -ne $sharedComputerCustomCheck) {
            $sharedComputerCustomCheck.IsChecked = $false
        }
    }
    if ($null -ne $visioProjectLineCombo) {
        $visioProjectLineCombo.IsEnabled = ($visioCheck.IsChecked -eq $true) -or ($projectCheck.IsChecked -eq $true)
    }
    if ($null -ne $visioProjectLinePanel) {
        $showVp = ($visioCheck.IsChecked -eq $true) -or ($projectCheck.IsChecked -eq $true)
        $visioProjectLinePanel.Visibility = if ($showVp) { 'Visible' } else { 'Collapsed' }
    }
    if ($channelProfileDefaultItem) {
        $channelProfileDefaultItem.IsEnabled = -not $custom
    }
    if ($custom -and $channelCombo.SelectedIndex -eq 0) {
        $channelCombo.SelectedIndex = 1
    }
    Sync-LanguageComboFromProfile
}

function Resolve-ChannelParameter {
    param(
        [bool]$IsCustomProfile,
        [int]$ChannelSelectedIndex
    )
    if ($IsCustomProfile) {
        if ($ChannelSelectedIndex -le 1) { return 'Current' }
        return 'SemiAnnualEnterprise'
    }
    switch ($ChannelSelectedIndex) {
        0 { return $null }
        1 { return 'Current' }
        2 { return 'SemiAnnualEnterprise' }
        default { return $null }
    }
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
    try {
        Test-M365AppsPrerequisites | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Prerequisites",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
    return $true
}

function Export-OdtConfigFromOptionsToPath {
    <#
    .SYNOPSIS
        Writes ODT configuration.xml from the same options used by the guided installer (retail profile or custom interactive).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DestinationPath,
        [Parameter(Mandatory)]
        [hashtable]$Options
    )
    $enc = New-Object System.Text.UTF8Encoding($false)
    if ($Options.retailProfile) {
        Assert-M365AppsLanguageCompatibleWithDeployment -LanguageId $Options.language -Preset '' `
            -CustomIncludeVisio:($Options.visio -eq '1') -CustomIncludeProject:($Options.project -eq '1')
        $xml = New-M365AppsO365ConfigurationForRetailProfile -RetailProfile $Options.retailProfile `
            -OfficeClientEdition $Options.bit -Channel $Options.channelOverride -LanguageId $Options.language `
            -AdditionalLanguageIds @($Options.additionalLanguageIds) -DisplayLevel $Options.ui `
            -AdditionalExcludeAppIds @($Options.excludeAppIds) -UpdatesEnabled:$Options.updatesEnabled `
            -UpdatesTargetVersion $Options.updatesTargetVersion -UpdatesDeadline $Options.updatesDeadline
        [System.IO.File]::WriteAllText($DestinationPath, $xml, $enc)
        Set-M365AppsConfigurationDisplayLevel -Path $DestinationPath -Level $Options.ui
        if ($Options.visio -eq '1' -or $Options.project -eq '1') {
            if (-not $Options.visioProjectLine) { throw 'Internal error: visioProjectLine missing for optional Visio/Project.' }
            Add-M365AppsOptionalVisioProjectProducts -Path $DestinationPath -LanguageId $Options.language `
                -IncludeVisio:($Options.visio -eq '1') -IncludeProject:($Options.project -eq '1') `
                -VisioProjectLine $Options.visioProjectLine -AdditionalLanguageIds @($Options.additionalLanguageIds)
        }
        return
    }
    Assert-M365AppsLanguageCompatibleWithDeployment -LanguageId $Options.language `
        -CustomIncludeVisio:($Options.visio -eq '1') -CustomIncludeProject:($Options.project -eq '1')
    $updEn = [bool]$Options.updatesEnabled
    $updTv = [string]$Options.updatesTargetVersion
    $updDl = [string]$Options.updatesDeadline
    $addLangs = @($Options.additionalLanguageIds)
    if ($Options.edition -eq 'ADDONS_ONLY') {
        $vpl = if ($Options.visioProjectLine) { $Options.visioProjectLine } else { 'M365Retail' }
        $xml = New-M365AppsInteractiveConfiguration -AddOnsOnly -LanguageId $Options.language `
            -OfficeClientEdition $Options.bit -Channel $Options.channel -DisplayLevel $Options.ui `
            -IncludeVisio:($Options.visio -eq '1') -IncludeProject:($Options.project -eq '1') -VisioProjectLine $vpl `
            -AdditionalLanguageIds $addLangs -UpdatesEnabled:$updEn -UpdatesTargetVersion $updTv -UpdatesDeadline $updDl
        [System.IO.File]::WriteAllText($DestinationPath, $xml, $enc)
        return
    }
    $sc = ($Options.sharedComputerCustom -eq $true)
    if ($Options.visioProjectLine) {
        $xml = New-M365AppsInteractiveConfiguration -ProductId $Options.edition -LanguageId $Options.language `
            -OfficeClientEdition $Options.bit -Channel $Options.channel -DisplayLevel $Options.ui `
            -IncludeVisio:($Options.visio -eq '1') -IncludeProject:($Options.project -eq '1') `
            -VisioProjectLine $Options.visioProjectLine -AutoActivate `
            -ExcludeAppIds @($Options.excludeAppIds) -AdditionalLanguageIds $addLangs `
            -UpdatesEnabled:$updEn -UpdatesTargetVersion $updTv -UpdatesDeadline $updDl `
            -SharedComputerLicensing:$sc
        [System.IO.File]::WriteAllText($DestinationPath, $xml, $enc)
        return
    }
    $xml = New-M365AppsInteractiveConfiguration -ProductId $Options.edition -LanguageId $Options.language `
        -OfficeClientEdition $Options.bit -Channel $Options.channel -DisplayLevel $Options.ui `
        -IncludeVisio:($Options.visio -eq '1') -IncludeProject:($Options.project -eq '1') -AutoActivate `
        -ExcludeAppIds @($Options.excludeAppIds) -AdditionalLanguageIds $addLangs `
        -UpdatesEnabled:$updEn -UpdatesTargetVersion $updTv -UpdatesDeadline $updDl `
        -SharedComputerLicensing:$sc
    [System.IO.File]::WriteAllText($DestinationPath, $xml, $enc)
}

function Build-UiInstallOptionsHashtable {
    <#
    .SYNOPSIS
        Collects guided-install options from the form. Returns $null if validation fails (message shown).
    #>
    $bit = if ($archCombo.SelectedIndex -eq 0) { '64' } else { '32' }
    $selLang = $langCombo.SelectedItem
    if (-not $selLang) {
        [System.Windows.Forms.MessageBox]::Show('Select a primary language on the Languages tab.', 'Office Auto Install',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $null
    }
    $languageName = [string]$selLang.Content
    $languageCode = if ($selLang.Tag) { [string]$selLang.Tag } else { Resolve-M365AppsLanguageId -Text $languageName }
    $uiLevel = if ($uiCombo.SelectedIndex -eq 0) { 'Full' } else { 'None' }
    $retailProfile = Get-M365RetailProfileFromPortalSelectors
    $isCustom = ($null -eq $retailProfile)
    $channelOverride = Resolve-ChannelParameter -IsCustomProfile $isCustom -ChannelSelectedIndex $channelCombo.SelectedIndex
    $excludeIds = @(Get-SelectedExcludeAppIds)
    $visio = if ($visioCheck.IsChecked) { '1' } else { '2' }
    $project = if ($projectCheck.IsChecked) { '1' } else { '2' }
    $vpl = $null
    if (($visio -eq '1' -or $project -eq '1') -and $null -ne $visioProjectLineCombo -and $visioProjectLineCombo.SelectedItem) {
        $vpl = [string]$visioProjectLineCombo.SelectedItem.Tag
    }
    $moreLangs = @(Get-SelectedAdditionalLanguageIds -PrimaryId $languageCode)
    $updEn = ($null -eq $updatesEnabledCheck) -or ($updatesEnabledCheck.IsChecked -eq $true)
    $updTv = if ($updatesTargetVersionBox) { [string]$updatesTargetVersionBox.Text.Trim() } else { '' }
    $updDl = if ($updatesDeadlineBox) { [string]$updatesDeadlineBox.Text.Trim() } else { '' }
    $sharedCustom = ($null -ne $sharedComputerCustomCheck -and $sharedComputerCustomCheck.IsChecked -eq $true)

    if ($isCustom) {
        $editionID = Get-EditionID -index $editionCombo.SelectedIndex
        $editionName = Get-EditionName -index $editionCombo.SelectedIndex
        if ($sharedCustom -and $editionID -ne 'O365ProPlusRetail') {
            [System.Windows.Forms.MessageBox]::Show(
                'Shared computer licensing applies only when the primary suite is Microsoft 365 Apps (Click-to-Run) on the custom path.',
                'Office Auto Install',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return $null
        }
        if ($editionID -eq 'ADDONS_ONLY' -and $visio -ne '1' -and $project -ne '1') {
            [System.Windows.Forms.MessageBox]::Show(
                'For "Visio and/or Project only", check at least one of the optional Visio or Project boxes.',
                'Office Auto Install',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return $null
        }
        if (($visio -eq '1' -or $project -eq '1') -and -not $vpl) {
            [System.Windows.Forms.MessageBox]::Show(
                'Select a Visio/Project product line, or uncheck both optional apps.',
                'Office Auto Install',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return $null
        }
        $channel = $channelOverride
        $profileLabel = Get-PortalDeploymentSummaryLabel -RetailProfile $null -IsCustom $true -EditionName $editionName
        $summaryLine = "$profileLabel | $languageName | ${bit}-bit"
        $ex = if ($editionID -eq 'ADDONS_ONLY') { @() } else { $excludeIds }
        return @{
            retailProfile = $null
            channelOverride = $null
            bit = $bit
            visio = $visio
            project = $project
            visioProjectLine = $vpl
            channel = $channel
            language = $languageCode
            languageName = $languageName
            ui = $uiLevel
            edition = $editionID
            editionName = $editionName
            profileLabel = $profileLabel
            summaryLine = $summaryLine
            excludeAppIds = $ex
            additionalLanguageIds = $moreLangs
            updatesEnabled = $updEn
            updatesTargetVersion = $updTv
            updatesDeadline = $updDl
            sharedComputerCustom = $sharedCustom
        }
    }
    if (($visio -eq '1' -or $project -eq '1') -and -not $vpl) {
        [System.Windows.Forms.MessageBox]::Show(
            'Select a Visio/Project product line, or uncheck both optional apps.',
            'Office Auto Install',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $null
    }
    $profileLabel = Get-PortalDeploymentSummaryLabel -RetailProfile $retailProfile -IsCustom $false -EditionName ''
    $summaryLine = "$profileLabel | $languageName | ${bit}-bit"
    return @{
        retailProfile = $retailProfile
        channelOverride = $channelOverride
        bit = $bit
        visio = $visio
        project = $project
        visioProjectLine = $vpl
        language = $languageCode
        languageName = $languageName
        ui = $uiLevel
        profileLabel = $profileLabel
        summaryLine = $summaryLine
        excludeAppIds = $excludeIds
        additionalLanguageIds = $moreLangs
        updatesEnabled = $updEn
        updatesTargetVersion = $updTv
        updatesDeadline = $updDl
        sharedComputerCustom = $false
    }
}

function Download-ODT {
    $output = Join-Path $installerFolder 'setup.exe'
    $odtUrl = 'https://officecdn.microsoft.com/pr/wsus/setup.exe'
    Update-Status "Downloading Office Deployment Tool..." 15
    Log "Downloading ODT from $odtUrl"
    try {
        $wc = New-Object System.Net.WebClient
        $eventJob = Register-ObjectEvent -InputObject $wc -EventName DownloadProgressChanged -Action {
            $pct = $Event.SourceEventArgs.ProgressPercentage
            $window.Dispatcher.Invoke([action]{
                $script:progressBar.Value = [Math]::Min(85, 15 + ($pct * 0.7))
                $script:statusLabel.Text = "Downloading... $pct%"
            })
        }
        $wc.DownloadFile($odtUrl, $output)
        $wc.Dispose()
        if ($eventJob) {
            Unregister-Event -SourceIdentifier $eventJob.Name -ErrorAction SilentlyContinue
            Remove-Job -Job $eventJob -ErrorAction SilentlyContinue
        }
        if (-not (Test-Path -LiteralPath $output) -or ((Get-Item -LiteralPath $output).Length -lt 100000)) {
            throw 'Download incomplete.'
        }
        Update-Status "Download completed." 85
        Log 'ODT saved.'
        Start-Sleep -Seconds 1
        return $true
    } catch {
        Log "Download failed: $_"
        [System.Windows.Forms.MessageBox]::Show("Download failed:`n$_", "Download Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

function Generate-Config {
    param($options)
    Update-Status "Creating configuration..." 87
    $configPath = Join-Path $installerFolder 'config.xml'
    $enc = New-Object System.Text.UTF8Encoding($false)
    if ($options.retailProfile) {
        Log "Generating config.xml for retail profile $($options.retailProfile) (baseline + user ExcludeApp merged)"
        Export-OdtConfigFromOptionsToPath -DestinationPath $configPath -Options $options
        Log "config.xml -> $configPath"
    } else {
        Log 'Generating config.xml (custom interactive)'
        Export-OdtConfigFromOptionsToPath -DestinationPath $configPath -Options $options
        Log "config.xml -> $configPath"
    }
    Update-Status "Configuration ready." 90
    Start-Sleep -Seconds 1
}

function Install-Office {
    param($options)
    Update-Status "Installing Office (may take a long time)..." 92
    Log 'Starting setup.exe /configure'
    $setupExe = Join-Path $installerFolder 'setup.exe'
    $configPath = Join-Path $installerFolder 'config.xml'
    if (-not (Test-Path -LiteralPath $setupExe)) {
        [System.Windows.Forms.MessageBox]::Show('setup.exe not found. Retry download.', 'Installation', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    try {
        Set-Location -LiteralPath $installerFolder
        $exitCode = Start-M365AppsSetup -SetupExePath $setupExe -ConfigurationPath $configPath -Wait
        if ($exitCode -eq 0) {
            Update-Status "Installation completed." 100
            Log "Exit code $exitCode"
            [System.Windows.Forms.MessageBox]::Show(
                "Installation finished (exit 0).`n`n$($options.summaryLine)",
                "Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return $true
        }
        Log "Exit code $exitCode"
        Update-Status "Finished with exit code $exitCode" 100
        [System.Windows.Forms.MessageBox]::Show(
            "Setup exit code: $exitCode. Verify apps in the Start menu.",
            "Setup",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return $true
    } catch {
        Log $_
        [System.Windows.Forms.MessageBox]::Show("$_", "Installation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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

function Restore-FullInstallUi {
    $installButton.IsEnabled = $true
    if ($null -ne $mainTabControl) { $mainTabControl.IsEnabled = $true }
    $productSuiteCombo.IsEnabled = $true
    if ($null -ne $deploymentTargetCombo) { $deploymentTargetCombo.IsEnabled = $true }
    $archCombo.IsEnabled = $true
    Update-ProfileDependentUI
    $channelCombo.IsEnabled = $true
    $langCombo.IsEnabled = $true
    if ($null -ne $additionalLangList) { $additionalLangList.IsEnabled = $true }
    if ($null -ne $updatesEnabledCheck) { $updatesEnabledCheck.IsEnabled = $true }
    if ($null -ne $updatesTargetVersionBox) { $updatesTargetVersionBox.IsEnabled = $true }
    if ($null -ne $updatesDeadlineBox) { $updatesDeadlineBox.IsEnabled = $true }
    $uiCombo.IsEnabled = $true
    if ($null -ne $sharedComputerCustomCheck) { $sharedComputerCustomCheck.IsEnabled = $true }
    Set-ExcludeAppsPanelEnabled -Enabled $true
}

$productSuiteCombo.Add_SelectionChanged({ Update-ProfileDependentUI })
if ($null -ne $deploymentTargetCombo) {
    $deploymentTargetCombo.Add_SelectionChanged({ Update-ProfileDependentUI })
}
$editionCombo.Add_SelectionChanged({ Sync-VisioProjectLineComboDefault; Update-ProfileDependentUI })
$visioCheck.Add_Checked({ Update-ProfileDependentUI })
$visioCheck.Add_Unchecked({ Update-ProfileDependentUI })
$projectCheck.Add_Checked({ Update-ProfileDependentUI })
$projectCheck.Add_Unchecked({ Update-ProfileDependentUI })
Sync-VisioProjectLineComboDefault
Update-ProfileDependentUI

$installButton.Add_Click({
    $installButton.IsEnabled = $false
    if ($null -ne $mainTabControl) { $mainTabControl.IsEnabled = $false }
    $productSuiteCombo.IsEnabled = $false
    if ($null -ne $deploymentTargetCombo) { $deploymentTargetCombo.IsEnabled = $false }
    $archCombo.IsEnabled = $false
    $editionCombo.IsEnabled = $false
    $visioCheck.IsEnabled = $false
    $projectCheck.IsEnabled = $false
    if ($null -ne $visioProjectLineCombo) { $visioProjectLineCombo.IsEnabled = $false }
    $channelCombo.IsEnabled = $false
    $langCombo.IsEnabled = $false
    if ($null -ne $additionalLangList) { $additionalLangList.IsEnabled = $false }
    if ($null -ne $updatesEnabledCheck) { $updatesEnabledCheck.IsEnabled = $false }
    if ($null -ne $updatesTargetVersionBox) { $updatesTargetVersionBox.IsEnabled = $false }
    if ($null -ne $updatesDeadlineBox) { $updatesDeadlineBox.IsEnabled = $false }
    $uiCombo.IsEnabled = $false
    if ($null -ne $sharedComputerCustomCheck) { $sharedComputerCustomCheck.IsEnabled = $false }
    Set-ExcludeAppsPanelEnabled -Enabled $false

    try {
        Log "=== Office Installer GUI Started ==="
        
        if (-not (Test-SystemRequirements)) {
            Restore-FullInstallUi
            $statusPanel.Visibility = "Collapsed"
            return
        }
        
        $options = Build-UiInstallOptionsHashtable
        if ($null -eq $options) {
            Restore-FullInstallUi
            $statusPanel.Visibility = "Collapsed"
            return
        }
        
        if (-not (Download-ODT)) {
            Restore-FullInstallUi
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
        Restore-FullInstallUi
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

