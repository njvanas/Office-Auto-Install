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
    Version: 3.7 - Deployment profiles (configs\ presets) + custom advanced path
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
        Width="920" Height="1050"
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
          <TextBlock Text="Automated Office deployment with customizable options"
                     FontSize="12"
                     Foreground="{StaticResource SiteTextMutedBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,4,0,0"
                     TextWrapping="Wrap"/>
        </StackPanel>
      </StackPanel>
    </Border>

    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Background="Transparent">
      <StackPanel Margin="40,28,40,20">

        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="Deployment profile"
                     FontSize="15" FontWeight="SemiBold"
                     Foreground="{StaticResource SiteTextBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,0,0,12"/>
          <TextBlock Text="Choose the scenario that matches your device (PC or VDI). Presets match the XML in configs\."
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
            <ComboBox x:Name="ProfileCombo" Style="{StaticResource SiteComboBoxStyle}">
              <ComboBoxItem Content="Microsoft 365 Apps — enterprise (physical / desktop)" IsSelected="True"/>
              <ComboBoxItem Content="Microsoft 365 Apps — enterprise (VDI / shared PC)"/>
              <ComboBoxItem Content="Microsoft 365 Apps — business (physical / desktop)"/>
              <ComboBoxItem Content="Microsoft 365 Apps — business (VDI / shared PC)"/>
              <ComboBoxItem Content="M365 Apps enterprise + Visio &amp; Project (physical / desktop)"/>
              <ComboBoxItem Content="M365 Apps enterprise + Visio &amp; Project (VDI / shared PC)"/>
              <ComboBoxItem Content="Custom — Office 2024, LTSC 2021, or build-your-own M365"/>
            </ComboBox>
          </Border>
        </StackPanel>

        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="System Configuration"
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
              <ComboBoxItem Content="64-bit (Recommended for most computers)" IsSelected="True"/>
              <ComboBoxItem Content="32-bit (For older systems)"/>
            </ComboBox>
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
            </ComboBox>
          </Border>
        </StackPanel>

        <StackPanel x:Name="OptionalSection" Margin="0,0,0,24" Visibility="Collapsed">
          <TextBlock Text="Optional Components (custom only)"
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
              <CheckBox x:Name="VisioCheck" Style="{StaticResource SiteCheckBoxStyle}"
                        Content="Include Visio Professional (for diagrams and flowcharts)"
                        Margin="0,0,0,12"/>
              <CheckBox x:Name="ProjectCheck" Style="{StaticResource SiteCheckBoxStyle}"
                        Content="Include Project Professional (for project management)"/>
            </StackPanel>
          </Border>
        </StackPanel>

        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="Update channel"
                     FontSize="15" FontWeight="SemiBold"
                     Foreground="{StaticResource SiteTextBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,0,0,12"/>
          <TextBlock Text="Deployment profiles use the channel baked into the preset XML unless you override it below."
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
              <ComboBoxItem x:Name="ChannelPresetDefaultItem" Content="Use preset default (from XML)" IsSelected="True"/>
              <ComboBoxItem Content="Monthly / Current (override)"/>
              <ComboBoxItem Content="Semi-annual Enterprise (override)"/>
            </ComboBox>
          </Border>
        </StackPanel>

        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="Language"
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
              <TextBlock Text="Languages are filtered by profile: Visio/Project bundles exclude combinations Microsoft does not support in one install (e.g. en-gb). Then click Install Office."
                         FontSize="12"
                         Foreground="{StaticResource SiteTextMutedBrush}"
                         FontFamily="Inter, Segoe UI"
                         TextWrapping="Wrap"/>
            </StackPanel>
          </Border>
        </StackPanel>

        <StackPanel Margin="0,0,0,24">
          <TextBlock Text="Installation Display"
                     FontSize="15" FontWeight="SemiBold"
                     Foreground="{StaticResource SiteTextBrush}"
                     FontFamily="Inter, Segoe UI"
                     Margin="0,0,0,12"/>
          <Border Background="{StaticResource SiteCardBrush}"
                  BorderBrush="{StaticResource SiteBorderBrush}"
                  BorderThickness="1"
                  CornerRadius="12"
                  Padding="20,16">
            <ComboBox x:Name="UICombo" Style="{StaticResource SiteComboBoxStyle}">
              <ComboBoxItem Content="Show installation progress (Recommended)" IsSelected="True"/>
              <ComboBoxItem Content="Install quietly in background"/>
            </ComboBox>
          </Border>
        </StackPanel>

      </StackPanel>
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
# - ProfileCombo: Preset deployment profile (configs\ XML) or Custom advanced build
# - ArchCombo: System architecture selection (32-bit/64-bit)
# - EditionCombo: Office edition selection (custom path only)
# - VisioCheck / ProjectCheck: Optional components (custom path only)
# - ChannelCombo: Preset default vs Current / Semi-annual override
# - LangCombo: Language selection for Office installation
# - UICombo: Installation display mode (Show progress / Quiet)
# - InstallButton: Primary action button to start installation
# - StatusPanel: Container for progress indicators (hidden by default)
# - ProgressBar: Visual progress indicator during download/install
# - StatusLabel: Text status updates during installation

try {
    $profileCombo = $window.FindName("ProfileCombo")
    $editionSection = $window.FindName("EditionSection")
    $optionalSection = $window.FindName("OptionalSection")
    $archCombo = $window.FindName("ArchCombo")
    $editionCombo = $window.FindName("EditionCombo")
    $visioCheck = $window.FindName("VisioCheck")
    $projectCheck = $window.FindName("ProjectCheck")
    $channelCombo = $window.FindName("ChannelCombo")
    $channelPresetDefaultItem = $window.FindName("ChannelPresetDefaultItem")
    $langCombo = $window.FindName("LangCombo")
    $uiCombo = $window.FindName("UICombo")
    $installButton = $window.FindName("InstallButton")
    $statusPanel = $window.FindName("StatusPanel")
    $progressBar = $window.FindName("ProgressBar")
    $statusLabel = $window.FindName("StatusLabel")
    
    if ($null -eq $profileCombo -or $null -eq $archCombo -or $null -eq $editionCombo -or $null -eq $installButton) {
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

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================
# These utility functions convert user selections to Office
# Deployment Tool (ODT) configuration values.

function Get-EditionID {
    param([int]$index)
    $editionMap = @{ 0 = "ProPlus2024Retail"; 1 = "ProPlus2021Volume"; 2 = "O365ProPlusRetail" }
    return $editionMap[$index]
}

function Get-EditionName {
    param([int]$index)
    $nameMap = @{ 0 = "Office 2024 Pro Plus"; 1 = "Office LTSC 2021"; 2 = "Microsoft 365 Apps" }
    return $nameMap[$index]
}

function Get-PresetNameFromProfileIndex {
    param([int]$index)
    $map = @{
        0 = 'O365ProPlus'
        1 = 'O365ProPlus-VDI'
        2 = 'O365Business'
        3 = 'O365Business-VDI'
        4 = 'O365ProPlusVisioProject'
        5 = 'O365ProPlusVisioProject-VDI'
    }
    if ($map.ContainsKey($index)) { return $map[$index] }
    return $null
}

function Get-ProfileSummaryLabel {
    param([int]$index)
    $labels = @(
        'M365 Apps enterprise (physical)'
        'M365 Apps enterprise (VDI)'
        'M365 Apps business (physical)'
        'M365 Apps business (VDI)'
        'M365 enterprise + Visio & Project (physical)'
        'M365 enterprise + Visio & Project (VDI)'
        'Custom (interactive XML)'
    )
    if ($index -ge 0 -and $index -lt $labels.Count) { return $labels[$index] }
    return 'Custom'
}

function Sync-LanguageComboFromProfile {
    try {
        $prevId = $null
        if ($langCombo.SelectedItem -and $null -ne $langCombo.SelectedItem.Tag) {
            $prevId = [string]$langCombo.SelectedItem.Tag
        }
        $preset = Get-PresetNameFromProfileIndex -index $profileCombo.SelectedIndex
        $incV = $false
        $incP = $false
        if ($profileCombo.SelectedIndex -eq 6) {
            $incV = [bool]$visioCheck.IsChecked
            $incP = [bool]$projectCheck.IsChecked
        }
        $langs = if ($null -eq $preset) {
            Get-M365AppsSupportedLanguages -IncludeVisio:$incV -IncludeProject:$incP
        } else {
            Get-M365AppsSupportedLanguages -Preset $preset
        }
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
}

function Update-ProfileDependentUI {
    $custom = ($profileCombo.SelectedIndex -eq 6)
    $editionSection.Visibility = if ($custom) { 'Visible' } else { 'Collapsed' }
    $optionalSection.Visibility = if ($custom) { 'Visible' } else { 'Collapsed' }
    $editionCombo.IsEnabled = $custom
    $visioCheck.IsEnabled = $custom
    $projectCheck.IsEnabled = $custom
    if ($channelPresetDefaultItem) {
        $channelPresetDefaultItem.IsEnabled = -not $custom
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
    if ($options.presetName) {
        Assert-M365AppsLanguageCompatibleWithDeployment -LanguageId $options.language -Preset $options.presetName
        Log "Generating config.xml from preset $($options.presetName)"
        $src = Get-M365AppsPresetConfigurationPath -Preset $options.presetName
        $ch = $options.channelOverride
        Copy-M365AppsConfigurationWithOverrides -SourcePath $src -DestinationPath $configPath `
            -OfficeClientEdition $options.bit -LanguageId $options.language -Channel $ch
        Set-M365AppsConfigurationDisplayLevel -Path $configPath -Level $options.ui
    } else {
        Assert-M365AppsLanguageCompatibleWithDeployment -LanguageId $options.language `
            -CustomIncludeVisio:($options.visio -eq '1') -CustomIncludeProject:($options.project -eq '1')
        Log 'Generating config.xml (custom interactive)'
        $xml = New-M365AppsInteractiveConfiguration -ProductId $options.edition -LanguageId $options.language `
            -OfficeClientEdition $options.bit -Channel $options.channel -DisplayLevel $options.ui `
            -IncludeVisio:($options.visio -eq '1') -IncludeProject:($options.project -eq '1') -AutoActivate
        $enc = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($configPath, $xml, $enc)
    }
    Log "config.xml -> $configPath"
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

$profileCombo.Add_SelectionChanged({ Update-ProfileDependentUI })
$visioCheck.Add_Checked({ Update-ProfileDependentUI })
$visioCheck.Add_Unchecked({ Update-ProfileDependentUI })
$projectCheck.Add_Checked({ Update-ProfileDependentUI })
$projectCheck.Add_Unchecked({ Update-ProfileDependentUI })
Update-ProfileDependentUI

$installButton.Add_Click({
    $installButton.IsEnabled = $false
    $profileCombo.IsEnabled = $false
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
            $profileCombo.IsEnabled = $true
            $archCombo.IsEnabled = $true
            Update-ProfileDependentUI
            $channelCombo.IsEnabled = $true
            $langCombo.IsEnabled = $true
            $uiCombo.IsEnabled = $true
            $statusPanel.Visibility = "Collapsed"
            return
        }
        
        $bit = if ($archCombo.SelectedIndex -eq 0) { "64" } else { "32" }
        $selLang = $langCombo.SelectedItem
        $languageName = [string]$selLang.Content
        $languageCode = if ($selLang.Tag) { [string]$selLang.Tag } else { Resolve-M365AppsLanguageId -Text $languageName }
        $uiLevel = if ($uiCombo.SelectedIndex -eq 0) { "Full" } else { "None" }
        $presetName = Get-PresetNameFromProfileIndex -index $profileCombo.SelectedIndex
        $isCustom = ($null -eq $presetName)
        
        $channelOverride = Resolve-ChannelParameter -IsCustomProfile $isCustom -ChannelSelectedIndex $channelCombo.SelectedIndex
        $profileLabel = Get-ProfileSummaryLabel -index $profileCombo.SelectedIndex
        
        if ($isCustom) {
            $editionID = Get-EditionID -index $editionCombo.SelectedIndex
            $editionName = Get-EditionName -index $editionCombo.SelectedIndex
            $visio = if ($visioCheck.IsChecked) { "1" } else { "2" }
            $project = if ($projectCheck.IsChecked) { "1" } else { "2" }
            $channel = $channelOverride
            $summaryLine = "$editionName | $languageName | ${bit}-bit | custom"
            $options = @{
                presetName = $null
                channelOverride = $null
                bit = $bit
                visio = $visio
                project = $project
                channel = $channel
                language = $languageCode
                languageName = $languageName
                ui = $uiLevel
                edition = $editionID
                editionName = $editionName
                profileLabel = $profileLabel
                summaryLine = $summaryLine
            }
        } else {
            $summaryLine = "$profileLabel | $languageName | ${bit}-bit | preset $($presetName).xml"
            $options = @{
                presetName = $presetName
                channelOverride = $channelOverride
                bit = $bit
                language = $languageCode
                languageName = $languageName
                ui = $uiLevel
                profileLabel = $profileLabel
                summaryLine = $summaryLine
            }
        }
        
        if (-not (Download-ODT)) {
            $installButton.IsEnabled = $true
            $profileCombo.IsEnabled = $true
            $archCombo.IsEnabled = $true
            Update-ProfileDependentUI
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
        $profileCombo.IsEnabled = $true
        $archCombo.IsEnabled = $true
        Update-ProfileDependentUI
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

