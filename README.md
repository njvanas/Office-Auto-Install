# Office Auto Installer

![Screenshot](./screenshot.png)

This enhanced PowerShell script provides a beautiful, user-friendly interface for downloading, configuring, and installing Microsoft Office 2024/2021/365 through official Microsoft channels, with optional components like Visio and Project.

## ✨ Features
- **Enhanced Visual Interface** - Beautiful, colorful menus and progress indicators
- **Comprehensive System Checks** - Validates internet connectivity and system PATH
- **Interactive Configuration** - Step-by-step guided setup with clear options
- **Real-time Progress** - Visual feedback during download and installation
- **Detailed Summaries** - Clear configuration review and completion status
- Choose 32-bit or 64-bit architecture
- Select from 7 supported languages (en-us, en-gb, fr-fr, de-de, nl-nl, es-es, pt-br)
- Choose update channel (Monthly or Semi-Annual)
- Optional install of Visio and/or Project
- Prompt for Silent or Full UI installation
- Automatic admin elevation
- Automatic system PATH validation and fixes
- Full logging for installation steps and errors
- Professional error handling with helpful messages

## 🚀 Getting Started
1. Clone or download the repo
2. Right-click the `.ps1` script and **Run with PowerShell** (it will prompt for elevation if needed)
3. If the script closes unexpectedly, it's likely due to PowerShell execution policy. Use the included batch file to run with policy bypass.
4. Follow the enhanced interactive prompts - the interface will guide you through each step
5. **Important:** Do not close the window during installation as this can corrupt the process
6. Upon completion, review the summary and check the generated logs

## 🎨 Interface Features
- **Colorful Headers** - Professional branding and clear section identification  
- **Progress Indicators** - Real-time feedback during operations
- **Smart Defaults** - Recommended options pre-selected for ease of use
- **Input Validation** - Prevents invalid selections with helpful error messages
- **Comprehensive Summaries** - Review configuration before installation and see detailed completion status

## 📋 Supported Configurations
- **Office Editions:** 2024 Pro Plus, LTSC 2021, Microsoft 365 Apps
- **Architectures:** 32-bit and 64-bit
- **Languages:** English (US/UK), French, German, Dutch, Spanish, Portuguese (Brazil)
- **Update Channels:** Monthly (Current) and Semi-Annual (Broad)
- **Additional Components:** Visio Professional, Project Professional
- **Installation Modes:** Full UI or Silent installation

## ⚠️ Disclaimer
This script downloads and installs Microsoft software through official Microsoft deployment tools. **You are solely responsible for ensuring your use complies with Microsoft licensing terms.** This tool does not modify, crack, or bypass any licensing mechanisms - it only facilitates the download and installation of official Microsoft Office packages. **Use at your own risk.** The author assumes no responsibility for any misuse, licensing violations, or damage.

## 🔧 Technical Details
- Uses Microsoft's official Office Deployment Tool (ODT)
- Downloads directly from Microsoft's CDN servers
- Generates standard XML configuration files
- Maintains full compatibility with Microsoft's licensing system
- Includes comprehensive logging for troubleshooting

## 📄 License
This project is licensed under the [MIT License](./LICENSE).