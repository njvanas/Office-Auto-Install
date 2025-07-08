# Office Auto Installer

![Screenshot](./screenshot.png)

This enhanced PowerShell script provides a beautiful, beginner-friendly interface for downloading, configuring, and installing Microsoft Office 2024/2021/365 through official Microsoft channels. Designed for any Windows user - no technical knowledge required!

## ✨ Features
### 🎯 **Beginner-Friendly Design**
- **Zero Technical Knowledge Required** - Simple, guided setup for everyone
- **Automatic Administrator Elevation** - Handles Windows permissions automatically
- **Smart Recommendations** - Suggests best options for typical users
- **Plain English Explanations** - No confusing technical jargon
- **Helpful Hints** - Guidance for every decision you need to make

### 🖥️ **Windows Integration**
- **Automatic Admin Rights** - Requests elevation when needed with clear explanation
- **System Requirements Check** - Validates disk space, RAM, Windows version, and internet
- **UAC Handling** - Smooth integration with Windows security prompts
- **Execution Policy Bypass** - Includes batch file for restricted systems
- **Professional Error Messages** - Clear explanations when something goes wrong

### 🎨 **Beautiful Interface**
- **Colorful Menus** - Easy-to-read, visually appealing interface
- **Progress Tracking** - Real-time feedback during download and installation
- **Step-by-Step Guidance** - Numbered steps with clear explanations
- **Smart Defaults** - Pre-selected recommended options
- **Comprehensive Summaries** - Review your choices before installation

### ⚙️ **Flexible Configuration**
- **Office Editions:** 2024 Pro Plus, LTSC 2021, Microsoft 365 Apps
- **Architectures:** 32-bit and 64-bit (with recommendations)
- **Languages:** 7 supported languages with clear descriptions
- **Update Channels:** Monthly or Semi-Annual with explanations
- **Optional Components:** Visio and Project with usage explanations
- **Installation Modes:** Visual or silent installation

## 🚀 Getting Started

### 📥 **Download & Run**
1. **Download** both files to the same folder:
   - `Install-Office.ps1` (main script)
   - `Install-Office(RunMeIfPowershellFails).bat` (backup launcher)

2. **Run the installer** (choose one method):
   - **Method 1:** Right-click `Install-Office.ps1` → "Run with PowerShell"
   - **Method 2:** Double-click `Install-Office(RunMeIfPowershellFails).bat`
   - **Method 3:** Open PowerShell as admin and run the script

3. **Follow the prompts:**
   - Click "Yes" when Windows asks for administrator permission
   - Follow the step-by-step guided setup
   - The script will explain each option in simple terms

4. **Wait for completion:**
   - Installation takes 10-30 minutes depending on internet speed
   - Don't close the window during installation
   - Perfect time for a coffee break! ☕

### 🆘 **If You Have Problems**
- **Script won't run?** Use the `.bat` file instead
- **Need admin rights?** The script will request them automatically
- **Antivirus blocking?** Temporarily disable it during installation
- **Still stuck?** Check the generated log file for details

## 🎯 **Perfect For**
- **Home Users** - Installing Office on personal computers
- **Small Businesses** - Setting up Office on multiple computers
- **IT Support** - Quick, reliable Office deployment
- **Students** - Easy Office installation for school work
- **Anyone** - Who wants Office installed without the hassle!

## 🔒 **Safety & Security**
- **100% Official Microsoft Tools** - Uses only Microsoft's official deployment tools
- **No Modifications** - Doesn't crack, patch, or modify Office in any way
- **Safe Downloads** - Downloads directly from Microsoft's servers
- **Transparent Process** - Full logging of all actions taken
- **Respects Licensing** - You must have proper Office licenses

## ⚠️ Disclaimer
This script downloads and installs Microsoft software through official Microsoft deployment tools. **You are responsible for having proper Microsoft Office licenses.** This tool does not crack, modify, or bypass any licensing - it only makes installation easier. **Use at your own risk.** The author assumes no responsibility for licensing compliance or any issues that may arise.

## 🛠️ **Technical Details**
- **Official Tools Only** - Uses Microsoft's Office Deployment Tool (ODT)
- **Direct Downloads** - Gets files straight from Microsoft's servers
- **Standard Configuration** - Creates proper XML config files
- **Full Logging** - Detailed logs for troubleshooting
- **Windows Compatible** - Works on Windows 10/11 (Windows 7/8 may work)
- **System Requirements** - 4GB+ disk space, 2GB+ RAM recommended

## 📄 License
This project is licensed under the [MIT License](./LICENSE).