# Office Auto Installer

![Screenshot](./screenshot.png)

This enhanced PowerShell script provides a beautiful, beginner-friendly interface for downloading, configuring, and installing Microsoft Office 2024/2021/365 through official Microsoft channels. **Now with PowerShell execution policy fixes** - designed for any Windows user with no technical knowledge required!

## ✨ Features
### 🎯 **Beginner-Friendly Design**
- **Zero Technical Knowledge Required** - Simple, guided setup for everyone
- **Automatic Administrator Elevation** - Handles Windows permissions automatically
- **Smart Recommendations** - Suggests best options for typical users
- **Plain English Explanations** - No confusing technical jargon
- **Helpful Hints** - Guidance for every decision you need to make

### 🖥️ **Windows Integration**
- **Execution Policy Fix** - Automatically handles PowerShell execution policy restrictions
- **Copy & Paste Ready** - Works when copied and pasted into PowerShell
- **Right-Click Protection** - Prevents window from closing immediately
- **Automatic Admin Rights** - Requests elevation when needed with clear explanation
- **System Requirements Check** - Validates disk space, RAM, Windows version, and internet
- **UAC Handling** - Smooth integration with Windows security prompts
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
1. **Download** the PowerShell script:
   - `Install-Office.ps1`

2. **Run the installer** (try these methods in order):

   **🥇 METHOD 1 - Bypass Execution Policy (RECOMMENDED):**
   1. Right-click on PowerShell → "Run as Administrator"
   2. Navigate to the script location: `cd "C:\Users\YourName\Downloads"`
   3. Run: `powershell -ExecutionPolicy Bypass -File "Install-Office.ps1"`
   4. Click "Yes" when Windows asks for permission

   **🥈 METHOD 2 - Copy & Paste Method:**
   1. Right-click on PowerShell → "Run as Administrator"
   2. Open the script file in Notepad
   3. Select all content (Ctrl+A) and copy (Ctrl+C)
   4. Paste into PowerShell window (Right-click → Paste)
   5. Press Enter to run

   **🥉 METHOD 3 - Change Execution Policy:**
   1. Right-click on PowerShell → "Run as Administrator"
   2. Run: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force`
   3. Run: `.\Install-Office.ps1`

3. **Follow the prompts:**
   - Click "Yes" when Windows asks for administrator permission
   - Follow the step-by-step guided setup
   - The script will explain each option in simple terms

4. **Wait for completion:**
   - Installation takes 10-30 minutes depending on internet speed
   - Don't close the window during installation
   - Perfect time for a coffee break! ☕

### 🆘 **If You Have Problems**
- **Script won't run?** Try Method 1 above with `-ExecutionPolicy Bypass`
- **"Cannot be loaded" error?** This is Windows security - use Method 1 or 2 above
- **Right-click doesn't work?** Use Method 1 or 2 instead - they're more reliable
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
- **Execution Policy Safe** - Automatically handles PowerShell restrictions
- **Direct Downloads** - Gets files straight from Microsoft's servers
- **Standard Configuration** - Creates proper XML config files
- **Full Logging** - Detailed logs for troubleshooting
- **Windows Compatible** - Works on Windows 10/11 (Windows 7/8 may work)
- **System Requirements** - 4GB+ disk space, 2GB+ RAM recommended

## 📄 License
This project is licensed under the [MIT License](./LICENSE).