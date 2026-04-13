#!/usr/bin/env bash
# Builds ZIPs for GitHub Releases. Called from release workflow with version tag as $1.
set -euo pipefail

VERSION="${1:?Usage: build-release-assets.sh <version e.g. v1.2.3>}"
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$ROOT"

echo "Building release assets for $VERSION"

rm -f ./*.zip
rm -rf release-package-console release-package-gui release-package-complete
mkdir -p release-package-console release-package-gui release-package-complete

if [[ -f Install-Office.ps1 ]]; then
  cp Install-Office.ps1 release-package-console/
fi

if [[ -d release-package-console && -n "$(ls -A release-package-console)" ]]; then
  cat >release-package-console/README.txt <<EOF
Office Auto Installer - Console Version Package - $VERSION
================================================================

QUICK START INSTRUCTIONS:
1. Save Install-Office.ps1 to a folder (e.g. Downloads)
2. Open PowerShell as Administrator
3. Run: cd "path\\to\\folder"
4. Run: powershell -NoProfile -ExecutionPolicy Bypass -File ".\\Install-Office.ps1"

WEB LAUNCH (admin PowerShell, recommended):
\$env:OFFICE_AUTO_INSTALL_USE_CONSOLE = "1"
irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex

WHAT'S INCLUDED:
- Install-Office.ps1 (Console installer script)
- README.txt (This file with instructions)

SYSTEM REQUIREMENTS:
- Windows 10/11 (Windows 7/8 may work)
- 4GB+ free disk space
- 2GB+ RAM recommended
- Internet connection (Office downloads during install)
- Administrator privileges

SUPPORTED OFFICE VERSIONS:
- Office 2024 Pro Plus (Latest - Recommended)
- Office LTSC 2021 (Stable long-term version)
- Microsoft 365 Apps (Cloud-connected)

SUPPORTED LANGUAGES:
- English (US/UK), French, German, Dutch, Spanish, Portuguese

OPTIONAL COMPONENTS:
- Visio Professional (Diagrams and flowcharts)
- Project Professional (Project management)

TROUBLESHOOTING:
If you have problems:
1. Run PowerShell as Administrator
2. Temporarily disable antivirus during installation if it blocks scripts
3. Check your internet connection

For more help, visit: https://njvanas.github.io/Office-Auto-Install/

IMPORTANT DISCLAIMER:
This tool uses Microsoft's official Office Deployment Tool.
You are responsible for having proper Office licenses.
This does not crack, modify, or bypass any licensing.

Licensed under MIT License
Copyright (c) 2025 Dolfie
EOF
  (cd release-package-console && zip -r "../Office-Auto-Installer-Console-${VERSION}.zip" .)
fi

if [[ -f Install-Office-GUI-WPF.ps1 ]]; then
  cp Install-Office-GUI-WPF.ps1 release-package-gui/
fi

if [[ -d release-package-gui && -n "$(ls -A release-package-gui)" ]]; then
  cat >release-package-gui/README.txt <<EOF
Office Auto Installer - GUI Version Package - $VERSION
================================================================

QUICK START INSTRUCTIONS:
1. Save Install-Office-GUI-WPF.ps1 to a folder (e.g. Downloads)
2. Open PowerShell as Administrator
3. Run: cd "path\\to\\folder"
4. Run: powershell -NoProfile -ExecutionPolicy Bypass -File ".\\Install-Office-GUI-WPF.ps1"

WEB LAUNCH (admin PowerShell, recommended):
irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex

WHAT'S INCLUDED:
- Install-Office-GUI-WPF.ps1 (GUI installer script)
- README.txt (This file with instructions)

FEATURES:
- GUI styled to match the project website (slate / blue)
- Pre-filled defaults
- Real-time progress tracking
- One-click ready

SYSTEM REQUIREMENTS:
- Windows 10/11 (Windows 7/8 may work)
- 4GB+ free disk space
- 2GB+ RAM recommended
- Internet connection (Office downloads during install)
- Administrator privileges

SUPPORTED OFFICE VERSIONS:
- Office 2024 Pro Plus (Latest - Recommended)
- Office LTSC 2021 (Stable long-term version)
- Microsoft 365 Apps (Cloud-connected)

SUPPORTED LANGUAGES:
- English (US/UK), French, German, Dutch, Spanish, Portuguese

OPTIONAL COMPONENTS:
- Visio Professional (Diagrams and flowcharts)
- Project Professional (Project management)

TROUBLESHOOTING:
If you have problems:
1. Run PowerShell as Administrator
2. Temporarily disable antivirus during installation if it blocks scripts
3. Check your internet connection

For more help, visit: https://njvanas.github.io/Office-Auto-Install/

IMPORTANT DISCLAIMER:
This tool uses Microsoft's official Office Deployment Tool.
You are responsible for having proper Office licenses.
This does not crack, modify, or bypass any licensing.

Licensed under MIT License
Copyright (c) 2025 Dolfie
EOF
  (cd release-package-gui && zip -r "../Office-Auto-Installer-GUI-${VERSION}.zip" .)
fi

for f in office.ps1 Install-Office.ps1 Install-Office-GUI-WPF.ps1 README.md LICENSE NOTICES.md M365AppsCore.ps1 Deploy-Microsoft365Apps.ps1; do
  if [[ -f $f ]]; then
    cp "$f" release-package-complete/
  fi
done
if [[ -d configs ]]; then
  cp -r configs release-package-complete/
fi

if [[ -d release-package-complete && -n "$(ls -A release-package-complete)" ]]; then
  cat >release-package-complete/README.txt <<EOF
Office Auto Installer - Complete Package - $VERSION
================================================================

Full payload: bootstrap, both installers, shared modules, sample configs, and docs.

QUICK START:
Web launch (admin PowerShell):
irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex

Or run a script from disk (admin PowerShell):
powershell -NoProfile -ExecutionPolicy Bypass -File ".\\Install-Office-GUI-WPF.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\\Install-Office.ps1"

WHAT'S INCLUDED:
- office.ps1 (Web bootstrap)
- Install-Office-GUI-WPF.ps1, Install-Office.ps1
- M365AppsCore.ps1, Deploy-Microsoft365Apps.ps1 (when present)
- configs/ (sample or reference configuration files, when present)
- README.md, LICENSE, NOTICES.md (when present)

For more help, visit: https://njvanas.github.io/Office-Auto-Install/

Licensed under MIT License
Copyright (c) 2025 Dolfie
EOF
  (cd release-package-complete && zip -r "../Office-Auto-Installer-Complete-${VERSION}.zip" .)
fi

ls -la ./*.zip
