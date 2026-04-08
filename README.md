# Office Auto Install

**Office-Auto-Install** is a PowerShell automation layer for deploying **Microsoft 365 Apps** using **Microsoft’s supported tools**. The **Office Deployment Tool (`setup.exe`)** is downloaded at install time from **Microsoft’s official Office CDN** (this repo does not ship Microsoft binaries). We provide **presets** in `configs\`, a single shared **`M365AppsCore.ps1`** (ODT helpers + full language list), and **one-command** bootstrap plus separate **GUI** and **console** installers.

## One command (recommended)

**Graphical installer** — open **PowerShell as Administrator**, paste, Enter:

```powershell
irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex
```

That fetches **`M365AppsCore.ps1`**, the chosen installer script, and **`configs\`**, into `%TEMP%`, then starts the installer.

### Modes (`office.ps1`)

Set **`OFFICE_AUTO_INSTALL_MODE`** before `irm ... | iex` (or use defaults):

| Who | Mode | What happens |
|-----|------|----------------|
| **Standard user** | `gui` (default) | WPF installer. |
| **Terminal prompts** | `console` | `Install-Office.ps1`. |
| **IT automation** | `deploy` | Downloads **`Deploy-Microsoft365Apps.ps1`** + **`configs\`**, then silent ODT install. **Elevated** PowerShell required. |

**Examples:**

```powershell
$env:OFFICE_AUTO_INSTALL_MODE = "gui"
irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex
```

```powershell
$env:OFFICE_AUTO_INSTALL_MODE = "console"
irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex
```

```powershell
$env:OFFICE_AUTO_INSTALL_MODE = "deploy"
$env:OFFICE_AUTO_INSTALL_PRESET = "O365ProPlus-VDI"
$env:OFFICE_AUTO_INSTALL_LANGUAGE = "en-us"
irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex
```

**Backward compatibility:** `OFFICE_AUTO_INSTALL_USE_CONSOLE=1` selects **console** if `MODE` is unset.

**Deploy mode — optional environment variables**

| Variable | Purpose |
|----------|---------|
| `OFFICE_AUTO_INSTALL_PRESET` | Preset name (default `O365ProPlus`). |
| `OFFICE_AUTO_INSTALL_LANGUAGE` / `OFFICE_AUTO_INSTALL_LANGUAGEID` | e.g. `en-us`. |
| `OFFICE_AUTO_INSTALL_CHANNEL` | e.g. `MonthlyEnterprise`, `Current`. |
| `OFFICE_AUTO_INSTALL_ARCH` | `32` or `64`. |
| `OFFICE_AUTO_INSTALL_UNINSTALL` | `1` removes Microsoft 365 Apps (Click-to-Run). |
| `OFFICE_AUTO_INSTALL_CONFIGURATION_FILE` | Path to your own ODT XML. |
| `OFFICE_AUTO_INSTALL_WORKING_DIRECTORY` | Working folder for ODT/setup. |
| `OFFICE_AUTO_INSTALL_SKIP_PREREQ` | `1` skips disk/network checks (testing). |
| `OFFICE_AUTO_INSTALL_SKIP_ADMIN` | `1` skips admin check (testing only). |
| `OFFICE_AUTO_INSTALL_EXCLUDE_APPS` | Comma-separated ODT `ExcludeApp` IDs (e.g. `Teams,OneDrive,Access`) merged into the suite product for preset or custom XML deploys. |

**Forks:** `OFFICE_AUTO_INSTALL_REPO`, `OFFICE_AUTO_INSTALL_BRANCH`.

**Raw GitHub** (if Pages is blocked):  
`irm "https://raw.githubusercontent.com/njvanas/Office-Auto-Install/main/office.ps1" | iex`

---

## Offline / cloned repository

Primary languages match **Microsoft 365 Apps** (ODT culture IDs). After dot-sourcing **`M365AppsCore.ps1`**, run `Get-M365AppsSupportedLanguages`.

```text
powershell -NoProfile -ExecutionPolicy Bypass -File ".\Install-Office-GUI-WPF.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\Install-Office.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\Deploy-Microsoft365Apps.ps1" -Preset O365ProPlus-VDI -LanguageId en-us
```

## Automation (direct script use)

```powershell
.\Deploy-Microsoft365Apps.ps1 -Preset O365ProPlus-VDI -LanguageId en-us
.\Deploy-Microsoft365Apps.ps1 -Preset O365ProPlus -Channel MonthlyEnterprise -LanguageId en-us
.\Deploy-Microsoft365Apps.ps1 -ConfigurationFile 'D:\Deploy\company.xml'
.\Deploy-Microsoft365Apps.ps1 -Preset O365ProPlus -LanguageId en-us -ExcludeApp Teams,OneDrive
.\Deploy-Microsoft365Apps.ps1 -Uninstall
```

**Parameters:** `-Preset`, `-ConfigurationFile`, `-Uninstall`, `-OfficeClientEdition`, `-Channel`, `-LanguageId`, `-ExcludeApp`, `-WorkingDirectory`, `-SkipPrerequisiteTest`, `-SkipAdministratorCheck`.

**Excluding apps:** The GUI and console installers offer optional exclusions. For automation, pass `-ExcludeApp` (repeatable or comma-separated) or set `OFFICE_AUTO_INSTALL_EXCLUDE_APPS`. IDs are the same as ODT `<ExcludeApp ID="…"/>` (for example `Access`, `Excel`, `Groove`, `Lync`, `OneDrive`, `OneNote`, `Outlook`, `OutlookForWindows`, `PowerPoint`, `Publisher`, `Teams`, `Word`). Merges apply to **suite** products in the XML; standalone Visio/Project products in multi-product presets are left unchanged. See [configuration options](https://learn.microsoft.com/microsoft-365-apps/deploy/office-deployment-tool-configuration-options).

### Presets (`configs\`)

| File | Use case |
|------|-----------|
| `O365ProPlus.xml` | Microsoft 365 Apps for enterprise, persistent machines |
| `O365ProPlus-VDI.xml` | AVD / Windows 365 — shared computer licensing; Teams/OneDrive excluded (deploy separately if needed) |
| `O365Business.xml` / `O365Business-VDI.xml` | Microsoft 365 Apps for business |
| `O365ProPlusVisioProject.xml` / `…-VDI` | M365 Apps + **Visio LTSC 2021** + **Project LTSC 2021** (volume licenses) |
| `O365ProPlusVisioProject-Retail.xml` / `…-Retail-VDI` | M365 Apps + **VisioProRetail** + **ProjectProRetail** (Microsoft 365 subscription / Click-to-Run Retail — Visio Plan 2, Project Plan 3/5) |
| `O365ProPlusVisioProject-2024.xml` / `…-2024-VDI` | M365 Apps + **Visio LTSC 2024** + **Project LTSC 2024** (volume) |
| `Uninstall-Microsoft365Apps.xml` | Removal |

**Visio / Project product IDs:** Microsoft’s ODT uses **`VisioProRetail`** / **`ProjectProRetail`** for subscription deployments alongside Microsoft 365 Apps, and **`VisioPro2024Volume`** / **`ProjectPro2024Volume`** (or **`*2024Retail`**) for Office LTSC 2024–aligned perpetual licenses. The older **`VisioPro2021Volume`** / **`ProjectPro2021Volume`** presets remain for LTSC 2021 volume agreements. See [supported product IDs](https://learn.microsoft.com/office365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run) and the [configuration options](https://learn.microsoft.com/microsoft-365-apps/deploy/office-deployment-tool-configuration-options) Product element.

Default presets exclude **new Outlook** (`OutlookForWindows`); edit XML to include or switch exclusions per [Microsoft’s configuration options](https://learn.microsoft.com/microsoft-365-apps/deploy/office-deployment-tool-configuration-options).

### Shared engine (`M365AppsCore.ps1`)

Dot-sourced by the installers (not run directly). Provides `Save-M365AppsOfficeDeploymentTool` (official CDN), `Start-M365AppsSetup`, `Get-M365AppsSupportedLanguages`, and configuration helpers.

### Official Microsoft references

| Topic | Link |
|-------|------|
| Office Deployment Tool (download) | https://www.microsoft.com/download/details.aspx?id=49117 |
| ODT overview | https://learn.microsoft.com/microsoft-365-apps/deploy/overview-office-deployment-tool |
| Configuration options | https://learn.microsoft.com/microsoft-365-apps/deploy/office-deployment-tool-configuration-options |
| Win32 app packaging for Intune (optional) | https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool |
| Microsoft license terms | https://www.microsoft.com/licensing/terms |

## Files

| Path | Role |
|------|------|
| `office.ps1` | Bootstrap (`irm \| iex`): downloads core + installer + `configs\`. |
| `M365AppsCore.ps1` | Shared engine (languages + ODT helpers); dot-sourced by other scripts. |
| `Install-Office-GUI-WPF.ps1` | Graphical installer (Windows PowerShell 5.1). |
| `Install-Office.ps1` | Console wizard. |
| `Deploy-Microsoft365Apps.ps1` | Silent / automation entry. |
| `configs\` | ODT XML presets. |
| `NOTICES.md` | License notices. |

## Requirements

- Valid **Microsoft 365 / Office license** for what you install.  
- **Administrator** rights to install or remove Office.  
- Network access to **Microsoft** endpoints for ODT and Office content.

## Language compatibility (built-in rules)

Deployments that include **Visio and/or Project** (volume) cannot use **en-gb**, **fr-ca**, or **es-mx** as the **primary** language in the same install — this matches Microsoft’s documented limitation for those apps. The GUI language list is **filtered** for Visio/Project profiles; `Assert-M365AppsLanguageCompatibleWithDeployment` in **`M365AppsCore.ps1`** blocks invalid combos in console and deploy (not for custom `-ConfigurationFile`).

## Troubleshooting: “language is not available”

1. **Blocked by this project** — Incompatible Visio/Project + language pairs fail early with a clear error (see above).

2. **Unknown / free-form culture** — **Unknown Office language** means the string was not recognized. Use a culture such as `en-us`, or pick from the GUI. Valid-looking `xx-yy` tags not yet in the catalog may pass through with a warning, then **Assert** still applies for Visio/Project.

3. **Microsoft setup (`setup.exe`)** — Other “language not available” messages come from **Microsoft’s CDN** for your **product/channel** combo. See [Overview of deploying languages for Microsoft 365 Apps](https://learn.microsoft.com/deployoffice/overview-deploying-languages-microsoft-365-apps).

## Safety

Uses **Microsoft’s** supported installation mechanisms only. **No** licensing bypass. **Use at your own risk.**

## License

[MIT License](./LICENSE)
