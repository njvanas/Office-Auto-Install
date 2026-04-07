# Office Auto Installer

Install Microsoft Office (2024 / 2021 / Microsoft 365) using **Microsoft’s official deployment tools**.  
**Typical use:** open **PowerShell as Administrator**, **paste one command**, press **Enter**. The website shows the same steps with a simple visual: [GitHub Pages site](https://njvanas.github.io/Office-Auto-Install/).

## What you do

1. **Copy** the command below (or use the **Copy** button on the site).
2. **Open** Windows PowerShell or Terminal **as Administrator** (e.g. Win → type *PowerShell* → Ctrl+Shift+Enter).
3. **Paste** into the window (right-click often pastes), then press **Enter**.
4. Stay **online** until setup finishes (often **10–30 minutes**). You need a **valid Office license**.

## Command (recommended)

```powershell
irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex
```

Only run `irm … | iex` from URLs you trust.

### If that URL is blocked

```powershell
irm "https://raw.githubusercontent.com/njvanas/Office-Auto-Install/main/office.ps1" | iex
```

### Text-only prompts (everything in the terminal)

```powershell
$env:OFFICE_AUTO_INSTALL_USE_CONSOLE = "1"
irm "https://njvanas.github.io/Office-Auto-Install/office.ps1" | iex
```

### Forks

Before `irm`, set:

```powershell
$env:OFFICE_AUTO_INSTALL_REPO = "yourname/Office-Auto-Install"
```

Optional: `OFFICE_AUTO_INSTALL_BRANCH` for a non-default branch.

## If something fails

- Run **PowerShell as Administrator** and try again.
- **`irm` / `iex` errors:** `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` (once), then retry.
- **Network:** the bootstrap uses GitHub; Office setup uses Microsoft. Both must be reachable.
- **More help:** [Issues](https://github.com/njvanas/Office-Auto-Install/issues) · [Releases](https://github.com/njvanas/Office-Auto-Install/releases/latest)

## Repository (for contributors / offline)

| File | Role |
|------|------|
| [`office.ps1`](./office.ps1) | Small bootstrap; downloads the payload script from GitHub and runs it. |
| `Install-Office-GUI-WPF.ps1` | Default payload (installer UI). |
| `Install-Office.ps1` | Payload when `OFFICE_AUTO_INSTALL_USE_CONSOLE=1`. |

**Offline (files on disk):** open **PowerShell as Administrator**, `cd` to the folder, then:

`powershell -NoProfile -ExecutionPolicy Bypass -File ".\Install-Office-GUI-WPF.ps1"`  
(or `Install-Office.ps1` for text-only prompts).

Clone: `git clone https://github.com/njvanas/Office-Auto-Install.git`

## Safety

Uses **official Microsoft** Office Deployment Tool and CDN. **No** licensing bypass or cracks. **You** are responsible for license compliance.

## Disclaimer

This tool only automates installation through Microsoft’s supported mechanisms. **Use at your own risk.**

## Technical note

[`office.ps1`](./office.ps1) fetches the payload over HTTPS from `raw.githubusercontent.com` and executes it with `Invoke-Expression`. ODT and Office bits come from Microsoft.

## License

[MIT License](./LICENSE)
