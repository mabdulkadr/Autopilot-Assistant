<div align="center">

# 🚀 Autopilot Assistant

**Windows Autopilot onboarding, simplified**

Collect hardware hashes, import devices into Intune Autopilot, and check upload results in one WPF window.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![UI](https://img.shields.io/badge/UI-WPF%20Tailwind%20Slate-3B82F6?style=for-the-badge&logo=windows&logoColor=white)](#%EF%B8%8F-usage)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-2.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-core-features) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [Troubleshooting](#-troubleshooting) • [License](#-license)

</div>

---

# 📖 Overview

Autopilot Assistant is a WPF tool for Windows Autopilot onboarding. It lets you:

* Collect the hardware hash (HWID) from a device
* Connect to Microsoft Graph with interactive delegated sign-in
* Import devices into Intune Autopilot
* Skip uploads when a serial number is already registered
* Check import status and keep a local log of each run

---

## 🖼️ Screenshots

![Autopilot Assistant main window with the device information panel, upload methods, and Message Center](Screenshot.png)

*Main window with the device information panel, the two upload methods, and the Message Center.*

---

# ✨ Core Features

### 🔹 Device Information Panel
* Device model, name, manufacturer, serial number
* Free storage (GB) and TPM version
* Internet connectivity status
* Session details: machine, user, and elevation

### 🔹 Method 1 — Collect HWID (No Upload)
* Exports a CSV in the format Autopilot expects
* Works without a Graph connection, so you can prepare the CSV offline

```csv
Device Serial Number,Windows Product ID,Hardware Hash
"ABC12345","00330-80000-00000-AAOEM","<base64HWID>"
```

### 🔹 Method 2 — Upload to Intune Autopilot
* Sign in to Microsoft Graph in a browser window
* Upload from CSV or JSON, or create a single-device CSV in the tool
* Optional Group Tag, Assigned User (UPN), and Assigned Computer Name
* Checks for duplicate serial numbers before uploading
* Polls Intune for the import result

### 🔹 Upload Result Engine
After the upload, a results window shows one of these states per device:

| Status | Meaning |
|---|---|
| ✅ Success | Imported into Autopilot |
| ⏳ Queued | Import pending in Intune |
| ⚠ Duplicate (Skipped) | Serial already registered |
| ❌ Failed | Import error; you can retry |

The window also shows totals for all devices, successes, failures, duplicates, and queued imports.

### 🔹 Retry Engine
The tool keeps track of failed rows. "Retry Failed Uploads" submits only those rows and leaves successful ones alone.

### 🔹 Message Center
A live log console with INFO, SUCCESS, WARN, and ERROR levels. The same messages go to:

```text
C:\AutopilotAssistant\Logs\app_YYYYMMDD.log
```

---

# 📂 Project Structure

```text
Autopilot-Assistant/
│
├── Autopilot Assistant.ps1      # WPF application source
├── Autopilot Assistant.exe      # Packaged executable
├── Screenshot.png               # Documentation screenshot
└── README.md
```

---

# 🚀 Usage

### Launch

**Option 1 — PowerShell script:**

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\Autopilot Assistant.ps1
```

**Option 2 — Packaged EXE (self-signed with PSWrap):**

```text
Autopilot Assistant.exe
```

The EXE was compiled and self-signed with [PSWrap](https://github.com/mabdulkadr/PSWrap), so it runs without a PowerShell console.

### Typical Workflow

1. Launch the tool as Administrator
2. Review the device information panel
3. Click **Connect** to sign in to Microsoft Graph
4. Choose the upload method. Provide a CSV path, or leave it empty to create one
5. Click **Upload Device to Autopilot**
6. Review the results and retry any failures

### Supported Input Formats

```csv
Device Serial Number,Windows Product ID,Hardware Hash
"PF12345","00330-80000-00000-AAOEM","<base64HWID>"
```

---

# ⚙️ Requirements

| Requirement | Details |
|-------------|---------|
| **OS** | Windows 10 / 11 |
| **PowerShell** | Windows PowerShell 5.1 |
| **Elevation** | Run as Administrator; required for HWID collection |
| **Modules** | `Microsoft.Graph.Authentication` + `MSAL.PS` (auto-checked) |
| **Graph permissions** | Delegated: `User.Read`, `Device.Read.All`, `DeviceManagementServiceConfig.ReadWrite.All` |

Install the modules once per system:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope AllUsers -Force
Install-Module MSAL.PS -Scope AllUsers -Force
```

The user who signs in needs enough Intune permission to import devices.

### Data & Logs

```text
C:\AutopilotAssistant\
├── HWID\        # exported Autopilot CSV files
└── Logs\        # daily application logs
```

The tool writes no data outside this folder.

---

# 🔍 Troubleshooting

| Symptom | Likely Cause | Fix |
|---------|--------------|-----|
| HWID collection fails | Not running as Administrator | Right-click and choose Run as Administrator |
| Graph connect fails | No internet or a proxy is blocking access | Check network and proxy settings |
| Module not found error | MSAL.PS or Graph module missing | Run the `Install-Module` commands above |
| Upload shows "Duplicate" | Serial already exists in Autopilot | Check existing devices, or use Retry Failed |
| UI freezes during upload | Rare; background job issue | Restart the tool |

---

# 🛡 Operational Notes

* Before each upload, the tool checks admin rights, internet access, the Graph modules, and duplicate serial numbers
* Sign-in uses delegated Graph scopes only. There is no device-code flow and no app-only secret to manage
* Tokens stay in memory for the session; nothing auth-related is written to disk
* Graph requests follow the system proxy configuration
* Uploads run in an STA runspace pool with a dispatcher timer, so the window stays responsive
* Test Group Tag and Assigned User values on a pilot device before onboarding a full fleet

---

## 👤 Author

**Mohammad Abdelkader Omar**  
GitHub: [@mabdulkadr](https://github.com/mabdulkadr)  
Website: [momar.tech](https://momar.tech)

---

## 📜 License

This project is licensed under the [MIT License](LICENSE).

---

## ⚠ Disclaimer

This skill and every script it generates are provided as-is with no warranty
of any kind. Test generated tools in a staging environment before deploying to
production. The authors assume no liability for any damage or data loss
resulting from their use.

---
<div align="center">

⭐ **If this tool saves you time, star the repo — it helps others find it.**

[Report an Issue](../../issues) · [momar.tech](https://momar.tech)

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/mabdulkadrx)

Built with [**PowerShell Enterprise Admin**](https://github.com/mabdulkadr/powershell-enterprise-admin-skill)

</div>
