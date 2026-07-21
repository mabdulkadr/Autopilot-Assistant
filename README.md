
# 🚀 Autopilot Assistant

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)
![GUI](https://img.shields.io/badge/GUI-WPF-purple.svg)
![Version](https://img.shields.io/badge/version-2.0-green.svg)

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-☕-FFDD00?style=for-the-badge)](https://www.buymeacoffee.com/mabdulkadrx)

---

## 📖 Overview

**Autopilot Assistant** is a modern WPF-based PowerShell tool designed to simplify **Windows Autopilot onboarding** and device provisioning workflows.

It provides a structured, production-ready interface for:

- Collecting **Hardware Hash (HWID)**
- Connecting securely to **Microsoft Graph (Interactive delegated auth)**
- Importing devices into **Intune Autopilot**
- Preventing duplicate uploads automatically
- Monitoring upload status with structured results
- Logging all operations locally

---

## 🖥 Screenshot

![Screenshot](Screenshot.png)

---

## ✨ Core Features

### 🔹 Device Information Panel
- Device Model
- Device Name
- Manufacturer
- Serial Number
- Free Storage (GB)
- TPM Version
- Internet Connectivity Status
- Session details (Machine / User / Elevation)

### 🔹 Method 1 — Collect HWID (No Upload)
- Exports compliant Autopilot CSV
- Does not require Graph connection
- Ideal for offline preparation

Output format:

```csv
Device Serial Number,Windows Product ID,Hardware Hash
"ABC12345","00330-80000-00000-AAOEM","<base64HWID>"
```

### 🔹 Method 2 — Upload to Intune Autopilot
- Connect to Microsoft Graph (browser interactive)
- Upload: CSV, JSON, or auto-generate single-device CSV
- Optional: Group Tag, Assigned User (UPN), Assigned Computer Name
- Built-in duplicate detection by serial number
- Automatic status polling (Complete / Queued / Failed)

### 🔹 Upload Result Engine
After upload, a structured results window shows:
- ✅ Success
- ⏳ Queued
- ⚠ Duplicate (Skipped)
- ❌ Failed

With summary counts: `Total | Success | Failed | Duplicate | Queued`

### 🔹 Retry Engine
- Automatically tracks failed rows
- "Retry Failed Uploads" re-submits only failed records
- Prevents reprocessing successful entries

### 🔹 Message Center
Real-time log console with levels: INFO, SUCCESS, WARN, ERROR

Logs are also written to:
```
C:\AutopilotAssistant\Logs\app_YYYYMMDD.log
```

---

## ⚙️ Requirements

### System
- Windows 10 / 11
- Windows PowerShell 5.1
- Run as **Administrator** (required for HWID collection)

### Modules

Required:

```powershell
Microsoft.Graph.Authentication
MSAL.PS
```

Install (recommended once per system):

```powershell
Install-Module Microsoft.Graph.Authentication -Scope AllUsers -Force
Install-Module MSAL.PS -Scope AllUsers -Force
```

---

## 🔐 Microsoft Graph Permissions (Delegated)

The tool requests:
- `User.Read`
- `Device.Read.All`
- `DeviceManagementServiceConfig.ReadWrite.All`

User must have sufficient Intune permissions.

---

## 🚀 How to Run

### Option 1 — PowerShell Script

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\Autopilot Assistant.ps1
```

### Option 2 — Packaged EXE (Self-Signed with PSWrap)

The `.exe` was compiled and self-signed using [PSWrap](https://github.com/mabdulkadr/PSWrap) — a PowerShell Script to EXE Compiler with built-in code signing support.

```
Autopilot Assistant.exe
```

No PowerShell console required.

---

## 🔄 Typical Workflow

1. Launch tool as Administrator
2. Review device information
3. Click **Connect** (Microsoft Graph)
4. Choose upload method:
   - Provide CSV
   - Or leave path empty to auto-create
5. Click **Upload Device to Autopilot**
6. Review results window

---

## 📂 Supported Input Formats

### CSV

```csv
Device Serial Number,Windows Product ID,Hardware Hash
"PF12345","00330-80000-00000-AAOEM","<base64HWID>"
```

---

## 📊 Operational Safeguards

- Pre-check: Admin rights validation
- Pre-check: Internet connectivity
- Pre-check: Graph module presence
- Pre-check: Serial duplication in Autopilot
- Polling logic for import status
- Proxy-aware Graph connectivity
- In-memory token caching (session only)

---

## 📁 Folder Structure

```
C:\AutopilotAssistant\
 ├── HWID\
 ├── Logs\
```

No data stored outside this root path.

---

## 🐛 Troubleshooting

| Symptom | Likely Cause | Fix |
|---------|--------------|-----|
| HWID collection fails | Not running as Administrator | Right-click → Run as Administrator |
| Graph connect fails | No internet or proxy blocking | Check network/proxy settings |
| Module not found error | MSAL.PS or Graph module missing | Run `Install-Module` commands from Requirements |
| Upload shows "Duplicate" | Serial already exists in Autopilot | Use "Retry Failed" or check existing devices |
| UI freezes during upload | Rare; check background jobs | Restart the tool |

---

## 🛡 Design Principles

- PowerShell 5.1 compatible
- STA RunspacePool (no UI freeze)
- DispatcherTimer async model
- No device-code authentication
- Delegated Graph authentication
- Clean WPF card-style interface

---

## 📜 License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

## 👤 Author

**Mohammad Abdulkader Omar**
Website: https://momar.tech
LinkedIn: https://www.linkedin.com/in/mabdulkadr/
Version: **2.0**

---

## ☕ Support

If this project helps you, consider supporting it:

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-☕-FFDD00?style=for-the-badge)](https://www.buymeacoffee.com/mabdulkadrx)

---

## ⚠ Disclaimer

These scripts are provided as-is. Test them in a staging environment before applying them to production. The author is not responsible for any unintended outcomes resulting from their use.
