<div align="center">

# Winget Update Manager
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=flat-square&logo=powershell&logoColor=white&labelColor=2C3E50)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D4?style=flat-square&logo=windows&logoColor=white&labelColor=34495E)
![License](https://img.shields.io/badge/License-MIT-50C878?style=flat-square&logo=opensourceinitiative&logoColor=white&labelColor=2C3E50)
![Version](https://img.shields.io/badge/Version-3.0.0-FF6B6B?style=flat-square&logoColor=white&labelColor=34495E)
![GitHub Stars](https://img.shields.io/github/stars/sterbweise/winget-update?style=flat-square&logo=github&color=FFD700&labelColor=2C3E50)
![GitHub Issues](https://img.shields.io/github/issues/sterbweise/winget-update?style=flat-square&logo=github&color=FF4757&labelColor=34495E)
![Last Commit](https://img.shields.io/github/last-commit/sterbweise/winget-update?style=flat-square&logo=git&color=2ED573&labelColor=34495E)
![Repo Size](https://img.shields.io/github/repo-size/sterbweise/winget-update?style=flat-square&logo=database&color=FFA726&labelColor=2C3E50)

</div>

<div align="center">
<img width="656" height="435" alt="image" src="https://github.com/user-attachments/assets/abfdb2de-4c08-444a-82cd-3680b71c6168" />
</div>

## Summary

**Winget Update Manager** is a comprehensive PowerShell automation tool designed for Windows system administrators and power users. It provides intelligent application lifecycle management using Windows Package Manager (winget) with advanced features including smart application detection, persistent exclusion management, and comprehensive logging capabilities.


<table align="center">
<tr>
<td align="center">🎯<br><b>Smart Detection</b><br>Intelligent app matching</td>
<td align="center">⚡<br><b>Multiple Modes</b><br>8 execution modes</td>
<td align="center">🛡️<br><b>Usage Ready</b><br>Advanced logging</td>
<td align="center">🔧<br><b>Easy Setup</b><br>One-line installation</td>
</tr>
</table>



## Quick Installation

### One-Line Installation

```powershell
iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/sterbweise/winget-update/main/install.ps1'))
```

*After installation, use `winget-update` from anywhere in PowerShell*


<<<<<<< HEAD
=======
<img src="https://github.com/user-attachments/assets/5863fdf3-6cd8-47d8-b346-4a574fc45c61" alt="image" width="600"/>
>>>>>>> 6d720eba713b73cd8438e17312523732c3b83334

## Table of Contents

<div align="center">

| [🔧 Installation](#installation) | [⚙️ System Requirements](#system-requirements) | [🎮 Quick Start](#quick-start) |
|:---:|:---:|:---:|
| [📚 Usage Guide](#usage) | [🔧 Configuration](#configuration) | [🚀 Advanced Features](#advanced-features) |
| [🛠️ Troubleshooting](#troubleshooting) | [🤝 Contributing](#contributing) | [📄 License](#license) |

</div>

---

## 🔧 Installation

<details>
<summary><b>📦 One-Line Installation (Recommended)</b></summary>

```powershell
iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/sterbweise/winget-update/main/install.ps1'))
```

**What this does:**
- ✅ Downloads and installs the latest version
- ✅ Creates a global `winget-update` command
- ✅ Sets up the necessary directory structure
- ✅ Configures PowerShell profile integration

</details>

<details>
<summary><b>📥 Manual Installation</b></summary>

```powershell
# Download the script
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/sterbweise/winget-update/main/winget-update.ps1" -OutFile "winget-update.ps1"

# Set execution policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Run the script
.\winget-update.ps1
```

</details>

<details>
<summary><b>🔄 Git Installation</b></summary>

```powershell
git clone https://github.com/sterbweise/winget-update.git
cd winget-update
.\winget-update.ps1 -Help
```

</details>


## ⚙️ System Requirements

<div align="center">

| Component | Requirement | Status |
|:---:|:---:|:---:|
| **🖥️ OS** | Windows 10 (1809+) / Windows 11 | ![Required](https://img.shields.io/badge/-Required-red?style=flat-square) |
| **⚡ PowerShell** | 5.1 / 7.0+ | ![Required](https://img.shields.io/badge/-Required-red?style=flat-square) |
| **📦 Winget** | 1.0+ | ![Required](https://img.shields.io/badge/-Required-red?style=flat-square) |
| **🔐 Privileges** | Administrator | ![Auto](https://img.shields.io/badge/-Auto--Prompt-yellow?style=flat-square) |
| **💾 Memory** | 512 MB RAM | ![Minimal](https://img.shields.io/badge/-Minimal-blue?style=flat-square) |
| **💿 Storage** | 100 MB | ![Minimal](https://img.shields.io/badge/-Minimal-blue?style=flat-square) |

</div>

## 🎮 Quick Start

### Basic Usage

<table>
<tr>
<td width="50%">

**🔄 Standard Update**
```powershell
winget-update
```
*Interactive mode with modern interface and real-time progress*

**👀 Preview Mode**
```powershell
winget-update -Mode dry-run
```
*See what would be updated without making changes*

</td>
<td width="50%">

**🤫 Silent Mode**
```powershell
winget-update -Mode silent
```
*Automated execution for scripts and scheduled tasks*

**🚀 Full Upgrade**
```powershell
winget-update -Mode full-upgrade
```
*Updates all packages including unknown versions*

</td>
</tr>
</table>


## 📚 Usage

### Execution Modes

<div align="center">

| Mode | Description | Use Case | Icon |
|:---:|:---|:---|:---:|
| `normal` | Interactive with prompts | Desktop environments | 🖥️ |
| `silent` | Automated execution | Scheduled tasks | 🤫 |
| `dry-run` | Simulation mode | Testing & planning | 👀 |
| `force` | Bypass restrictions | Emergency updates | ⚡ |
| `verbose` | Detailed logging | Debugging | 📝 |
| `no-interaction` | Zero user input | Server environments | 🤖 |
| `full-upgrade` | Include all packages | Complete refresh | 🚀 |
| `safe-upgrade` | Conservative approach | Production systems | 🛡️ |

</div>

### 🎯 Smart Exclusion Management

<details>
<summary><b>Temporary Exclusions</b></summary>

```powershell
# Exclude apps from current session only
winget-update -ExcludeApps "Microsoft.Edge,Discord.Discord"
```
- ✅ Supports both friendly names and exact IDs
- ✅ Use comma separation for multiple applications
- ✅ Only affects current update session

</details>

<details>
<summary><b>Permanent Exclusions</b></summary>

```powershell
# Smart addition to permanent exclusion list
winget-update -AddPersistentExcludeApps "Visual Studio Code,Chrome,Spotify"

# Smart removal from permanent exclusion list
winget-update -RemovePersistentExcludeApps "Discord,Spotify"
```
- 🧠 Intelligent name matching
- 🔍 Shows suggestions for ambiguous matches
- 💾 Automatically saves to `persistent_exclude_apps.txt`

</details>

### ⚙️ Advanced Parameters

<details>
<summary><b>Custom Parameters</b></summary>

```powershell
# Pass custom parameters to winget
winget-update -CustomParams "--include-unknown --force"

# Production automation example
winget-update -Mode silent -ExcludeApps "Microsoft.VisualStudio.2022.Community" -CustomParams "--accept-source-agreements"

# Maximum logging for debugging
winget-update -Mode verbose -CustomParams "--include-unknown --verbose-logs"
```

</details>


## 🔧 Configuration

### 📁 File Structure

```
📦 winget-update/
├── 📜 winget-update.ps1              # Main executable script
├── 📜 install.ps1                    # Installation script
├── 📋 persistent_exclude_apps.txt    # Permanent exclusion list
├── 📁 logs/                          # Log file directory
│   └── 📄 winget_update_*.log         # Timestamped execution logs
└── 📖 README.md                      # This documentation
```

### ⚙️ Smart Application Categories

<div align="center">

| Category | Priority | Examples |
|:---:|:---:|:---|
| **Development** | 🔴 High | Visual Studio, Git, Docker |
| **Security** | 🔴 High | Antivirus, VPN, Firewall |
| **Browser** | 🟡 Medium | Chrome, Firefox, Edge |
| **Productivity** | 🟡 Medium | Office, Teams, Notion |
| **Media** | 🟢 Low | VLC, Spotify, OBS |
| **Gaming** | 🟢 Low | Steam, Discord, Epic |

</div>



## 🚀 Advanced Features

### Intelligent Detection

<table>
<tr>
<td width="50%">

**🔍 Smart Name Matching**
```powershell
# Input: "Visual Studio Code"
# Output: Microsoft.VisualStudioCode
```

**🎯 Fuzzy Search**
```powershell
# Input: "Chrome"
# Output: Multiple matches with selection
```

</td>
<td width="50%">

**📊 Automatic Categorization**
- High Priority: Dev tools, security
- Medium Priority: Browsers, productivity
- Low Priority: Entertainment, gaming

**⚡ Performance Optimization**
- Concurrent update processing
- Intelligent retry mechanisms
- Resource usage monitoring

</td>
</tr>
</table>

### 🤖 Automation Integration

<details>
<summary><b>Windows Task Scheduler</b></summary>

```powershell
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-Command winget-update -Mode silent"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At "2:00 AM"
Register-ScheduledTask -TaskName "WingetUpdateManager" -Action $action -Trigger $trigger -RunLevel Highest
```

</details>

<details>
<summary><b>PowerShell Profile</b></summary>

```powershell
# Add to $PROFILE for custom aliases
function Update-Apps { winget-update -Mode silent }
function Preview-Updates { winget-update -Mode dry-run }
function Update-Verbose { winget-update -Mode verbose }
```

</details>


## 🛠️ Troubleshooting

<div align="center">

### Common Issues Quick Fix

</div>

<details>
<summary><b>❌ winget Command Not Found</b></summary>

**Symptoms:** `'winget' is not recognized as an internal or external command`

**Solutions:**
```powershell
# Method 1: Install from Microsoft Store
# Search for "App Installer" and install

# Method 2: Manual installation
Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe

# Method 3: Download from GitHub
# Visit: https://github.com/microsoft/winget-cli/releases
```

</details>

<details>
<summary><b>🔒 Script Execution Policy Errors</b></summary>

**Symptoms:** `Execution of scripts is disabled on this system`

**Solutions:**
```powershell
# Recommended: Set policy for current user
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Alternative: Bypass for single execution
powershell.exe -ExecutionPolicy Bypass -File ".\winget-update.ps1"
```

</details>

<details>
<summary><b>🛡️ Access Denied During Updates</b></summary>

**Symptoms:** Permission errors, "Access is denied" messages

**Solutions:**
```powershell
# Check admin status
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
Write-Host "Running as Admin: $isAdmin"

# Run as Administrator
Start-Process powershell -Verb RunAs -ArgumentList "-File `"$PWD\winget-update.ps1`""
```

</details>

<details>
<summary><b>🔍 Smart Detection Issues</b></summary>

**Symptoms:** Applications not found, "No matches found" errors

**Solutions:**
```powershell
# List all installed applications
winget list | Out-GridView

# Search with partial name
winget search "Visual Studio"

# Use exact ID
winget-update -AddPersistentExcludeApps "Microsoft.VisualStudioCode"
```

</details>

### 🔧 Advanced Debugging

<details>
<summary><b>System Information Collection</b></summary>

```powershell
$info = @{
    "PowerShell Version" = $PSVersionTable.PSVersion
    "Windows Version" = (Get-CimInstance Win32_OperatingSystem).Caption
    "Winget Version" = (winget --version)
    "Execution Policy" = (Get-ExecutionPolicy)
    "Is Admin" = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}
$info | Format-Table -AutoSize
```

</details>



## 🤝 Contributing

<div align="center">

### We Welcome Contributions!

</div>

<table align="center">
<tr>
<td align="center">🍴<br><b>Fork</b><br>Fork the repository</td>
<td align="center">🌿<br><b>Branch</b><br>Create feature branch</td>
<td align="center">✨<br><b>Develop</b><br>Add your changes</td>
<td align="center">🧪<br><b>Test</b><br>Test thoroughly</td>
<td align="center">📤<br><b>Submit</b><br>Create pull request</td>
</tr>
</table>

### 🛠️ Development Setup

```powershell
# Clone repository
git clone https://github.com/sterbweise/winget-update.git
cd winget-update

# Install development tools
Install-Module -Name Pester -Force
Install-Module -Name PSScriptAnalyzer -Force

# Run code analysis
Invoke-ScriptAnalyzer -Path ".\winget-update.ps1"
```



## 📄 License

<div align="center">

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for full details.

**Summary:** You are free to use, modify, distribute, and sell this software. No warranty is provided.

</div>

---

<div align="center">

## 🌟 Support the Project

If you find this project helpful, please consider:

[![⭐ Star](https://img.shields.io/badge/⭐-Star%20this%20repo-yellow?style=for-the-badge)](https://github.com/sterbweise/winget-update)
[![🐛 Report Bug](https://img.shields.io/badge/🐛-Report%20Bug-red?style=for-the-badge)](https://github.com/sterbweise/winget-update/issues)
[![💡 Request Feature](https://img.shields.io/badge/💡-Request%20Feature-blue?style=for-the-badge)](https://github.com/sterbweise/winget-update/issues)

---

**❤️ Developed by [Sterbweise](https://github.com/sterbweise)**

*Making Windows package management intelligent and effortless*

![GitHub Profile](https://img.shields.io/badge/GitHub-Sterbweise-181717?style=for-the-badge&logo=github)

</div>
