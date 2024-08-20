# Winget Update Tools

![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?logo=gnometerminal&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-0078D6?logo=windows&logoColor=white)
![GitHub release (latest by date)](https://img.shields.io/github/v/release/sterbweise/winget-update)
![GitHub last commit](https://img.shields.io/github/last-commit/sterbweise/winget-update)
![GitHub issues](https://img.shields.io/github/issues/sterbweise/winget-update)
![GitHub stars](https://img.shields.io/github/stars/sterbweise/winget-update)
![GitHub licence](https://img.shields.io/github/license/sterbweise/winget-update)

An advanced PowerShell script to automate application updates via winget.

## Table of Contents

- [Winget Update Tools](#winget-update-tools)
  - [Table of Contents](#table-of-contents)
  - [About](#about)
  - [Features](#features)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
  - [Usage](#usage)
    - [Main Options](#main-options)
  - [Update Modes](#update-modes)
  - [Examples](#examples)
    - [Standard Update](#standard-update)
    - [Silent Update Excluding Certain Applications](#silent-update-excluding-certain-applications)
    - [Adding Applications to the Permanent Exclusion List](#adding-applications-to-the-permanent-exclusion-list)
    - [Forced Update with Custom Parameters](#forced-update-with-custom-parameters)
  - [Contributing](#contributing)
  - [Licence](#licence)

## About

Winget-Update is a PowerShell script designed to simplify and automate the process of updating applications via the Windows Package Manager (winget). It offers increased flexibility and advanced features for users and system administrators.

## Features

- Automatic application updates via winget
- Multiple update modes (normal, silent, forced, etc.)
- Temporary or permanent application exclusion
- Support for custom parameters for winget
- User-friendly command-line interface

## Prerequisites

- Windows 10 (version 1809 or later) or Windows 11
- PowerShell 5.1 or later
- [Windows Package Manager (winget)](https://github.com/microsoft/winget-cli)

## Installation

1. Ensure winget is installed on your system.
   - To check, open PowerShell and type `winget --version`
   - If winget is not recognised, install it from the [Microsoft Store](https://www.microsoft.com/p/app-installer/9nblggh4nns1)

2. Obtain the `winget-update.ps1` script:
   
   Option A: Clone the Git repository (recommended)
   - Open PowerShell
   - Navigate to the folder where you want to clone the repository
   - Run the command:
     ```
     git clone https://github.com/sterbweise/winget-update.git
     ```

   Option B: Download the script directly
   - Visit [https://github.com/sterbweise/winget-update](https://github.com/sterbweise/winget-update)
   - Click on the `winget-update.ps1` file
   - Click on the "Raw" button
   - Right-click and select "Save as..."
   - Choose the save location and click "Save"

3. Add the folder containing the script to your PATH environment variable:
   - Open PowerShell as administrator
   - Run the following command, replacing `C:\Path\To\The\Folder` with the actual path to the folder containing the script:
     ```powershell
     [Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Path\To\The\Folder", "Machine")
     ```
   - Close and reopen PowerShell for the changes to take effect

4. Create a PowerShell alias for the script:
   - Open your PowerShell profile by typing `notepad $PROFILE` in PowerShell
   - If the file doesn't exist, create it
   - Add the following line to the file:
     ```powershell
     Set-Alias -Name winget-update -Value winget-update.ps1
     ```
   - Save and close the file
   - Reload your PowerShell profile by typing `. $PROFILE`

You can now run `winget-update` from any location in PowerShell.

## Usage

Open PowerShell and navigate to the directory containing the script. Run it with the desired parameters:
powershell
.\winget-update.ps1 [options]

### Main Options

- `-ExcludeApps` or `-e`: Specifies applications to exclude from the current update.
- `-Mode` or `-m`: Sets the update mode.
- `-AddPermanentExcludeApps` or `-ape`: Adds applications to the permanent exclusion list.
- `-RemovePermanentExcludeApps` or `-rpe`: Removes applications from the permanent exclusion list.
- `-CustomParams` or `-cp`: Specifies custom parameters to pass directly to winget.
- `-Help`: Displays detailed help for the script.

## Update Modes

- `normal`: Default mode, interactive update.
- `silent`: Silent mode, automatically accepts all agreements.
- `force`: Forced mode, ignores version checks and bypasses some restrictions.
- `verbose`: Provides detailed logging information.
- `no-interaction`: Runs without user interaction, suitable for automated scripts.
- `full-upgrade`: Includes unknown versions and pinned packages in the upgrade.
- `safe-upgrade`: Performs a conservative upgrade, accepting only package agreements.

## Examples

### Standard Update
    .\winget-update.ps1

### Silent Update Excluding Certain Applications
    .\winget-update.ps1 -ExcludeApps "App1,App2" -Mode silent

### Adding Applications to the Permanent Exclusion List
    .\winget-update.ps1 -AddPermanentExcludeApps "App3,App4"

### Forced Update with Custom Parameters
    .\winget-update.ps1 -Mode force -CustomParams "--no-upgrade"

## Contributing

Contributions are welcome! Feel free to open an issue or submit a pull request.

1. Fork the project
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Licence

This project is licensed under the MIT Licence. See the [LICENCE](LICENCE) file for details.

---

Developed with ❤️ by [Sterbweise](https://github.com/sterbweise)
