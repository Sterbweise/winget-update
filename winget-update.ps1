<#
.SYNOPSIS
    Advanced PowerShell script for updating applications using winget.

.DESCRIPTION
    This professional-grade script automates the process of updating applications
    using the Windows Package Manager (winget). It offers various operation modes,
    application exclusion capabilities, and custom parameter support for winget.

.NOTES
    File Name      : winget-update.ps1
    Author         : sterbweise
    Prerequisite   : PowerShell 5.1 or later, Windows Package Manager (winget)
    Version        : 1.5
    Date           : 2024-08-20

     _    _ _                  _     _   _           _       _       
    | |  | (_)                | |   | | | |         | |     | |      
    | |  | |_ _ __   __ _  ___| |_  | | | |_ __   __| | __ _| |_ ___ 
    | |/\| | | '_ \ / _` |/ _ \ __| | | | | '_ \ / _` |/ _` | __/ _ \
    \  /\  / | | | | (_| |  __/ |_  | |_| | |_) | (_| | (_| | ||  __/
     \/  \/|_|_| |_|\__, |\___|\__|  \___/| .__/ \__,_|\__,_|\__\___|
                     __/ |               | |                         
                    |___/                |_|                         

    "Elevating your system's potential, one update at a time."
#>

# Define script parameters
param(
    [Parameter(Mandatory=$false)]
    [Alias("e", "exclude")]
    [string[]]$ExcludeApps,

    [Parameter(Mandatory=$false)]
    [Alias("m")]
    [ValidateSet("normal", "silent", "force", "verbose", "no-interaction", "full-upgrade", "safe-upgrade")]
    [string]$Mode = "normal",

    [Parameter(Mandatory=$false)]
    [Alias("ape", "add-permanent-exclude")]
    [string[]]$AddPermanentExcludeApps,

    [Parameter(Mandatory=$false)]
    [Alias("rpe", "remove-permanent-exclude")]
    [string[]]$RemovePermanentExcludeApps,

    [Parameter(Mandatory=$false)]
    [Alias("cp", "custom-params")]
    [string]$CustomParams,

    [Parameter(Mandatory=$false)]
    [switch]$Help
)

# ... rest of the script ...

# Define the path for the permanent exclude list file
$PermanentExcludeFile = Join-Path $PSScriptRoot "permanent_exclude_apps.txt"

# Create the permanent exclude file if it doesn't exist
if (-not (Test-Path $PermanentExcludeFile)) {
    New-Item -Path $PermanentExcludeFile -ItemType File -Force | Out-Null
    Write-Host "Created permanent exclude file: $PermanentExcludeFile"
}

# Function to display comprehensive help information
function Show-Help {
    $helpText = @"
Winget-Update - Advanced PowerShell Script for Updating Applications via winget

SYNOPSIS
    This script provides a robust solution for automating application updates using the Windows Package Manager (winget).

USAGE
    .\winget-update.ps1 [-ExcludeApps <app1,app2,...>] [-Mode <mode>] [-AddPermanentExcludeApps <app1,app2,...>] [-RemovePermanentExcludeApps <app1,app2,...>] [-CustomParams <params>] [-Help]

OPTIONS
    -ExcludeApps, -e, -exclude <app1,app2,...>
        Specify applications to exclude from the current update session (comma-separated)

    -Mode, -m <mode>
        Set the update mode (normal, silent, force, verbose, no-interaction, full-upgrade, safe-upgrade)

    -AddPermanentExcludeApps, -ape, -add-permanent-exclude <app1,app2,...>
        Add applications to the permanent exclude list (comma-separated)

    -RemovePermanentExcludeApps, -rpe, -remove-permanent-exclude <app1,app2,...>
        Remove applications from the permanent exclude list (comma-separated)

    -CustomParams, -cp, -custom-params <params>
        Specify custom parameters to pass directly to winget

    -Help
        Display this comprehensive help message

UPDATE MODES
    normal         : Default mode, interactive update
    silent         : Silent mode, automatically accepts all agreements
    force          : Forced mode, ignores version checks and bypasses some restrictions
    verbose        : Provides detailed logging information
    no-interaction : Runs without any user interaction, suitable for automated scripts
    full-upgrade   : Includes unknown versions and pinned packages in the upgrade
    safe-upgrade   : Performs a conservative upgrade, accepting only package agreements

EXAMPLES
    .\winget-update.ps1
    .\winget-update.ps1 -ExcludeApps "App1,App2" -Mode silent
    .\winget-update.ps1 -e "App1,App2" -m force
    .\winget-update.ps1 -AddPermanentExcludeApps "App3,App4"
    .\winget-update.ps1 -RemovePermanentExcludeApps "App3,App4"
    .\winget-update.ps1 -CustomParams "--no-upgrade"


NOTES
    Author  : sterbweise
    Version : 1.5
    Date    : 2024-08-20

    This script is designed for both casual users and system administrators,
    providing flexibility and power in managing system updates.
"@
    Write-Host $helpText
    exit
}

# Display help if requested
if ($Help) {
    Show-Help
}

# Function to retrieve the permanent exclude list
function Get-PermanentExcludeList {
    if (Test-Path $PermanentExcludeFile) {
        return Get-Content $PermanentExcludeFile
    }
    return @()
}

# Function to add applications to the permanent exclude list
function Add-PermanentExcludeApps {
    param(
        [string[]]$Apps
    )
    $currentList = Get-PermanentExcludeList
    $newApps = $Apps | Where-Object { $currentList -notcontains $_ }
    if ($newApps.Count -gt 0) {
        $newList = $currentList + "`n" + $newApps
        $newList | Set-Content $PermanentExcludeFile
        Write-Host "Added to permanent exclude list: $($newApps -join ', ')"
    } else {
        Write-Host "No new apps to add to the permanent exclude list."
    }
    (Get-Content $PermanentExcludeFile) | Where-Object {$_.Trim() -ne ""} | Set-Content $PermanentExcludeFile
}

# Function to remove applications from the permanent exclude list
function Remove-PermanentExcludeApps {
    param(
        [string[]]$Apps
    )
    $currentList = Get-PermanentExcludeList
    $newList = $currentList | Where-Object { $Apps -notcontains $_ }
    $newList | Set-Content $PermanentExcludeFile
    Write-Host "Removed from permanent exclude list: $($Apps -join ', ')"
    (Get-Content $PermanentExcludeFile) | Where-Object {$_.Trim() -ne ""} | Set-Content $PermanentExcludeFile
}

# Add new permanent exclude apps if specified
if ($AddPermanentExcludeApps) {
    Add-PermanentExcludeApps -Apps $AddPermanentExcludeApps
    exit
}

# Remove permanent exclude apps if specified
if ($RemovePermanentExcludeApps) {
    Remove-PermanentExcludeApps -Apps $RemovePermanentExcludeApps
    exit
}

# Retrieve the current permanent exclude list
$CurrentPermanentExcludeApps = Get-PermanentExcludeList

# Function to get the list of available updates from winget
function Get-WingetUpdates {
    # Execute winget update command and capture the output
    $wingetOutput = winget update | Out-String
    $lines = $wingetOutput -split "`r`n"
    
    # Locate the header line in the output
    $headerLine = $lines | Where-Object { $_ -match "Name\s+Id\s+Version\s+Available\s+Source" }

    if ($headerLine) {
        # Calculate the index positions for parsing the output
        $idIndex = $headerLine.IndexOf("Id")
        $versionIndex = $headerLine.IndexOf("Version") - 1

        $appIds = @()
        foreach ($line in $lines) {
            # Parse each line to extract the application ID
            if ($line -match "^\s*\S+" -and $line -notmatch "Name\s+Id\s+Version\s+Available\s+Source" -and $line.Length -ge $headerLine.Length) {
                if ($line -match "[a-zA-Z0-9]") {
                    $appId = $line.Substring($idIndex, $versionIndex - $idIndex).Trim()
                    $appIds += $appId
                }
            }
        }
        return $appIds
    } else {
        Write-Host "Unable to find the header line in winget update output."
        return @()
    }
}

# Function to update applications
function Update-Apps {
    param(
        [string[]]$AppIds,
        [string]$Mode,
        [string]$CustomParams
    )

    # Define the base winget command
    $baseCommand = "winget upgrade --id"

    # Set additional parameters based on the selected mode
    switch ($Mode) {
        "normal" { $additionalParams = "" }
        "silent" { $additionalParams = "--silent --accept-package-agreements" }
        "force" { $additionalParams = "--force --accept-package-agreements --accept-source-agreements --ignore-local-archive-malware-scan" }
        "verbose" { $additionalParams = "--verbose-logs" }
        "no-interaction" { $additionalParams = "--disable-interactivity --silent --accept-package-agreements --accept-source-agreements" }
        "full-upgrade" { $additionalParams = "--include-unknown --include-pinned --accept-package-agreements --accept-source-agreements" }
        "safe-upgrade" { $additionalParams = "--accept-package-agreements" }
        default { $additionalParams = "" }
    }

    # Append custom parameters if provided
    if ($CustomParams) {
        $additionalParams += " $CustomParams"
    }

    # Iterate through each application ID and perform the update
    foreach ($appId in $AppIds) {
        if (($ExcludeApps -notcontains $appId) -and ($CurrentPermanentExcludeApps -notcontains $appId)) {
            Write-Host "Updating application with ID: $appId (Mode: $Mode)"
            Invoke-Expression "$baseCommand $appId $additionalParams"
        } else {
            Write-Host "Application excluded from update: $appId"
        }
    }
}

# Main program execution
$appIdsToUpdate = Get-WingetUpdates

if ($appIdsToUpdate.Count -gt 0) {
    # Display winget version information
    winget update
    Write-Host "`n"
    Write-Host "Update mode: $Mode"
    Write-Host "Temporarily excluded applications: $($ExcludeApps -join ', ')"
    Write-Host "Permanently excluded applications: $($CurrentPermanentExcludeApps -join ', ')"
    if ($CustomParams) {
        Write-Host "Custom parameters: $CustomParams"
    }
    
    # Prompt for user confirmation
    $confirmation = Read-Host "Do you want to proceed with the update? (Y/N)"
    if ($confirmation -eq "Y") {
        Update-Apps -AppIds $appIdsToUpdate -Mode $Mode -CustomParams $CustomParams
    } else {
        Write-Host "Update cancelled."
    }
} else {
    Write-Host "No updates available."
}
