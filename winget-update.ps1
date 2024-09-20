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
    Version        : v2.0.0
    Date           : 2024-09-19

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
    $width = 100
    $separator = "=" * $width

    function Write-ColoredLine($text, $color) {
        Write-Host $text.PadRight($width) -ForegroundColor $color
    }

    Clear-Host
    Write-ColoredLine $separator "Cyan"
    Write-ColoredLine "Winget-Update - Advanced PowerShell Script for Updating Applications via winget" "Cyan"
    Write-ColoredLine $separator "Cyan"
    Write-Host ""

    Write-ColoredLine "SYNOPSIS" "Yellow"
    Write-Host "    This script provides a robust solution for automating application updates using the"
    Write-Host "    Windows Package Manager (winget)."
    Write-Host ""

    Write-ColoredLine "USAGE" "Yellow"
    Write-Host "    .\winget-update.ps1 [-ExcludeApps <app1,app2,...>] [-Mode <mode>]"
    Write-Host "                        [-AddPermanentExcludeApps <app1,app2,...>]"
    Write-Host "                        [-RemovePermanentExcludeApps <app1,app2,...>]"
    Write-Host "                        [-CustomParams <params>] [-Help]"
    Write-Host ""

    Write-ColoredLine "OPTIONS" "Yellow"
    Write-Host "    -ExcludeApps, -e, -exclude <app1,app2,...>" -ForegroundColor Green
    Write-Host "        Specify applications to exclude from the current update session (comma-separated)"
    Write-Host ""
    Write-Host "    -Mode, -m <mode>" -ForegroundColor Green
    Write-Host "        Set the update mode (normal, silent, force, verbose, no-interaction, full-upgrade, safe-upgrade)"
    Write-Host ""
    Write-Host "    -AddPermanentExcludeApps, -ape, -add-permanent-exclude <app1,app2,...>" -ForegroundColor Green
    Write-Host "        Add applications to the permanent exclude list (comma-separated)"
    Write-Host ""
    Write-Host "    -RemovePermanentExcludeApps, -rpe, -remove-permanent-exclude <app1,app2,...>" -ForegroundColor Green
    Write-Host "        Remove applications from the permanent exclude list (comma-separated)"
    Write-Host ""
    Write-Host "    -CustomParams, -cp, -custom-params <params>" -ForegroundColor Green
    Write-Host "        Specify custom parameters to pass directly to winget"
    Write-Host ""
    Write-Host "    -Help" -ForegroundColor Green
    Write-Host "        Display this comprehensive help message"
    Write-Host ""

    Write-ColoredLine "UPDATE MODES" "Yellow"
    Write-Host "    normal         : Default mode, interactive update" -ForegroundColor Cyan
    Write-Host "    silent         : Silent mode, automatically accepts all agreements" -ForegroundColor Cyan
    Write-Host "    force          : Forced mode, ignores version checks and bypasses some restrictions" -ForegroundColor Cyan
    Write-Host "    verbose        : Provides detailed logging information" -ForegroundColor Cyan
    Write-Host "    no-interaction : Runs without any user interaction, suitable for automated scripts" -ForegroundColor Cyan
    Write-Host "    full-upgrade   : Includes unknown versions and pinned packages in the upgrade" -ForegroundColor Cyan
    Write-Host "    safe-upgrade   : Performs a conservative upgrade, accepting only package agreements" -ForegroundColor Cyan
    Write-Host ""

    Write-ColoredLine "EXAMPLES" "Yellow"
    Write-Host "    .\winget-update.ps1" -ForegroundColor Magenta
    Write-Host "    .\winget-update.ps1 -ExcludeApps `"App1,App2`" -Mode silent" -ForegroundColor Magenta
    Write-Host "    .\winget-update.ps1 -e `"App1,App2`" -m force" -ForegroundColor Magenta
    Write-Host "    .\winget-update.ps1 -AddPermanentExcludeApps `"App3,App4`"" -ForegroundColor Magenta
    Write-Host "    .\winget-update.ps1 -RemovePermanentExcludeApps `"App3,App4`"" -ForegroundColor Magenta
    Write-Host "    .\winget-update.ps1 -CustomParams `"--no-upgrade`"" -ForegroundColor Magenta
    Write-Host ""

    Write-ColoredLine "NOTES" "Yellow"
    Write-Host "    Author  : sterbweise" -ForegroundColor White
    Write-Host "    Version : v1.0.0" -ForegroundColor White
    Write-Host "    Date    : 2024-08-20" -ForegroundColor White
    Write-Host ""
    Write-Host "    This script is designed for both casual users and system administrators,"
    Write-Host "    providing flexibility and power in managing system updates."
    Write-Host ""

    Write-ColoredLine $separator "Cyan"
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
    $wingetOutput = winget update --include-unknown | Out-String
    $lines = $wingetOutput -split "`r`n"
    
    $updates = @()
    $headerFound = $false
    $headerLine = ""
    
    foreach ($line in $lines) {
        if ($line -match "Name\s+Id\s+Version\s+Available\s+Source") {
            $headerFound = $true
            $headerLine = $line
            continue
        }

        if ($headerFound -and $line.Trim() -ne "" -and $line -match "[a-zA-Z0-9]") {
            # Utiliser une expression régulière pour extraire les informations
            if ($line -match "^(.*?)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)$") {
                $name = $matches[1].Trim()
                $id = $matches[2].Trim()
                $version = $matches[3].Trim()
                $available = $matches[4].Trim()
                $source = $matches[5].Trim()

                if ($id -and $id -ne "Id") {
                    # Vérifier si l'application existe
                    $appExists = winget show -e --id $id | Out-String
                    if ($appExists -match $id) {
                        $updates += [PSCustomObject]@{
                            Name = $name
                            Id = $id
                            Version = $version
                            Available = $available
                            Source = $source
                        }
                    }
                }
            }
        }
    }

    return $updates
}

function Show-UpdateSummary {
    param(
        [Array]$Updates,
        [string]$Mode,
        [string[]]$ExcludeApps,
        [string[]]$PermanentExcludeApps,
        [string]$CustomParams
    )
    
    $width = 100
    $separator = "─" * $width

    function Write-CenteredText($text, $color = "White") {
        $spaces = " " * (($width - $text.Length) / 2)
        Write-Host "$spaces$text" -ForegroundColor $color
    }

    Clear-Host
    Write-CenteredText "WINGET UPDATE SUMMARY" "Cyan"
    Write-Host $separator

    # Tool information and version check
    $currentVersion = "v2.0.0"
    try {
        $latestRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/Sterbweise/winget-update/releases/latest"
        $latestVersion = $latestRelease.tag_name
        if ($latestVersion -ne $currentVersion) {
            $updateUrl = "https://github.com/Sterbweise/winget-update/releases/latest"
            $toolVersionStatus = "$currentVersion → $latestVersion"
            $toolVersionColors = @("DarkRed", "DarkGreen")
            Write-Host "A new version is available! " -ForegroundColor Yellow -NoNewline
            Write-Host "(" -NoNewline
            Write-Host "$updateUrl" -ForegroundColor Cyan -NoNewline
            Write-Host ")"
        } else {
            $toolVersionStatus = "$currentVersion (Latest)"
            $toolVersionColors = @("White")
        }
    } catch {
        $toolVersionStatus = "$currentVersion (Unable to check for updates)"
        $toolVersionColors = @("Red")
    }

    $wingetVersion = (winget --version).Trim()

    # Create a table-like structure with colors only for specific values
    $tableData = @(
        @("Tool Version", $toolVersionStatus, $toolVersionColors),
        @("Winget Version", $wingetVersion, "White"),
        @("Updates Available", $Updates.Count, "White"),
        @("Update Mode", $Mode, "White"),
        @("Excluded Apps", "$($ExcludeApps.Count + $PermanentExcludeApps.Count)", "White"),
        @("Execution Time", $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), "White"),
        @("Custom Params", $(if ($CustomParams) { $CustomParams } else { "None" }), "White")
    )

    foreach ($row in $tableData) {
        $column = "{0,-20}: " -f $row[0]
        
        Write-Host $column -NoNewline
        
        if ($row[2].Count -eq 2) {
            $parts = $row[1].Split('→')
            Write-Host $parts[0].Trim() -ForegroundColor $row[2][0] -NoNewline
            Write-Host " → " -NoNewline 
            Write-Host $parts[1].Trim() -ForegroundColor $row[2][1]
        } else {
            Write-Host $row[1] -ForegroundColor $row[2]
        }
    }

    Write-Host $separator

    # Apps to update
    Write-Host "APPS TO UPDATE:" -ForegroundColor Green
    if ($Updates.Count -eq 0) {
        Write-Host "  No updates available." -ForegroundColor Cyan
    } else {
        foreach ($update in $Updates) {
            if (($ExcludeApps -notcontains $update.Id) -and ($PermanentExcludeApps -notcontains $update.Id)) {
                Write-Host ("  {0} ({1})" -f $update.Name, $update.Id) -NoNewline
                Write-Host (" {0}" -f $update.Version) -ForegroundColor DarkCyan -NoNewline
                Write-Host " →" -NoNewline
                Write-Host (" {0}" -f $update.Available) -ForegroundColor Cyan  -NoNewline
                Write-Host " from" -ForegroundColor DarkYellow -NoNewline
                Write-Host (" {0}" -f $update.Source)
            }
        }
    }

    Write-Host $separator

    # Excluded apps
    if ($ExcludeApps.Count -gt 0 -or $PermanentExcludeApps.Count -gt 0) {
        Write-Host "EXCLUDED APPS:" -ForegroundColor Yellow
        if ($ExcludeApps.Count -gt 0) {
            Write-Host "  Temporary: " -NoNewline -ForegroundColor Yellow
            Write-Host ($ExcludeApps -join ", ") -ForegroundColor DarkYellow
        }
        if ($PermanentExcludeApps.Count -gt 0) {
            Write-Host "  Permanent: " -NoNewline -ForegroundColor Magenta
            Write-Host ($PermanentExcludeApps -join ", ") -ForegroundColor DarkMagenta
        }
        Write-Host $separator
    }

    Write-CenteredText "For issues or feature requests, please visit:" "White"
    Write-CenteredText "https://github.com/Sterbweise/winget-update/issues" "Cyan"
    Write-Host $separator
}

function Update-Apps {
    param(
        [Array]$Updates,
        [string]$Mode,
        [string]$CustomParams,
        [string[]]$ExcludeApps,
        [string[]]$PermanentExcludeApps
    )

    $baseCommand = "winget"
    $additionalParams = @()
    
    # Create log directory if it doesn't exist
    $logDir = Join-Path $PSScriptRoot "logs"
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir | Out-Null
    }

    # Create a new log file with current date and time
    $currentDate = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $logFile = Join-Path $logDir "winget_update_$currentDate.log"

    # Delete log files older than 1 week
    Get-ChildItem $logDir -Filter "winget_update_*.log" | ForEach-Object {
        try {
            if ($_.BaseName -match "^winget_update_(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})") {
                $fileDate = [DateTime]::ParseExact($Matches[1], "yyyy-MM-dd_HH-mm-ss", $null)
                if ($fileDate -lt (Get-Date).AddDays(-7)) {
                    Remove-Item $_.FullName -Force -ErrorAction Stop
                }
            } else {
                Write-Warning "Skipping file with unexpected name format: $($_.FullName)"
            }
        } catch {
            Write-Warning "Error processing file $($_.FullName): $($_.Exception.Message)"
        }
    }

    switch ($Mode) {
        "normal" { }
        "silent" { $additionalParams += @("--silent", "--accept-package-agreements") }
        "force" { $additionalParams += @("--force", "--accept-package-agreements", "--accept-source-agreements", "--ignore-local-archive-malware-scan") }
        "verbose" { $additionalParams += @("--verbose-logs") }
        "no-interaction" { $additionalParams += @("--disable-interactivity", "--silent", "--accept-package-agreements", "--accept-source-agreements") }
        "full-upgrade" { $additionalParams += @("--include-unknown", "--include-pinned", "--accept-package-agreements", "--accept-source-agreements") }
        "safe-upgrade" { $additionalParams += @("--accept-package-agreements") }
    }

    if ($CustomParams) {
        $additionalParams += $CustomParams -split ' '
    }

    $totalApps = ($Updates | Where-Object { ($ExcludeApps -notcontains $_.Id) -and ($PermanentExcludeApps -notcontains $_.Id) }).Count
    $updateResults = @()

    function Write-CenteredTitle {
        param($Title)
        $width = 120
        $padding = [math]::Max(0, ($width - $Title.Length) / 2)
        $centeredTitle = (" " * $padding) + $Title
        Write-Host $centeredTitle -ForegroundColor Cyan
        Write-Host ("─" * $width)
    }

    Clear-Host
    Write-CenteredTitle "WINGET UPDATE PROGRESS"
    Write-Host ""

    # Log start of update process
    $logContent = @"
========================================
WINGET UPDATE LOG
========================================
Start Time: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Mode: $Mode
Custom Parameters: $CustomParams
Total Apps to Update: $totalApps
========================================

"@

    $logContent | Out-File -FilePath $logFile

    foreach ($update in $Updates) {
        if (($ExcludeApps -notcontains $update.Id) -and ($PermanentExcludeApps -notcontains $update.Id)) {
            $spinnerChars = '⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏'
            $spinnerIndex = 0

            function Write-UpdateLine {
                param($Spinner, $Status)
                $cursorLeft = $host.UI.RawUI.CursorPosition.X
                $cursorTop = $host.UI.RawUI.CursorPosition.Y
            
                if ($Status -eq $null) {
                    Write-Host "`r[" -NoNewline
                    Write-Host $Spinner -ForegroundColor Yellow -NoNewline
                    Write-Host "] " -NoNewline
                } else {
                    Write-Host "`r[" -NoNewline
                    if ($Status -eq "✓") {
                        Write-Host $Status -ForegroundColor Green -NoNewline
                    } else {
                        Write-Host $Status -ForegroundColor Red -NoNewline
                    }
                    Write-Host "] " -NoNewline
                }
                
                Write-Host "$($update.Name)  " -NoNewline
                Write-Host $update.Version -ForegroundColor DarkCyan -NoNewline
                Write-Host " → " -NoNewline
                Write-Host "$($update.Available) " -ForegroundColor Cyan -NoNewline
                Write-Host " from " -ForegroundColor DarkYellow -NoNewline
                Write-Host "$($update.Source)".PadRight($host.UI.RawUI.WindowSize.Width - $cursorLeft) -NoNewline
            
                $host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates $cursorLeft, $cursorTop
            }
            
            Write-UpdateLine $spinnerChars[0] $null
            
            $arguments = @("upgrade", "--id", $update.Id) + $additionalParams
            $tempOutputFile = [System.IO.Path]::GetTempFileName()
            $process = Start-Process -FilePath $baseCommand -ArgumentList $arguments -NoNewWindow -PassThru -RedirectStandardOutput $tempOutputFile
            
            while (!$process.HasExited) {
                $spinnerIndex = ($spinnerIndex + 1) % $spinnerChars.Length
                Write-UpdateLine $spinnerChars[$spinnerIndex] $null
                Start-Sleep -Milliseconds 100
            }
            
            $exitCode = $process.ExitCode
            $success = $exitCode -eq 0
            $status = if ($success) { "✓" } else { "✗" }
            
            Write-UpdateLine $null $status
        
            $updateOutput = Get-Content $tempOutputFile -Raw
            Remove-Item $tempOutputFile
            
            $updateResults += [PSCustomObject]@{
                Name = $update.Name
                Id = $update.Id
                Success = $success
            }

            # Log update result
            $logContent = @"
Update: $($update.Name) ($($update.Id))
From Version: $($update.Version)
To Version: $($update.Available)
Status: $($success ? 'Success' : 'Failure')
Exit Code: $exitCode
Output:
$updateOutput
----------------------------------------

"@
            $logContent | Out-File -FilePath $logFile -Append
            
            # Filter and log non-empty lines
            $updateOutput -split "`n" | Where-Object { 
                $_ -is [string] -and
                $_ -match "^\s*\S+" -and 
                $_ -match "\S" -and 
                $_ -match "[a-zA-Z0-9]" -and 
                $_ -notmatch "^\s*-\s*$" -and 
                $_ -ne "\" -and
                $_ -ne "-" -and
                $_.Trim() -ne ""
            } | ForEach-Object {
                $_ | Out-File -FilePath $logFile -Append
            }

            Write-Host ""
        }
    }

    Write-Host ""
    Write-Host ""
    Write-CenteredTitle "UPDATE PROCESS SUMMARY"
    Write-Host ""
    $successCount = ($updateResults | Where-Object { $_.Success }).Count
    $failCount = ($updateResults | Where-Object { -not $_.Success }).Count
    Write-Host "✅ Successful updates: " -NoNewline
    Write-Host $successCount -ForegroundColor Green
    Write-Host "❌ Failed updates:     " -NoNewline
    Write-Host $failCount -ForegroundColor Red

    # Log summary
    $logContent = @"
========================================
SUMMARY
========================================
End Time: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Successful updates: $successCount
Failed updates: $failCount
========================================

"@
    $logContent | Out-File -FilePath $logFile -Append

    if ($failCount -gt 0) {
        Write-Host "`nFailed updates:"
        $updateResults | Where-Object { -not $_.Success } | ForEach-Object {
            Write-Host "  ❌ " -ForegroundColor Red -NoNewline
            Write-Host $_.Name -ForegroundColor Yellow -NoNewline
            Write-Host " ($($_.Id))"
        }
    }

    Write-Host "`nLog file: " -NoNewline
    Write-Host $logFile -ForegroundColor Cyan
    Write-Host "`nPress any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Main program execution
$updates = Get-WingetUpdates

if ($updates.Count -gt 0) {
    Show-UpdateSummary -Updates $updates -Mode $Mode -ExcludeApps $ExcludeApps -PermanentExcludeApps $CurrentPermanentExcludeApps -CustomParams $CustomParams
    
    Write-Host "`nDo you want to proceed with the update? (" -NoNewline
    Write-Host "Y" -ForegroundColor Green -NoNewline
    Write-Host "/" -NoNewline
    Write-Host "N" -ForegroundColor Red -NoNewline
    Write-Host "): " -NoNewline
    $confirmation = Read-Host

    if ($confirmation -eq "Y") {
        Update-Apps -Updates $updates -Mode $Mode -CustomParams $CustomParams -ExcludeApps $ExcludeApps -PermanentExcludeApps $CurrentPermanentExcludeApps
    } else {
        Write-Host "Update cancelled." -ForegroundColor Red
    }
} else {
    Write-Host "No updates available." -ForegroundColor Green
}
