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
    Version        : v2.1.0
    Date           : 2025-02-07

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
    [string[]]$ExcludeApps = @(),

    [Parameter(Mandatory=$false)]
    [Alias("m")]
    [ValidateSet("normal", "silent", "force", "verbose", "no-interaction", "full-upgrade", "safe-upgrade")]
    [string]$Mode = "normal",

    [Parameter(Mandatory=$false)]
    [Alias("ape", "add-persistent-exclude")]
    [string[]]$AddPersistentExcludeApps = @(),

    [Parameter(Mandatory=$false)]
    [Alias("rpe", "remove-persistent-exclude")]
    [string[]]$RemovePersistentExcludeApps = @(),

    [Parameter(Mandatory=$false)]
    [Alias("cp", "custom-params")]
    [string]$CustomParams = "",

    [Parameter(Mandatory=$false)]
    [switch]$Help
)

# ... rest of the script ...

# Define the path for the persistent exclude list file
$PersistentExcludeFile = Join-Path $PSScriptRoot "persistent_exclude_apps.txt"

# Create the persistent exclude file if it doesn't exist
if (-not (Test-Path $PersistentExcludeFile)) {
    New-Item -Path $PersistentExcludeFile -ItemType File -Force | Out-Null
    Write-Host "Created persistent exclude file: $PersistentExcludeFile"
}

# Function to display comprehensive help information
function Show-Help {
    # Get console width for better centering
    $width = [Math]::Min(100, $Host.UI.RawUI.WindowSize.Width - 1)
    $separator = "─" * $width
    $doubleSeparator = "═" * $width

    function Write-CenteredText {
        param(
            [string]$Text,
            [string]$Color = "White",
            [switch]$NoNewline
        )
        $spaces = " " * [Math]::Max(0, [Math]::Floor(($width - $Text.Length) / 2))
        if ($NoNewline) {
            Write-Host ($spaces + $Text) -ForegroundColor $Color -NoNewline
        } else {
            Write-Host ($spaces + $Text) -ForegroundColor $Color
        }
    }

    Clear-Host
    Write-Host $doubleSeparator -ForegroundColor Cyan
    Write-CenteredText "WINGET UPDATE HELP" "Cyan"
    Write-Host $doubleSeparator -ForegroundColor Cyan
    Write-Host ""

    Write-Host "🎯 SYNOPSIS" -ForegroundColor Yellow
    Write-Host $separator -ForegroundColor DarkGray
    Write-Host "    Advanced PowerShell script for automating application updates using the Windows"
    Write-Host "    Package Manager (winget). Features multiple operation modes, exclusion capabilities,"
    Write-Host "    and extensive customization options."
    Write-Host ""

    Write-Host "📋 USAGE" -ForegroundColor Yellow
    Write-Host $separator -ForegroundColor DarkGray
    Write-Host "    .\winget-update.ps1 [-ExcludeApps <apps>] [-Mode <mode>] [-CustomParams <params>]"
    Write-Host "    .\winget-update.ps1 [-AddPersistentExcludeApps <apps>] [-RemovePersistentExcludeApps <apps>]"
    Write-Host "    .\winget-update.ps1 [-Help]"
    Write-Host ""

    Write-Host "⚙️ OPTIONS" -ForegroundColor Yellow
    Write-Host $separator -ForegroundColor DarkGray
    
    $options = @(
        @{
            Name = "-ExcludeApps, -e, -exclude"
            Description = "Specify applications to exclude from the current update session"
            Example = "-ExcludeApps `"App1,App2,App3`""
            Color = "Green"
        },
        @{
            Name = "-Mode, -m"
            Description = "Set the update mode (see UPDATE MODES section below)"
            Example = "-Mode silent"
            Color = "Green"
        },
        @{
            Name = "-AddPersistentExcludeApps, -ape"
            Description = "Add applications to the persistent exclude list"
            Example = "-AddPersistentExcludeApps `"App1,App2`""
            Color = "Green"
        },
        @{
            Name = "-RemovePersistentExcludeApps, -rpe"
            Description = "Remove applications from the persistent exclude list"
            Example = "-RemovePersistentExcludeApps `"App1,App2`""
            Color = "Green"
        },
        @{
            Name = "-CustomParams, -cp"
            Description = "Specify custom parameters to pass directly to winget"
            Example = "-CustomParams `"--no-upgrade`""
            Color = "Green"
        },
        @{
            Name = "-Help"
            Description = "Display this help message"
            Example = "-Help"
            Color = "Green"
        }
    )

    foreach ($option in $options) {
        Write-Host "    $($option.Name)" -ForegroundColor $option.Color
        Write-Host "        $($option.Description)"
        Write-Host "        Example: $($option.Example)"
        Write-Host ""
    }

    Write-Host "🔄 UPDATE MODES" -ForegroundColor Yellow
    Write-Host $separator -ForegroundColor DarkGray
    
    $modes = @(
        @{
            Name = "normal"
            Description = "Default interactive mode with standard update behavior"
            Icon = "🔵"
        },
        @{
            Name = "silent"
            Description = "Silent mode, automatically accepts all agreements"
            Icon = "🔇"
        },
        @{
            Name = "force"
            Description = "Forces updates, bypasses most restrictions and warnings"
            Icon = "⚡"
        },
        @{
            Name = "verbose"
            Description = "Provides detailed logging information for troubleshooting"
            Icon = "📝"
        },
        @{
            Name = "no-interaction"
            Description = "Runs without any user interaction, suitable for automation"
            Icon = "🤖"
        },
        @{
            Name = "full-upgrade"
            Description = "Includes unknown versions and pinned packages in the upgrade"
            Icon = "🔄"
        },
        @{
            Name = "safe-upgrade"
            Description = "Conservative upgrade mode with minimal risk"
            Icon = "🛡️"
        }
    )

    foreach ($mode in $modes) {
        Write-Host "    $($mode.Icon) $($mode.Name)" -ForegroundColor Cyan
        Write-Host "        $($mode.Description)"
        Write-Host ""
    }

    Write-Host "📚 EXAMPLES" -ForegroundColor Yellow
    Write-Host $separator -ForegroundColor DarkGray
    $examples = @(
        @{
            Command = ".\winget-update.ps1"
            Description = "Run update with default settings"
        },
        @{
            Command = ".\winget-update.ps1 -Mode silent"
            Description = "Run silent update without user interaction"
        },
        @{
            Command = ".\winget-update.ps1 -ExcludeApps `"Microsoft.Edge,Mozilla.Firefox`""
            Description = "Update all apps except specified browsers"
        },
        @{
            Command = ".\winget-update.ps1 -Mode force -CustomParams `"--no-upgrade`""
            Description = "Force check for updates without upgrading"
        },
        @{
            Command = ".\winget-update.ps1 -AddPersistentExcludeApps `"App1,App2`""
            Description = "Add apps to persistent exclusion list"
        }
    )

    foreach ($example in $examples) {
        Write-Host "    📎 $($example.Command)" -ForegroundColor Magenta
        Write-Host "       $($example.Description)"
        Write-Host ""
    }

    Write-Host "ℹ️ NOTES" -ForegroundColor Yellow
    Write-Host $separator -ForegroundColor DarkGray
    Write-Host "    • Version     : v2.1.0"
    Write-Host "    • Author      : sterbweise"
    Write-Host "    • Last Update : 2025-02-07"
    Write-Host "    • Requirements: PowerShell 5.1+, Windows Package Manager (winget)"
    Write-Host "    • Repository  : https://github.com/Sterbweise/winget-update"
    Write-Host ""

    Write-Host "💡 TIPS" -ForegroundColor Yellow
    Write-Host $separator -ForegroundColor DarkGray
    Write-Host "    • Use silent mode for automated tasks"
    Write-Host "    • Check logs in ./logs directory for troubleshooting"
    Write-Host "    • Persistent excludes persist across sessions"
    Write-Host "    • Use verbose mode when reporting issues"
    Write-Host ""

    Write-Host $doubleSeparator -ForegroundColor Cyan
    Write-CenteredText "Press any key to exit" "Cyan"
    Write-Host $doubleSeparator -ForegroundColor Cyan
    
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# Display help if requested
if ($Help) {
    Show-Help
}

# Function to retrieve the persistent exclude list
function Get-PersistentExcludeList {
    if (Test-Path $PersistentExcludeFile) {
        return Get-Content $PersistentExcludeFile
    }
    return @()
}

# Function to add applications to the persistent exclude list
function Add-PersistentExcludeApps {
    param(
        [string[]]$Apps
    )
    $currentList = Get-PersistentExcludeList
    $newApps = $Apps | Where-Object { $currentList -notcontains $_ }
    if ($newApps.Count -gt 0) {
        $newList = $currentList + "`n" + $newApps
        $newList | Set-Content $PersistentExcludeFile
        Write-Host "Added to persistent exclude list: $($newApps -join ', ')"
    } else {
        Write-Host "No new apps to add to the persistent exclude list."
    }
    (Get-Content $PersistentExcludeFile) | Where-Object {$_.Trim() -ne ""} | Set-Content $PersistentExcludeFile
}

# Function to remove applications from the persistent exclude list
function Remove-PersistentExcludeApps {
    param(
        [string[]]$Apps
    )
    $currentList = Get-PersistentExcludeList
    $newList = $currentList | Where-Object { $Apps -notcontains $_ }
    $newList | Set-Content $PersistentExcludeFile
    Write-Host "Removed from persistent exclude list: $($Apps -join ', ')"
    (Get-Content $PersistentExcludeFile) | Where-Object {$_.Trim() -ne ""} | Set-Content $PersistentExcludeFile
}

# Add new persistent exclude apps if specified
if ($AddPersistentExcludeApps) {
    Add-PersistentExcludeApps -Apps $AddPersistentExcludeApps
    exit
}

# Remove persistent exclude apps if specified
if ($RemovePersistentExcludeApps) {
    Remove-PersistentExcludeApps -Apps $RemovePersistentExcludeApps
    exit
}

# Retrieve the current persistent exclude list
$CurrentPersistentExcludeApps = Get-PersistentExcludeList

# Function to escape special regex characters
function Escape-SpecialCharacters {
    param(
        [string]$String
    )
    return [regex]::Escape($String)
}

# Function to get the list of available updates from winget
function Get-WingetUpdates {
    $wingetOutput = winget update --include-unknown | Out-String
    $lines = $wingetOutput -split "`r`n"
    
    $updates = @()
    $headerFound = $false
    $startProcessing = $false
    
    foreach ($line in $lines) {
        # Skip warning and information lines
        if ($line -match "following packages have|explicit targeting") {
            continue
        }

        if ($line -match "Name\s+Id\s+Version\s+Available\s+Source") {
            $headerFound = $true
            continue
        }

        # Wait for a line after the header before starting
        if ($headerFound -and $line -match "^-+") {
            $startProcessing = $true
            continue
        }

        if ($startProcessing -and $line.Trim() -ne "" -and $line -match "[a-zA-Z0-9]") {
            # Use precise regex to capture columns
            if ($line -match "^([^│]+?)\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)$") {
                $name = $matches[1].Trim()
                $id = $matches[2].Trim()
                $version = $matches[3].Trim()
                $available = $matches[4].Trim()
                $source = $matches[5].Trim()

                # Verify it's a valid line
                if ($id -ne "Id" -and $version -ne "Version" -and 
                    -not ($name -match "following packages|explicit targeting")) {
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

    # Filter duplicates and return only valid updates
    return @($updates | Where-Object { $_.Id -notmatch "explicit|upgrade available" } | Sort-Object Id -Unique)
}

function Show-UpdateSummary {
    param(
        [Array]$Updates = @(),
        [string]$Mode = "normal",
        [string[]]$ExcludeApps = @(),
        [string[]]$PersistentExcludeApps = @(),
        [string]$CustomParams = ""
    )

    Clear-Host
    $currentVersion = "v2.1.0"
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

    $width = 72  # Largeur totale augmentée
    $nameWidth = 30
    $versionWidth = 19
    $availableWidth = 20
    $border = "─" * ($width-2)
    Write-Host "`n┌" -NoNewline -ForegroundColor Blue
    Write-Host ("─" * 28) -NoNewline -ForegroundColor Blue
    Write-Host " " -NoNewline
    Write-Host "WINGET UPDATE" -NoNewline -ForegroundColor Cyan
    Write-Host " " -NoNewline
    Write-Host ("─" * 27) -NoNewline -ForegroundColor Blue
    Write-Host "┐" -ForegroundColor Blue

    # Version information section
    Write-Host "│ " -NoNewline -ForegroundColor Blue
    Write-Host "ℹ️  Script Version: " -NoNewline
    if ($toolVersionColors.Count -eq 2) {
        $parts = $toolVersionStatus.Split('→')
        Write-Host $parts[0].Trim() -ForegroundColor $toolVersionColors[0] -NoNewline
        Write-Host " → " -NoNewline
        Write-Host ($parts[1].Trim() + " ").PadRight($width - 32) -ForegroundColor $toolVersionColors[1] -NoNewline
    } else {
        Write-Host ($toolVersionStatus + " ").PadRight($width - 32) -ForegroundColor $toolVersionColors[0] -NoNewline
    }
    Write-Host "│" -ForegroundColor Blue

    # Updates count section
    Write-Host "│ " -NoNewline -ForegroundColor Blue
    Write-Host ("📦 Available Updates: " + $Updates.Count + " packages").PadRight($width-3) -NoNewline -ForegroundColor Yellow
    Write-Host "│" -ForegroundColor Blue

    # Mode section
    Write-Host "│ " -NoNewline -ForegroundColor Blue
    Write-Host ("⚙️  Mode: " + $Mode).PadRight($width-3) -NoNewline -ForegroundColor DarkCyan
    Write-Host "│" -ForegroundColor Blue

    # Custom params if present
    if ($CustomParams) {
        Write-Host "│ " -NoNewline -ForegroundColor Blue
        Write-Host ("🛠️  Params: " + $CustomParams).PadRight($width-3) -NoNewline -ForegroundColor DarkCyan
        Write-Host "│" -ForegroundColor Blue
    }

    # Separator
    Write-Host "├$border┤" -ForegroundColor Blue

    # Updates list header
    if ($Updates.Count -gt 0) {
        Write-Host "│ " -NoNewline -ForegroundColor Blue
        Write-Host "NAME".PadRight($nameWidth) -NoNewline -ForegroundColor Cyan
        Write-Host "VERSION".PadRight($versionWidth) -NoNewline -ForegroundColor Cyan
        Write-Host "AVAILABLE".PadRight($availableWidth) -NoNewline -ForegroundColor Cyan
        Write-Host "│" -ForegroundColor Blue
        Write-Host "├$border┤" -ForegroundColor Blue

        # List each update
        foreach ($update in $Updates) {
            $isExcluded = ($ExcludeApps -contains $update.Id) -or ($PersistentExcludeApps -contains $update.Id)
            
            Write-Host "│ " -NoNewline -ForegroundColor Blue
            
            # Name truncation with ellipsis if necessary
            $name = if ($update.Name.Length -gt $nameWidth) {
                $update.Name.Substring(0, $nameWidth-3) + "..."
            } else {
                $update.Name
            }
            Write-Host $name.PadRight($nameWidth) -NoNewline -ForegroundColor $(if ($isExcluded) { "DarkGray" } else { "White" })
            
            # Version
            Write-Host $update.Version.PadRight($versionWidth) -NoNewline -ForegroundColor $(if ($isExcluded) { "DarkGray" } else { "Yellow" })
            
            # Available version with arrow
            $availableText = "→     " + $update.Available
            Write-Host $availableText.PadRight($availableWidth) -NoNewline -ForegroundColor $(if ($isExcluded) { "DarkGray" } else { "Green" })
            
            Write-Host "│" -ForegroundColor Blue
        }
    }

    # Exclusions section if any
    if ($ExcludeApps.Count -gt 0 -or $PersistentExcludeApps.Count -gt 0) {
        Write-Host "├$border┤" -ForegroundColor Blue
        Write-Host "│ " -NoNewline -ForegroundColor Blue
        Write-Host "🚫 Excluded:".PadRight($width-3) -NoNewline -ForegroundColor Yellow
        Write-Host "│" -ForegroundColor Blue

        if ($PersistentExcludeApps.Count -gt 0) {
            $persExcludes = ($PersistentExcludeApps -join ", ")
            $firstLine = "  Persistent: " + $persExcludes
            $maxLength = $width - 4  # -4 pour "│ " et " │"

            $lines = for ($i = 0; $i -lt $firstLine.Length; $i += $maxLength) {
                if ($i -eq 0) {
                    $firstLine.Substring($i, [Math]::Min($maxLength, $firstLine.Length - $i))
                } else {
                    "    " + $firstLine.Substring($i, [Math]::Min($maxLength - 4, $firstLine.Length - $i))
                }
            }

            foreach ($line in $lines) {
                Write-Host "│ " -NoNewline -ForegroundColor Blue
                Write-Host $line.PadRight($width-3) -NoNewline -ForegroundColor DarkMagenta
                Write-Host "│" -ForegroundColor Blue
            }
        }
    }

    # Bottom border
    Write-Host "└$border┘" -ForegroundColor Blue
    Write-Host ""

    # Single prompt for confirmation
    if ($Updates.Count -gt 0) {
        Write-Host "Continue with updates? (" -NoNewline
        Write-Host "Y" -NoNewline -ForegroundColor Green
        Write-Host "/" -NoNewline
        Write-Host "N" -NoNewline -ForegroundColor Red
        Write-Host "): " -NoNewline
    }
}

function Write-Status {
    param(
        [string]$Status,
        [string]$Icon,
        [string]$Color = "White",
        [switch]$NoNewLine
    )
    
    # Retour au début de la ligne d'état
    Write-Host "`r   " -NoNewline
    # Effacer la ligne précédente
    Write-Host (" " * 80) -NoNewline
    Write-Host "`r   " -NoNewline
    Write-Host $Icon -NoNewline
    Write-Host " " -NoNewline
    if ($NoNewLine) {
        Write-Host $Status -ForegroundColor $Color -NoNewline
    } else {
        Write-Host $Status -ForegroundColor $Color
    }
}

function Write-ProgressBar {
    param(
        [string]$Current,
        [string]$Total,
        [int]$PercentComplete
    )
    $width = 40
    $completed = [math]::Round($width * ($PercentComplete / 100))
    $remaining = $width - $completed
    
    $progressBar = "█" * $completed + "░" * $remaining
    Write-Host "`r   📥 " -NoNewline
    Write-Host "$progressBar $Current / $Total ($PercentComplete%)" -NoNewline -ForegroundColor Yellow
}

# Fonction helper pour convertir les tailles en MB
function Convert-ToMB {
    param (
        [double]$size,
        [string]$unit
    )
    
    switch ($unit) {
        "KB" { return $size / 1024 }
        "MB" { return $size }
        "GB" { return $size * 1024 }
        default { return $size }
    }
}

function Get-RemoteFileSize {
    param([string]$Url)
    try {
        $request = [System.Net.WebRequest]::Create($Url)
        $request.Method = "HEAD"
        $response = $request.GetResponse()
        $fileSize = $response.ContentLength
        $response.Close()
        
        # Convert to appropriate size format
        if ($fileSize -gt 1GB) {
            return "$([math]::Round($fileSize/1GB, 2)) GB"
        }
        elseif ($fileSize -gt 1MB) {
            return "$([math]::Round($fileSize/1MB, 2)) MB"
        }
        else {
            return "$([math]::Round($fileSize/1KB, 2)) KB"
        }
    }
    catch {
        return "Unknown"
    }
}

function Update-Apps {
    param(
        [Array]$Updates = @(),
        [string]$Mode = "normal",
        [string]$CustomParams = "",
        [string[]]$ExcludeApps = @(),
        [string[]]$PersistentExcludeApps = @()
    )
    Clear-Host
    # Créer le dossier de logs s'il n'existe pas
    $logDir = Join-Path $PSScriptRoot "logs"
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir | Out-Null
    }

    # Créer un nouveau fichier log pour cette session
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $sessionLogFile = Join-Path $logDir "winget_update_$timestamp.log"
    
    # Fonction pour écrire dans le log
    function Write-Log {
        param([string]$Message)
        $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
        Add-Content -Path $sessionLogFile -Value $logMessage
    }

    $baseCommand = "winget"
    $additionalParams = @()
    
    switch ($Mode) {
        "normal" { }
        "silent" { $additionalParams += @("--silent", "--accept-package-agreements") }
        "force" { $additionalParams += @("--force", "--accept-package-agreements", "--accept-source-agreements") }
        "verbose" { $additionalParams += @("--verbose") }
        "no-interaction" { $additionalParams += @("--silent", "--accept-package-agreements") }
        "full-upgrade" { $additionalParams += @("--include-unknown", "--accept-package-agreements") }
        "safe-upgrade" { $additionalParams += @("--accept-package-agreements") }
    }

    if ($CustomParams) {
        $additionalParams += $CustomParams -split ' '
    }

    function Format-FileSize {
        param([string]$Size)
        if ($Size -match "(\d+\.?\d*)\s*(KB|MB|GB|B)") {
            $value = [double]$matches[1]
            $unit = $matches[2]
            switch ($unit) {
                "B"  { return "$value B" }
                "KB" { return "$([math]::Round($value, 2)) KB" }
                "MB" { return "$([math]::Round($value, 2)) MB" }
                "GB" { return "$([math]::Round($value, 2)) GB" }
            }
        }
        return $Size
    }

    # Calculer le nombre réel de packages à mettre à jour (exclus les packages exclus)
    $totalApps = @($Updates | Where-Object {
        $id = $_.Id
        ($ExcludeApps -notcontains $id) -and ($PersistentExcludeApps -notcontains $id)
    }).Count

    $currentApp = 0
    $updateResults = @()

    $downloadStarted = $false
    $lastProgressLine = ""

    foreach ($update in $Updates) {
        if (($ExcludeApps -notcontains $update.Id) -and ($PersistentExcludeApps -notcontains $update.Id)) {
            $currentApp++
            $logPath = $null  # Initialisation de la variable logPath
            
            # Package Header
            Write-Host "`n┌─────────────────────────────────────────────────────────────┐" -ForegroundColor Blue
            Write-Host "│ 📦 Package $currentApp of $totalApps" -ForegroundColor Blue
            Write-Host "└─────────────────────────────────────────────────────────────┘" -ForegroundColor Blue
            Write-Host "   Name        : " -NoNewline; Write-Host $update.Name -ForegroundColor Cyan
            Write-Host "   ID          : " -NoNewline; Write-Host $update.Id -ForegroundColor DarkGray
            Write-Host "   Version     : " -NoNewline
            Write-Host $update.Version -NoNewline -ForegroundColor Yellow
            Write-Host " → " -NoNewline -ForegroundColor Gray
            Write-Host $update.Available -ForegroundColor Green

            # Get package details
            try {
                $packageInfo = winget show --id $update.Id --accept-source-agreements | Out-String
                
                # Extract package information using regex (removed Author)
                $details = @{
                    'Publisher' = if ($packageInfo -match "Publisher:\s*(.+)") { $matches[1] } else { "Unknown" }
                    'License' = if ($packageInfo -match "License:\s*(.+)") { $matches[1] } else { "Unknown" }
                }

                # Try to get size from Download Size field first
                if ($packageInfo -match "Download Size:\s*(.+)") {
                    $details.Size = Format-FileSize $matches[1]
                }
                # If not found, try to get from Installer URL
                elseif ($packageInfo -match "Installer Url:\s*(.+)") {
                    $details.Size = Get-RemoteFileSize $matches[1]
                }
                else {
                    $details.Size = "Unknown"
                }

                Write-Host "   Publisher   : " -NoNewline; Write-Host $details.Publisher -ForegroundColor DarkCyan
                Write-Host "   License     : " -NoNewline; Write-Host $details.License -ForegroundColor DarkCyan
                Write-Host "   Size        : " -NoNewline; Write-Host $details.Size -ForegroundColor DarkCyan
            } catch {
                Write-Host "   ⚠️ Could not retrieve additional package information" -ForegroundColor Yellow
            }

            Write-Status -Status "Starting update process..." -Icon "🔄" -Color "Yellow"

            try {
                $downloadStarted = $false
                $lastLine = ""
                
                # Créer un processus avec redirection synchrone
                $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                $pinfo.FileName = "winget"
                $pinfo.Arguments = "upgrade --id $($update.Id) $($additionalParams -join ' ')"
                $pinfo.RedirectStandardOutput = $true
                $pinfo.RedirectStandardError = $true
                $pinfo.UseShellExecute = $false
                $pinfo.CreateNoWindow = $true

                $p = New-Object System.Diagnostics.Process
                $p.StartInfo = $pinfo
                $p.Start() | Out-Null

                # Lire la sortie caractère par caractère pour un affichage en temps réel
                $reader = $p.StandardOutput
                $buffer = ""
                
                while (-not $reader.EndOfStream) {
                    $char = [char]$reader.Read()
                    if ($char -eq "`n") {
                        $line = $buffer.Trim()
                        Write-Log $line

                        if ($line -match "Downloading (.+)") {
                            $downloadUrl = $matches[1]
                            Write-Status -Status "Downloading from: $downloadUrl" -Icon "⏬" -Color "Yellow"
                            $downloadStarted = $true
                        }
                        elseif ($downloadStarted -and $line -match "(\d+(?:\.\d+)?)\s*(KB|MB|GB)\s*/\s*(\d+(?:\.\d+)?)\s*(KB|MB|GB)") {
                            try {
                                $currentValue = [double]$matches[1]
                                $currentUnit = $matches[2]
                                $totalValue = [double]$matches[3]
                                $totalUnit = $matches[4]

                                $currentSize = "$currentValue $currentUnit"
                                $totalSize = "$totalValue $totalUnit"

                                $currentBytes = Convert-ToMB $currentValue $currentUnit
                                $totalBytes = Convert-ToMB $totalValue $totalUnit

                                if ($totalBytes -gt 0) {
                                    $percent = [math]::Min([math]::Round(($currentBytes / $totalBytes) * 100), 100)
                                    Write-ProgressBar -Current $currentSize -Total $totalSize -PercentComplete $percent
                                }
                            }
                            catch {
                                Write-Log "Error processing download progress: $_"
                            }
                        }
                        elseif ($line -match "Successfully verified installer hash") {
                            if ($downloadStarted) {
                                Write-Host ""
                            }
                            Write-Status -Status "Hash verification successful" -Icon "✅" -Color "Green"
                            $downloadStarted = $false
                        }
                        elseif ($line -match "Starting package install") {
                            Write-Status -Status "Installing package..." -Icon "🔧" -Color "Yellow"
                        }
                        elseif ($line -match "Successfully installed") {
                            Write-Status -Status "Installation completed successfully!" -Icon "✅" -Color "Green"
                        }
                        elseif ($line -match "Installer failed with exit code: (.+)") {
                            Write-Status -Status "Installation failed (Error code: $($matches[1]))" -Icon "❌" -Color "Red"
                        }

                        $buffer = ""
                    }
                    else {
                        $buffer += $char
                    }
                }

                $p.WaitForExit()
                
                # Vérifier si le processus s'est terminé normalement
                if ($p.ExitCode -eq 0) {
                    $status = "Success"
                }
                else {
                    $status = "Failed"
                    Write-Status -Status "Process exited with code: $($p.ExitCode)" -Icon "❌" -Color "Red"
                }

                $updateResults += [PSCustomObject]@{
                    Name = $update.Name
                    Id = $update.Id
                    OldVersion = $update.Version
                    NewVersion = $update.Available
                    Status = $status
                    ExitCode = $p.ExitCode
                    LogFile = $logPath
                    Size = $details.Size
                }

            } catch {
                Write-Status -Status "Error: $_" -Icon "❌" -Color "Red"
                $updateResults += [PSCustomObject]@{
                    Name = $update.Name
                    Id = $update.Id
                    Status = "Error"
                    ExitCode = -1
                    LogFile = $null
                    Size = "Unknown"
                }
            }
        }
    }

    # Final Summary
    Write-Host "`n┌" -NoNewline -ForegroundColor Blue
    Write-Host ("─" * 23) -NoNewline -ForegroundColor Blue
    Write-Host " " -NoNewline
    Write-Host "UPDATE SUMMARY" -NoNewline -ForegroundColor Cyan
    Write-Host " " -NoNewline
    Write-Host ("─" * 23) -NoNewline -ForegroundColor Blue
    Write-Host "┐" -ForegroundColor Blue

    $successCount = @($updateResults | Where-Object { $_.Status -eq "Success" }).Count
    $failCount = @($updateResults | Where-Object { $_.Status -ne "Success" }).Count

    Write-Host "📊 Results Summary" -ForegroundColor Yellow
    Write-Host "   ✅ Successful: " -NoNewline -ForegroundColor Green
    Write-Host $successCount
    Write-Host "   ❌ Failed   : " -NoNewline -ForegroundColor Red
    Write-Host $failCount

    if ($updateResults.Count -gt 0) {
        Write-Host "`n📋 Detailed Results:" -ForegroundColor Yellow
        Write-Host "─" * 80 -ForegroundColor DarkGray
        foreach ($result in $updateResults) {
            Write-Host "   • " -NoNewline
            Write-Host $result.Name -NoNewline -ForegroundColor Cyan
            Write-Host " ($($result.Id))" -NoNewline -ForegroundColor DarkGray
            Write-Host " - " -NoNewline
            Write-Host $result.Status -ForegroundColor $(
                switch ($result.Status) {
                    "Success" { "Green" }
                    "Failed" { "Red" }
                    default { "Yellow" }
                }
            )
            if ($result.Status -eq "Success") {
                Write-Host "     ↳ " -NoNewline
                Write-Host "$($result.OldVersion) → $($result.NewVersion)" -ForegroundColor Green
                if ($result.Size -and $result.Size -ne "Unknown") {
                    Write-Host "     ↳ Size: $($result.Size)" -ForegroundColor DarkCyan
                }
            }
        }
    }

    # À la fin du script, afficher le chemin du log de session
    Write-Host "`n📝 Session log: " -NoNewline
    Write-Host $sessionLogFile -ForegroundColor Cyan
}

# Main program execution
$updates = Get-WingetUpdates

# Handle potential null array
if (@($updates).Count -gt 0) {
    Show-UpdateSummary -Updates $updates -Mode $Mode -ExcludeApps $ExcludeApps -PersistentExcludeApps $CurrentPersistentExcludeApps -CustomParams $CustomParams
    
    $confirmation = Read-Host

    if ($confirmation -eq "Y") {
        Update-Apps -Updates $updates -Mode $Mode -CustomParams $CustomParams -ExcludeApps $ExcludeApps -PersistentExcludeApps $CurrentPersistentExcludeApps
    } else {
        Write-Host "Update cancelled." -ForegroundColor Red
    }
} else {
    Write-Host "No updates available." -ForegroundColor Green
}
