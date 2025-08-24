<#
.SYNOPSIS
    Advanced PowerShell script for updating applications using winget.

.DESCRIPTION
    This script automates the process of updating applications
    using the Windows Package Manager (winget). It offers various operation modes,
    application exclusion capabilities, and custom parameter support for winget.

.NOTES
    File Name      : winget-update.ps1
    Author         : sterbweise
    Prerequisite   : PowerShell 5.1 or later, Windows Package Manager (winget)
    Version        : v3.1.0
    Date           : 2025-07-26

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
    [ValidateSet("normal", "silent", "force", "verbose", "no-interaction", "full-upgrade", "safe-upgrade", "dry-run")]
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

# Check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Require administrator privileges for package updates
if (-not $Help -and -not (Test-Administrator)) {
    Write-Host ""
    Write-Host "⚠️  Administrator privileges required" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "This script requires administrator privileges to update system packages." -ForegroundColor White
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor White
    Write-Host ""
    Write-Host "Right-click PowerShell → 'Run as Administrator'" -ForegroundColor DarkGray
    Write-Host ""
    exit 1
}

# Define the path for the persistent exclude list file
$PersistentExcludeFile = Join-Path $PSScriptRoot "persistent_exclude_apps.txt"

# Create required directories and files
$requiredPaths = @(
    (Join-Path $PSScriptRoot "logs"),
    (Join-Path $PSScriptRoot "reports"),
    (Join-Path $PSScriptRoot "config")
)

foreach ($path in $requiredPaths) {
    if (-not (Test-Path $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
}

# Create the persistent exclude file if it doesn't exist
if (-not (Test-Path $PersistentExcludeFile)) {
    New-Item -Path $PersistentExcludeFile -ItemType File -Force | Out-Null
    Write-Host "✅ Created persistent exclude file: $PersistentExcludeFile" -ForegroundColor Green
}

# Function to clean old log files
function Clear-OldLogs {
    param([int]$RetentionDays = 30)
    
    $logDir = Join-Path $PSScriptRoot "logs"
    if (Test-Path $logDir) {
        $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
        $oldLogs = @(Get-ChildItem $logDir -Filter "*.log" | Where-Object { $_.LastWriteTime -lt $cutoffDate })
        
        if ($oldLogs.Count -gt 0) {
            $oldLogs | Remove-Item -Force
            Write-Host "🧹 Cleaned $($oldLogs.Count) old log files (older than $RetentionDays days)" -ForegroundColor Yellow
        }
    }
}

# Clean old logs on startup
Clear-OldLogs

# Function to display comprehensive help information
function Show-Help {
    # Get console width for better formatting
    $width = [Math]::Min(120, $Host.UI.RawUI.WindowSize.Width - 2)
    $width = [Math]::Max(40, $width)  # Ensure minimum width
    
    # Safe string creation
    $separator = ""
    $doubleSeparator = ""
    $thinSeparator = ""
    for ($i = 0; $i -lt $width; $i++) {
        $separator += "─"
        $doubleSeparator += "═"
        $thinSeparator += "┄"
    }

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

    function Write-Section {
        param(
            [string]$Title,
            [string]$Icon,
            [string]$Color = "Cyan"
        )
        Write-Host ""
        Write-Host "$Icon $Title" -ForegroundColor $Color
        Write-Host $separator -ForegroundColor DarkGray
    }

    function Write-Parameter {
        param(
            [string]$Name,
            [string]$Aliases,
            [string]$Type,
            [string]$Description,
            [string]$Example,
            [string]$Notes = ""
        )
        Write-Host "    " -NoNewline
        Write-Host $Name -ForegroundColor Green -NoNewline
        if ($Aliases) {
            Write-Host ", " -NoNewline -ForegroundColor DarkGray
            Write-Host $Aliases -ForegroundColor Yellow
        } else {
            Write-Host ""
        }
        Write-Host "        Type: " -NoNewline -ForegroundColor DarkGray
        Write-Host $Type -ForegroundColor Magenta
        Write-Host "        $Description" -ForegroundColor White
        Write-Host "        Example: " -NoNewline -ForegroundColor DarkGray
        Write-Host $Example -ForegroundColor DarkCyan
        if ($Notes) {
            Write-Host "        Note: " -NoNewline -ForegroundColor DarkYellow
            Write-Host $Notes -ForegroundColor Yellow
        }
        Write-Host ""
    }

    Clear-Host
    
    # Header Banner with perfect alignment
    Write-Host "╔$doubleSeparator╗" -ForegroundColor Magenta
    
    # First line with title
    $title1 = "WINGET UPDATE MANAGER v3.1.0"
    $padding1 = $width - $title1.Length
    $leftPad1 = [Math]::Max(0, [Math]::Floor($padding1 / 2))
    $rightPad1 = [Math]::Max(0, $padding1 - $leftPad1)
    Write-Host "║" -NoNewline -ForegroundColor Magenta
    for ($i = 0; $i -lt $leftPad1; $i++) { Write-Host " " -NoNewline }
    Write-Host $title1 -NoNewline -ForegroundColor White
    for ($i = 0; $i -lt $rightPad1; $i++) { Write-Host " " -NoNewline }
    Write-Host "║" -ForegroundColor Magenta
    
    # Second line with subtitle
    $title2 = "Advanced Windows Package Management Solution"
    $padding2 = [Math]::Max(0, $width - $title2.Length)
    $leftPad2 = [Math]::Max(0, [Math]::Floor($padding2 / 2))
    $rightPad2 = [Math]::Max(0, $padding2 - $leftPad2)
    Write-Host "║" -NoNewline -ForegroundColor Magenta
    for ($i = 0; $i -lt $leftPad2; $i++) { Write-Host " " -NoNewline }
    Write-Host $title2 -NoNewline -ForegroundColor DarkCyan
    for ($i = 0; $i -lt $rightPad2; $i++) { Write-Host " " -NoNewline }
    Write-Host "║" -ForegroundColor Magenta
    
    Write-Host "╚$doubleSeparator╝" -ForegroundColor Magenta

    # Synopsis
    Write-Section "SYNOPSIS" "🎯" "Yellow"
    Write-Host "    Professional-grade PowerShell script for automated Windows application management" -ForegroundColor White
    Write-Host "    using Windows Package Manager (winget). Features intelligent application detection," -ForegroundColor White
    Write-Host "    smart exclusion management, multiple operation modes, comprehensive logging," -ForegroundColor White
    Write-Host "    and enterprise-ready automation capabilities." -ForegroundColor White

    # Syntax
    Write-Section "SYNTAX" "📋" "Yellow"
    Write-Host "    Basic Update Operations:" -ForegroundColor Cyan
    Write-Host "        .\winget-update.ps1 [-Mode <String>] [-ExcludeApps <String[]>] [-CustomParams <String>]" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    Persistent Exclusion Management:" -ForegroundColor Cyan
    Write-Host "        .\winget-update.ps1 -AddPersistentExcludeApps <String[]>" -ForegroundColor DarkGray
    Write-Host "        .\winget-update.ps1 -RemovePersistentExcludeApps <String[]>" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    Information & Help:" -ForegroundColor Cyan
    Write-Host "        .\winget-update.ps1 -Help" -ForegroundColor DarkGray

    # Parameters
    Write-Section "PARAMETERS" "⚙️" "Yellow"
    
    Write-Parameter -Name "-Mode" -Aliases "-m" -Type "String" `
        -Description "Specifies the update execution mode. Controls behavior, interaction level, and safety measures." `
        -Example ".\winget-update.ps1 -Mode silent" `
        -Notes "See UPDATE MODES section for detailed mode descriptions"

    Write-Parameter -Name "-ExcludeApps" -Aliases "-e, -exclude" -Type "String[]" `
        -Description "Temporarily excludes specified applications from the current update session only." `
        -Example ".\winget-update.ps1 -ExcludeApps `"Microsoft.Edge,Mozilla.Firefox`"" `
        -Notes "Supports both application names and IDs. Use comma separation for multiple apps."

    Write-Parameter -Name "-AddPersistentExcludeApps" -Aliases "-ape, -add-persistent-exclude" -Type "String[]" `
        -Description "🧠 INTELLIGENT: Adds applications to permanent exclusion list with smart detection." `
        -Example ".\winget-update.ps1 -ape `"Visual Studio Code,Chrome`"" `
        -Notes "Auto-detects installed apps, converts names to IDs, suggests matches if ambiguous"

    Write-Parameter -Name "-RemovePersistentExcludeApps" -Aliases "-rpe, -remove-persistent-exclude" -Type "String[]" `
        -Description "🧠 INTELLIGENT: Removes applications from permanent exclusion list with smart matching." `
        -Example ".\winget-update.ps1 -rpe `"Wireshark,Discord`"" `
        -Notes "Finds matches by name or ID, shows friendly names, suggests alternatives"

    Write-Parameter -Name "-CustomParams" -Aliases "-cp" -Type "String" `
        -Description "Passes additional parameters directly to winget upgrade command." `
        -Example ".\winget-update.ps1 -CustomParams `"--include-unknown --force`"" `
        -Notes "Advanced users only. Refer to winget upgrade --help for available options"

    Write-Parameter -Name "-Help" -Aliases "" -Type "Switch" `
        -Description "Displays this comprehensive help documentation." `
        -Example ".\winget-update.ps1 -Help"

    # Update Modes
    Write-Section "UPDATE MODES" "🔄" "Yellow"
    Write-Host "    The script supports multiple execution modes for different scenarios:" -ForegroundColor White
    Write-Host ""

    $modes = @(
        @{ Name = "normal"; Icon = "🔵"; Description = "Default interactive mode with user prompts and confirmations"; Use = "General use, manual updates" },
        @{ Name = "silent"; Icon = "🔇"; Description = "Automated mode with minimal output, accepts all agreements"; Use = "Scheduled tasks, CI/CD pipelines" },
        @{ Name = "force"; Icon = "⚡"; Description = "Aggressive mode bypassing restrictions and warnings"; Use = "Emergency updates, troubleshooting" },
        @{ Name = "verbose"; Icon = "📝"; Description = "Detailed logging with comprehensive diagnostic information"; Use = "Debugging, issue reporting" },
        @{ Name = "no-interaction"; Icon = "🤖"; Description = "Zero user interaction, fully automated execution"; Use = "Server environments, automation" },
        @{ Name = "full-upgrade"; Icon = "🔄"; Description = "Includes unknown versions and pinned packages"; Use = "Complete system refresh" },
        @{ Name = "safe-upgrade"; Icon = "🛡️"; Description = "Conservative approach with enhanced safety checks"; Use = "Production systems, critical environments" },
        @{ Name = "dry-run"; Icon = "🔍"; Description = "Simulation mode showing planned actions without execution"; Use = "Testing, planning, verification" }
    )

    foreach ($mode in $modes) {
        Write-Host "    $($mode.Icon) " -NoNewline -ForegroundColor White
        Write-Host $mode.Name -NoNewline -ForegroundColor Green
        Write-Host " - " -NoNewline -ForegroundColor DarkGray
        Write-Host $mode.Description -ForegroundColor White
        Write-Host "        Use case: " -NoNewline -ForegroundColor DarkGray
        Write-Host $mode.Use -ForegroundColor DarkCyan
        Write-Host ""
    }

    # Smart Features
    Write-Section "🧠 INTELLIGENT FEATURES" "🤖" "Magenta"
    Write-Host "    Advanced capabilities that make this script enterprise-ready:" -ForegroundColor White
    Write-Host ""
    
    Write-Host "    🔍 Smart Application Detection" -ForegroundColor Green
    Write-Host "        • Automatically scans installed applications" -ForegroundColor DarkGray
    Write-Host "        • Matches application names to official IDs" -ForegroundColor DarkGray
    Write-Host "        • Provides suggestions for partial matches" -ForegroundColor DarkGray
    Write-Host ""
    
    Write-Host "    🎯 Intelligent Exclusion Management" -ForegroundColor Green
    Write-Host "        • Converts friendly names to precise IDs automatically" -ForegroundColor DarkGray
    Write-Host "        • Handles ambiguous matches with user-friendly suggestions" -ForegroundColor DarkGray
    Write-Host "        • Cross-references with installed applications database" -ForegroundColor DarkGray
    Write-Host ""
    
    Write-Host "    � Professeional Reporting" -ForegroundColor Green
    Write-Host "        • Comprehensive update summaries with statistics" -ForegroundColor DarkGray
    Write-Host "        • Detailed logging with timestamps and session IDs" -ForegroundColor DarkGray
    Write-Host "        • Color-coded status indicators and progress tracking" -ForegroundColor DarkGray
    Write-Host ""

    # Examples
    Write-Section "EXAMPLES" "📚" "Yellow"
    
    $examples = @(
        @{ 
            Title = "Basic Operations"
            Commands = @(
                @{ Cmd = ".\winget-update.ps1"; Desc = "Standard update with default settings" },
                @{ Cmd = ".\winget-update.ps1 -Mode dry-run"; Desc = "Preview updates without installing" },
                @{ Cmd = ".\winget-update.ps1 -Mode silent"; Desc = "Automated silent update" }
            )
        },
        @{ 
            Title = "Exclusion Management"
            Commands = @(
                @{ Cmd = ".\winget-update.ps1 -e `"Microsoft.Edge,Discord`""; Desc = "Exclude apps from current session" },
                @{ Cmd = ".\winget-update.ps1 -ape `"Visual Studio Code`""; Desc = "Add app to permanent exclusion (smart detection)" },
                @{ Cmd = ".\winget-update.ps1 -rpe `"Chrome`""; Desc = "Remove app from permanent exclusion (smart matching)" }
            )
        },
        @{ 
            Title = "Advanced Usage"
            Commands = @(
                @{ Cmd = ".\winget-update.ps1 -Mode force -cp `"--include-unknown`""; Desc = "Force update including unknown versions" },
                @{ Cmd = ".\winget-update.ps1 -Mode verbose -e `"App1,App2`""; Desc = "Detailed logging with specific exclusions" },
                @{ Cmd = ".\winget-update.ps1 -Mode no-interaction"; Desc = "Fully automated execution for servers" }
            )
        },
        @{ 
            Title = "Smart Exclusion Examples"
            Commands = @(
                @{ Cmd = ".\winget-update.ps1 -ape `"Wireshark`""; Desc = "Finds: WiresharkFoundation.Wireshark automatically" },
                @{ Cmd = ".\winget-update.ps1 -ape `"Chrome`""; Desc = "Shows multiple matches: Google.Chrome, etc." },
                @{ Cmd = ".\winget-update.ps1 -rpe `"Visual Studio`""; Desc = "Smart removal with friendly name display" }
            )
        }
    )

    foreach ($exampleGroup in $examples) {
        Write-Host "    $($exampleGroup.Title):" -ForegroundColor Cyan
        foreach ($example in $exampleGroup.Commands) {
            Write-Host "        📎 " -NoNewline -ForegroundColor DarkGray
            Write-Host $example.Cmd -ForegroundColor DarkCyan
            Write-Host "           $($example.Desc)" -ForegroundColor White
        }
        Write-Host ""
    }

    # File Structure
    Write-Section "FILE STRUCTURE" "📁" "Yellow"
    Write-Host "    The script creates and manages the following files and directories:" -ForegroundColor White
    Write-Host ""
    Write-Host "    📄 persistent_exclude_apps.txt" -ForegroundColor Green
    Write-Host "        Contains permanently excluded application IDs" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    📁 logs/" -ForegroundColor Green
    Write-Host "        ├── winget_update_YYYY-MM-DD_HH-MM-SS.log    (Standard logs)" -ForegroundColor DarkGray
    Write-Host "        └── winget_update_YYYY-MM-DD_HH-MM-SS_ID.log (Session-specific logs)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    📁 backup/" -ForegroundColor Green
    Write-Host "        Automatic backups of configuration files" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    📁 logs/" -ForegroundColor Green
    Write-Host "        └── winget_update_*.log   (Logs des mises à jour)" -ForegroundColor DarkGray

    # Requirements
    Write-Section "REQUIREMENTS" "🔧" "Yellow"
    Write-Host "    System Requirements:" -ForegroundColor Cyan
    Write-Host "        • Windows 10 version 1809 (build 17763) or later" -ForegroundColor White
    Write-Host "        • PowerShell 5.1 or PowerShell 7+" -ForegroundColor White
    Write-Host "        • Windows Package Manager (winget) installed and configured" -ForegroundColor White
    Write-Host "        • Administrator privileges (recommended for system-wide updates)" -ForegroundColor White
    Write-Host ""
    Write-Host "    Optional Components:" -ForegroundColor Cyan
    Write-Host "        • Windows Terminal (for enhanced display)" -ForegroundColor DarkGray
    Write-Host "        • Git (for version checking and updates)" -ForegroundColor DarkGray

    # Troubleshooting
    Write-Section "TROUBLESHOOTING" "🔧" "Yellow"
    Write-Host "    Common Issues and Solutions:" -ForegroundColor White
    Write-Host ""
    Write-Host "    🚫 'winget' is not recognized" -ForegroundColor Red
    Write-Host "        Solution: Install Windows Package Manager from Microsoft Store" -ForegroundColor Green
    Write-Host ""
    Write-Host "    🚫 Access denied errors" -ForegroundColor Red
    Write-Host "        Solution: Run PowerShell as Administrator" -ForegroundColor Green
    Write-Host ""
    Write-Host "    🚫 Script execution policy errors" -ForegroundColor Red
    Write-Host "        Solution: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor Green
    Write-Host ""
    Write-Host "    🚫 Application not found in smart detection" -ForegroundColor Red
    Write-Host "        Solution: Ensure application is installed via winget, try exact ID instead" -ForegroundColor Green

    # Footer with perfect alignment
    Write-Host ""
    Write-Host "╔$doubleSeparator╗" -ForegroundColor DarkGray
    
    # System Information header
    $sysInfoTitle = "SYSTEM INFORMATION"
    $sysInfoPadding = [Math]::Max(0, $width - $sysInfoTitle.Length)
    $sysInfoLeftPad = [Math]::Max(0, [Math]::Floor($sysInfoPadding / 2))
    $sysInfoRightPad = [Math]::Max(0, $sysInfoPadding - $sysInfoLeftPad)
    Write-Host "║" -NoNewline -ForegroundColor DarkGray
    for ($i = 0; $i -lt $sysInfoLeftPad; $i++) { Write-Host " " -NoNewline }
    Write-Host $sysInfoTitle -NoNewline -ForegroundColor White
    for ($i = 0; $i -lt $sysInfoRightPad; $i++) { Write-Host " " -NoNewline }
    Write-Host "║" -ForegroundColor DarkGray
    
    # Version line
    $versionText = "    Version: v3.1.0"
    $versionPadding = [Math]::Max(0, $width - $versionText.Length)
    Write-Host "║" -NoNewline -ForegroundColor DarkGray
    Write-Host "    Version: " -NoNewline -ForegroundColor DarkGray
    Write-Host "v3.1.0" -NoNewline -ForegroundColor Green
    for ($i = 0; $i -lt $versionPadding; $i++) { Write-Host " " -NoNewline }
    Write-Host "║" -ForegroundColor DarkGray
    
    # Author line
    $authorText = "    Author: Sterbweise"
    $authorPadding = [math]::Max(0, $width - $authorText.Length)
    Write-Host "║" -NoNewline -ForegroundColor DarkGray
    Write-Host "    Author: " -NoNewline -ForegroundColor DarkGray
    Write-Host "Sterbweise" -NoNewline -ForegroundColor Yellow
    if ($authorPadding -gt 0) {
        for ($i = 0; $i -lt $authorPadding; $i++) { Write-Host " " -NoNewline }
    }
    Write-Host "║" -ForegroundColor DarkGray
    
    # License line
    $licenseText = "    License: MIT"
    $licensePadding = [math]::Max(0, $width - $licenseText.Length)
    Write-Host "║" -NoNewline -ForegroundColor DarkGray
    Write-Host "    License: " -NoNewline -ForegroundColor DarkGray
    Write-Host "MIT" -NoNewline -ForegroundColor Green
    if ($licensePadding -gt 0) {
        for ($i = 0; $i -lt $licensePadding; $i++) { Write-Host " " -NoNewline }
    }
    Write-Host "║" -ForegroundColor DarkGray
    
    # Repository line
    $repoLine = "    Repository: https://github.com/Sterbweise/winget-update"
    $repoPadding = [math]::Max(0, $width - $repoLine.Length)
    Write-Host "║" -NoNewline -ForegroundColor DarkGray
    Write-Host "    Repository: " -NoNewline -ForegroundColor DarkGray
    Write-Host "https://github.com/Sterbweise/winget-update" -NoNewline -ForegroundColor Blue
    if ($repoPadding -gt 0) {
        for ($i = 0; $i -lt $repoPadding; $i++) { Write-Host " " -NoNewline }
    }
    Write-Host "║" -ForegroundColor DarkGray
    
    # Last Updated line
    $updateLine = "    Last Updated: July 20, 2025"
    $updatePadding = [math]::Max(0, $width - $updateLine.Length)
    Write-Host "║" -NoNewline -ForegroundColor DarkGray
    Write-Host "    Last Updated: " -NoNewline -ForegroundColor DarkGray
    Write-Host "July 20, 2025" -NoNewline -ForegroundColor Magenta
    if ($updatePadding -gt 0) {
        for ($i = 0; $i -lt $updatePadding; $i++) { Write-Host " " -NoNewline }
    }
    Write-Host "║" -ForegroundColor DarkGray
    
    Write-Host "╚$doubleSeparator╝" -ForegroundColor DarkGray
    
    Write-Host ""
    Write-CenteredText "Press any key to exit help and return to terminal..." "DarkYellow"
    Write-Host ""
    
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
        $content = Get-Content $PersistentExcludeFile -ErrorAction SilentlyContinue
        if ($content) {
            return @($content | Where-Object { $_.Trim() -ne "" })
        }
    }
    return @()
}

# Function to get installed applications for smart suggestions
function Get-InstalledApplications {
    try {
        $installedApps = @()
        
        # Get winget list output
        $wingetOutput = winget list --accept-source-agreements 2>$null
        
        if ($wingetOutput) {
            # Parse winget output (skip header lines)
            $lines = $wingetOutput | Select-Object -Skip 2
            
            foreach ($line in $lines) {
                if ($line -and $line.Trim() -ne "" -and $line -notmatch "^-+") {
                    # Parse the line to extract Name and Id
                    $parts = $line -split '\s{2,}' # Split on multiple spaces
                    if ($parts.Count -ge 2) {
                        $name = $parts[0].Trim()
                        $id = $parts[1].Trim()
                        
                        if ($name -and $id -and $id -notmatch "^Available|^Upgrades") {
                            $installedApps += [PSCustomObject]@{
                                Name = $name
                                Id = $id
                            }
                        }
                    }
                }
            }
        }
        
        return $installedApps
    } catch {
        Write-Host "⚠️  Warning: Could not retrieve installed applications list." -ForegroundColor Yellow
        return @()
    }
}

# Function to find application suggestions
function Find-ApplicationSuggestions {
    param(
        [string]$SearchTerm,
        [array]$InstalledApps
    )
    
    $suggestions = @()
    
    # Ensure InstalledApps is an array
    if (-not $InstalledApps) {
        return @()
    }
    $InstalledApps = @($InstalledApps)
    
    # Exact name match
    $exactNameMatch = $InstalledApps | Where-Object { $_.Name -eq $SearchTerm }
    if ($exactNameMatch) {
        return @($exactNameMatch)
    }
    
    # Exact ID match
    $exactIdMatch = $InstalledApps | Where-Object { $_.Id -eq $SearchTerm }
    if ($exactIdMatch) {
        return @($exactIdMatch)
    }
    
    # Partial name matches (case insensitive)
    $nameMatches = @($InstalledApps | Where-Object { $_.Name -like "*$SearchTerm*" })
    $suggestions += $nameMatches
    
    # Partial ID matches (case insensitive)
    $idMatches = @($InstalledApps | Where-Object { $_.Id -like "*$SearchTerm*" })
    $suggestions += $idMatches
    
    # Remove duplicates
    $uniqueSuggestions = @()
    foreach ($suggestion in $suggestions) {
        $exists = $false
        foreach ($unique in $uniqueSuggestions) {
            if ($unique.Id -eq $suggestion.Id) {
                $exists = $true
                break
            }
        }
        if (-not $exists) {
            $uniqueSuggestions += $suggestion
        }
    }
    
    return $uniqueSuggestions | Select-Object -First 10 # Limit to 10 suggestions
}

# Enhanced function to add applications to the persistent exclude list with smart suggestions
function Add-PersistentExcludeApps {
    param(
        [string[]]$Apps
    )
    
    if (-not $Apps -or $Apps.Count -eq 0) {
        Write-Host "❌ No applications specified to add to the persistent exclude list." -ForegroundColor Red
        return
    }
    
    Write-Host "🔍 Analyzing installed applications..." -ForegroundColor Cyan
    $installedApps = @(Get-InstalledApplications)
    
    if (-not $installedApps -or $installedApps.Count -eq 0) {
        Write-Host "⚠️  Warning: Could not retrieve installed applications list." -ForegroundColor Yellow
        Write-Host "   💡 Make sure winget is working properly: winget list" -ForegroundColor DarkGray
        return
    }
    
    $currentList = @(Get-PersistentExcludeList)
    $validAppsToAdd = @()
    
    foreach ($app in $Apps) {
        if ($currentList -contains $app) {
            Write-Host "ℹ️  '$app' is already in the persistent exclude list." -ForegroundColor Yellow
            continue
        }
        
        # Find suggestions for this app
        $suggestions = @(Find-ApplicationSuggestions -SearchTerm $app -InstalledApps $installedApps)
        
        if (-not $suggestions -or $suggestions.Count -eq 0) {
            Write-Host "❌ No installed application found matching '$app'." -ForegroundColor Red
            Write-Host "   💡 Make sure the application is installed and try using the exact name or ID." -ForegroundColor DarkGray
            continue
        }
        
        if ($suggestions -and $suggestions.Count -eq 1) {
            # Exact match found - use the ID
            $selectedApp = $suggestions[0]
            $appToAdd = $selectedApp.Id
            
            Write-Host "✅ Found exact match for '$app':" -ForegroundColor Green
            Write-Host "   📦 Name: $($selectedApp.Name)" -ForegroundColor White
            Write-Host "   🆔 ID: $($selectedApp.Id)" -ForegroundColor Cyan
            Write-Host "   ➕ Adding ID to exclude list..." -ForegroundColor DarkGray
            
            $validAppsToAdd += $appToAdd
        } else {
            # Multiple matches found - show suggestions
            Write-Host "🔍 Multiple applications found matching '$app':" -ForegroundColor Yellow
            Write-Host "   Please specify which one you want to exclude:" -ForegroundColor DarkGray
            Write-Host ""
            
            for ($i = 0; $i -lt $suggestions.Count; $i++) {
                $suggestion = $suggestions[$i]
                Write-Host "   [$($i + 1)] " -NoNewline -ForegroundColor Cyan
                Write-Host "$($suggestion.Name)" -NoNewline -ForegroundColor White
                Write-Host " (" -NoNewline -ForegroundColor DarkGray
                Write-Host "$($suggestion.Id)" -NoNewline -ForegroundColor Yellow
                Write-Host ")" -ForegroundColor DarkGray
            }
            
            Write-Host ""
            Write-Host "   💡 To exclude a specific app, use:" -ForegroundColor DarkGray
            Write-Host "      .\winget-update.ps1 -ape `"$($suggestions[0].Id)`"" -ForegroundColor Green
            Write-Host ""
        }
    }
    
    # Add valid apps to the list
    if ($validAppsToAdd.Count -gt 0) {
        $allApps = @($currentList) + @($validAppsToAdd)
        $allApps | Where-Object { $_.Trim() -ne "" } | Set-Content $PersistentExcludeFile
        
        Write-Host "✅ Successfully added to persistent exclude list:" -ForegroundColor Green
        foreach ($app in $validAppsToAdd) {
            Write-Host "   • $app" -ForegroundColor DarkGray
        }
    }
}

# Function to find matches in the persistent exclude list
function Find-ExcludeListMatches {
    param(
        [string]$SearchTerm,
        [array]$ExcludeList,
        [array]$InstalledApps
    )
    
    $matches = @()
    
    # Exact match in exclude list
    $exactMatch = $ExcludeList | Where-Object { $_ -eq $SearchTerm }
    if ($exactMatch) {
        return @($exactMatch)
    }
    
    # Find the app in installed apps to get both name and ID
    $installedApp = $InstalledApps | Where-Object { $_.Name -eq $SearchTerm -or $_.Id -eq $SearchTerm }
    if ($installedApp) {
        # Check if the ID is in the exclude list
        $idMatch = $ExcludeList | Where-Object { $_ -eq $installedApp.Id }
        if ($idMatch) {
            return @($idMatch)
        }
        
        # Check if the name is in the exclude list
        $nameMatch = $ExcludeList | Where-Object { $_ -eq $installedApp.Name }
        if ($nameMatch) {
            return @($nameMatch)
        }
    }
    
    # Partial matches in exclude list (case insensitive)
    $partialMatches = $ExcludeList | Where-Object { $_ -like "*$SearchTerm*" }
    $matches += $partialMatches
    
    # Find apps in installed list that match the search term and check if their IDs are in exclude list
    $installedMatches = $InstalledApps | Where-Object { $_.Name -like "*$SearchTerm*" -or $_.Id -like "*$SearchTerm*" }
    foreach ($installedMatch in $installedMatches) {
        $excludeMatch = $ExcludeList | Where-Object { $_ -eq $installedMatch.Id -or $_ -eq $installedMatch.Name }
        if ($excludeMatch) {
            $matches += $excludeMatch
        }
    }
    
    # Remove duplicates
    $uniqueMatches = @()
    foreach ($match in $matches) {
        if ($uniqueMatches -notcontains $match) {
            $uniqueMatches += $match
        }
    }
    
    return $uniqueMatches | Select-Object -First 10
}

# Enhanced function to remove applications from the persistent exclude list with smart matching
function Remove-PersistentExcludeApps {
    param(
        [string[]]$Apps
    )
    
    if (-not $Apps -or $Apps.Count -eq 0) {
        Write-Host "❌ No applications specified to remove from the persistent exclude list." -ForegroundColor Red
        return
    }
    
    Write-Host "🔍 Analyzing exclude list and installed applications..." -ForegroundColor Cyan
    $currentList = @(Get-PersistentExcludeList)
    $installedApps = Get-InstalledApplications
    
    $validAppsToRemove = @()
    
    foreach ($app in $Apps) {
        # Find matches for this app in the exclude list
        $matches = @(Find-ExcludeListMatches -SearchTerm $app -ExcludeList $currentList -InstalledApps $installedApps)
        
        if ($matches.Count -eq 0) {
            Write-Host "❌ No application matching '$app' found in the persistent exclude list." -ForegroundColor Red
            
            # Show what's currently in the exclude list for reference
            if ($currentList.Count -gt 0) {
                Write-Host "   📋 Current exclude list contains:" -ForegroundColor DarkGray
                foreach ($excludedApp in $currentList) {
                    # Try to find the friendly name for this ID
                    $friendlyApp = $installedApps | Where-Object { $_.Id -eq $excludedApp }
                    if ($friendlyApp) {
                        Write-Host "      • $($friendlyApp.Name) ($excludedApp)" -ForegroundColor DarkGray
                    } else {
                        Write-Host "      • $excludedApp" -ForegroundColor DarkGray
                    }
                }
            }
            continue
        }
        
        if ($matches.Count -eq 1) {
            # Exact match found
            $appToRemove = $matches[0]
            
            # Find friendly name if possible
            $friendlyApp = $installedApps | Where-Object { $_.Id -eq $appToRemove -or $_.Name -eq $appToRemove }
            
            Write-Host "✅ Found match for '$app' in exclude list:" -ForegroundColor Green
            if ($friendlyApp) {
                Write-Host "   📦 Name: $($friendlyApp.Name)" -ForegroundColor White
                Write-Host "   🆔 ID: $($friendlyApp.Id)" -ForegroundColor Cyan
            } else {
                Write-Host "   🆔 Entry: $appToRemove" -ForegroundColor Cyan
            }
            Write-Host "   ➖ Removing from exclude list..." -ForegroundColor DarkGray
            
            $validAppsToRemove += $appToRemove
        } else {
            # Multiple matches found - show options
            Write-Host "🔍 Multiple entries found in exclude list matching '$app':" -ForegroundColor Yellow
            Write-Host "   Please specify which one you want to remove:" -ForegroundColor DarkGray
            Write-Host ""
            
            for ($i = 0; $i -lt $matches.Count; $i++) {
                $match = $matches[$i]
                $friendlyApp = $installedApps | Where-Object { $_.Id -eq $match -or $_.Name -eq $match }
                
                Write-Host "   [$($i + 1)] " -NoNewline -ForegroundColor Cyan
                if ($friendlyApp) {
                    Write-Host "$($friendlyApp.Name)" -NoNewline -ForegroundColor White
                    Write-Host " (" -NoNewline -ForegroundColor DarkGray
                    Write-Host "$match" -NoNewline -ForegroundColor Yellow
                    Write-Host ")" -ForegroundColor DarkGray
                } else {
                    Write-Host "$match" -ForegroundColor Yellow
                }
            }
            
            Write-Host ""
            Write-Host "   💡 To remove a specific app, use:" -ForegroundColor DarkGray
            Write-Host "      .\winget-update.ps1 -rpe `"$($matches[0])`"" -ForegroundColor Green
            Write-Host ""
        }
    }
    
    # Remove valid apps from the list
    if ($validAppsToRemove.Count -gt 0) {
        $newList = @($currentList | Where-Object { $validAppsToRemove -notcontains $_ })
        $newList | Where-Object { $_.Trim() -ne "" } | Set-Content $PersistentExcludeFile
        
        Write-Host "✅ Successfully removed from persistent exclude list:" -ForegroundColor Green
        foreach ($app in $validAppsToRemove) {
            # Try to show friendly name
            $friendlyApp = $installedApps | Where-Object { $_.Id -eq $app -or $_.Name -eq $app }
            if ($friendlyApp) {
                Write-Host "   • $($friendlyApp.Name) ($app)" -ForegroundColor DarkGray
            } else {
                Write-Host "   • $app" -ForegroundColor DarkGray
            }
        }
    }
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

# Retrieve the current persistent exclude list and ensure it's an array
$CurrentPersistentExcludeApps = @(Get-PersistentExcludeList)

# Function to escape special regex characters
function Escape-SpecialCharacters {
    param(
        [string]$String
    )
    return [regex]::Escape($String)
}

# Function to get the list of available updates from winget
function Get-WingetUpdates {
    Write-Host "🔍 Scanning for available updates..." -ForegroundColor Cyan
    Write-Host "DEBUG: Get-WingetUpdates called" -ForegroundColor Magenta
    
    try {
        # Execute winget with enhanced error handling
        $wingetOutput = & winget upgrade --include-unknown 2>&1
        
        if ($LASTEXITCODE -ne 0) {
            Write-Host "❌ Winget command failed with exit code: $LASTEXITCODE" -ForegroundColor Red
            return @()
        }
        
        $lines = $wingetOutput -split "`r`n"
        $updates = @()
        $headerFound = $false
        $startProcessing = $false
        
        foreach ($line in $lines) {
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
                    if ($id -ne "Id" -and $version -ne "Version" -and -not ($name -match "following packages|explicit targeting")) {
                        $updateInfo = [PSCustomObject]@{
                            Name = $name
                            Id = $id
                            Version = $version
                            Available = $available
                            Source = $source
                            Priority = Get-UpdatePriority -Id $id
                            Category = Get-PackageCategory -Id $id
                            SecurityUpdate = Test-SecurityUpdate -Id $id
                        }
                        
                        $updates += $updateInfo
                    }
                }
            }
        }

        # Ensure $updates is always an array and remove duplicates
        $updates = @($updates)
        Write-Host "DEBUG: Raw updates count: $($updates.Count)" -ForegroundColor Magenta
        $updates | Where-Object { $_.Name -match "v2ray|2dust" -or $_.Id -match "v2ray|2dust" } | ForEach-Object {
            Write-Host "DEBUG: Found v2rayN: $($_.Name) ($($_.Id))" -ForegroundColor Magenta
        }
        
        $uniqueUpdates = @($updates | Sort-Object Id -Unique | Sort-Object Priority, Name)
        
        Write-Host "✅ Found $($uniqueUpdates.Count) available updates" -ForegroundColor Green
        return $uniqueUpdates
        
    } catch {
        Write-Host "❌ Error scanning for updates: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Function to determine update priority
function Get-UpdatePriority {
    param([string]$Id)
    
    # Define priority categories (1 = highest, 3 = lowest)
    
    # High Priority: Critical development tools, security software, system utilities
    $highPriority = @(
        # Development Tools
        "Microsoft.WindowsTerminal", "Microsoft.PowerShell", "Microsoft.VisualStudioCode", 
        "Microsoft.VisualStudio.2022.Community", "Microsoft.VisualStudio.2022.Professional", "Microsoft.VisualStudio.2022.Enterprise",
        "Microsoft.VisualStudio.2022.BuildTools", "Git.Git", "GitHub.GitHubDesktop", "GitKraken.GitKraken",
        "JetBrains.IntelliJIDEA.Ultimate", "JetBrains.IntelliJIDEA.Community", "JetBrains.PyCharm.Professional", "JetBrains.PyCharm.Community",
        "JetBrains.WebStorm", "JetBrains.PhpStorm", "JetBrains.CLion", "JetBrains.DataGrip", "JetBrains.Rider",
        "Docker.DockerDesktop", "Kubernetes.kubectl", "Microsoft.AzureCLI", "Amazon.AWSCLI", "Google.CloudSDK",
        
        # Security & System
        "Microsoft.WindowsDefender", "Malwarebytes.Malwarebytes", "ESET.EndpointSecurity", "Bitdefender.Bitdefender",
        "Microsoft.Sysinternals.ProcessExplorer", "Microsoft.Sysinternals.ProcessMonitor", "Microsoft.Sysinternals.Autoruns",
        
        # Critical Runtimes
        "Microsoft.DotNet.Runtime.6", "Microsoft.DotNet.Runtime.7", "Microsoft.DotNet.Runtime.8",
        "Microsoft.VCRedist.2015+.x64", "Microsoft.VCRedist.2015+.x86", "Oracle.JavaRuntimeEnvironment",
        "Python.Python.3.11", "Python.Python.3.12", "NodeJS.NodeJS"
    )
    
    # Medium Priority: Popular applications, browsers, productivity tools
    $mediumPriority = @(
        # Browsers
        "Google.Chrome", "Mozilla.Firefox", "Microsoft.Edge", "Opera.Opera", "BraveSoftware.BraveBrowser",
        "Vivaldi.Vivaldi", "Tor.TorBrowser",
        
        # Productivity & Office
        "Microsoft.Office", "Microsoft.Teams", "Zoom.Zoom", "Slack.Slack", "Discord.Discord",
        "Notion.Notion", "Obsidian.Obsidian", "Evernote.Evernote", "Dropbox.Dropbox", "Google.Drive",
        "Adobe.Acrobat.Reader.64-bit", "SumatraPDF.SumatraPDF", "Foxit.FoxitReader",
        
        # Media & Graphics
        "Adobe.Photoshop", "Adobe.Illustrator", "Adobe.Premiere", "Adobe.AfterEffects",
        "GIMP.GIMP", "Inkscape.Inkscape", "Blender.Blender", "OBSProject.OBSStudio",
        "VLC.VLC", "MPC-HC.MPC-HC", "Spotify.Spotify", "Audacity.Audacity",
        
        # Utilities
        "7zip.7zip", "WinRAR.WinRAR", "PeaZip.PeaZip", "Notepad++.Notepad++", "Sublime.SublimeText.4",
        "Microsoft.PowerToys", "Sysinternals.ProcessExplorer", "WiresharkFoundation.Wireshark",
        
        # Cloud & Sync
        "Microsoft.OneDrive", "Google.BackupAndSync", "Amazon.Kindle", "Evernote.Evernote"
    )
    
    if ($Id -in $highPriority) { return 1 }
    if ($Id -in $mediumPriority) { return 2 }
    
    # Additional pattern-based priorities
    if ($Id -match "Microsoft\.(VisualStudio|DotNet|PowerShell|WindowsTerminal)") { return 1 }
    if ($Id -match "JetBrains\.|Docker\.|Kubernetes\.|Git\.") { return 1 }
    if ($Id -match "Microsoft\.|Google\.|Adobe\.|Mozilla\.") { return 2 }
    if ($Id -match "Antivirus|Security|Defender") { return 1 }
    
    return 3
}

# Function to categorize packages
function Get-PackageCategory {
    param([string]$Id)
    
    switch -Regex ($Id) {
        # Development Tools
        "Microsoft\.(VisualStudio|PowerShell|WindowsTerminal|DotNet|AzureCLI)" { return "Development" }
        "JetBrains\.|IntelliJ|PyCharm|WebStorm|PhpStorm|CLion|DataGrip|Rider" { return "Development" }
        "Git\.|GitHub\.|GitKraken|Sourcetree" { return "Development" }
        "Python\.|NodeJS\.|Go\.|Rust\.|Java\.|Oracle\.Java" { return "Development" }
        "Docker\.|Kubernetes\.|Postman\.|Insomnia" { return "Development" }
        "Notepad\+\+|Sublime|Atom|Brackets|VSCodium" { return "Development" }
        "Amazon\.(AWSCLI|SAM)|Google\.CloudSDK|Terraform" { return "Development" }
        
        # Browsers
        "Google\.Chrome|Mozilla\.Firefox|Microsoft\.Edge|Opera|Brave|Vivaldi|Tor" { return "Browser" }
        
        # Microsoft Products
        "Microsoft\.(Office|Teams|OneDrive|Skype|PowerToys)" { return "Microsoft" }
        "Microsoft\.(VCRedist|DirectX|XNAFramework)" { return "Microsoft Runtime" }
        
        # Google Products
        "Google\.(Chrome|Drive|Earth|BackupAndSync)" { return "Google" }
        
        # Adobe Products
        "Adobe\.(Photoshop|Illustrator|Premiere|AfterEffects|Acrobat|Reader)" { return "Adobe" }
        
        # Media & Entertainment
        "VLC|MPC-HC|PotPlayer|Kodi|Plex" { return "Media Player" }
        "Spotify|iTunes|Audacity|Foobar2000|AIMP" { return "Audio" }
        "OBS|Streamlabs|XSplit|Bandicam|Camtasia" { return "Recording" }
        "GIMP|Inkscape|Paint\.NET|Krita|Canva" { return "Graphics" }
        "Blender|Unity|UnrealEngine|Autodesk" { return "3D/CAD" }
        
        # Gaming
        "Steam|Epic|GOG|Origin|Ubisoft|Rockstar|Battle\.net" { return "Gaming Platform" }
        "Discord|TeamSpeak|Mumble|Ventrilo" { return "Gaming Communication" }
        
        # Productivity & Office
        "Zoom|Skype|Teams|Slack|WhatsApp|Telegram" { return "Communication" }
        "Notion|Obsidian|Evernote|OneNote|Joplin" { return "Note Taking" }
        "Dropbox|OneDrive|GoogleDrive|Box|Mega" { return "Cloud Storage" }
        "WinRAR|7zip|PeaZip|Bandizip|WinZip" { return "Archive" }
        
        # System & Utilities
        "CCleaner|Malwarebytes|Avast|AVG|Kaspersky|ESET|Bitdefender" { return "Security" }
        "CPU-Z|GPU-Z|HWiNFO|Speccy|CrystalDiskInfo" { return "System Info" }
        "Wireshark|Fiddler|Nmap|PuTTY|WinSCP" { return "Network Tools" }
        "VirtualBox|VMware|Hyper-V|QEMU" { return "Virtualization" }
        "Sysinternals|ProcessExplorer|ProcessMonitor|Autoruns" { return "System Tools" }
        
        # Runtimes & Libraries
        "Runtime|Redist|Framework|Library" { return "Runtime" }
        
        # Specific Applications
        "Kindle|Calibre|SumatraPDF|Foxit" { return "Reading" }
        "FileZilla|WinSCP|Cyberduck|Core FTP" { return "FTP Client" }
        "Thunderbird|Outlook|MailBird|eM Client" { return "Email Client" }
        "Tor|ProtonVPN|NordVPN|ExpressVPN" { return "VPN/Privacy" }
        
        # Fallback patterns
        "Microsoft\." { return "Microsoft" }
        "Google\." { return "Google" }
        "Mozilla\." { return "Mozilla" }
        "Adobe\." { return "Adobe" }
        "Apple\." { return "Apple" }
        "Amazon\." { return "Amazon" }
        
        default { return "Other" }
    }
}

# Function to check if update is security-related
function Test-SecurityUpdate {
    param([string]$Id)
    
    # Simple heuristic - in production, this would check against security databases
    $securityApps = @("Microsoft.Edge", "Google.Chrome", "Mozilla.Firefox", "Adobe.Acrobat.Reader.64-bit")
    return ($Id -in $securityApps)
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
    
    # Ensure all arrays are properly initialized to prevent Count errors
    $Updates = @($Updates)
    $ExcludeApps = @($ExcludeApps)
    $PersistentExcludeApps = @($PersistentExcludeApps)
    
    # Enhanced header with system info
    $systemInfo = Get-SystemInfo
    $currentVersion = "v3.1.0"
    
    # Check for script updates
    try {
        $latestRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/Sterbweise/winget-update/releases/latest" -TimeoutSec 5
        $latestVersion = $latestRelease.tag_name
        if ($latestVersion -ne $currentVersion) {
            $updateUrl = "https://github.com/Sterbweise/winget-update/releases/latest"
            $toolVersionStatus = "$currentVersion → $latestVersion"
            $toolVersionColors = @("Red", "White", "Green")  # Old version, Arrow, New version
            Write-Host "🔔 A new version is available! " -ForegroundColor Yellow -NoNewline
            Write-Host "($updateUrl)" -ForegroundColor Cyan
        } else {
            $toolVersionStatus = "$currentVersion (Latest)"
            $toolVersionColors = @("Green")
        }
    } catch {
        $toolVersionStatus = "$currentVersion (Unable to check)"
        $toolVersionColors = @("Yellow")
    }

    # AUTOMATIC WIDTH CALCULATION - Fits any terminal
    $consoleWidth = try { $Host.UI.RawUI.WindowSize.Width } catch { 100 }
    $width = [Math]::Min(100, $consoleWidth - 2)  # Leave 2 chars margin
    # Safe border creation
    $borderWidth = [Math]::Max(0, $width - 2)
    $border = ""
    $thinBorder = ""
    for ($i = 0; $i -lt $borderWidth; $i++) {
        $border += "═"
        $thinBorder += "─"
    }
    
    # Main header
    Write-Host "`n╔$border╗" -ForegroundColor DarkGray
    Write-Host "║" -NoNewline -ForegroundColor DarkGray
    $title = "WINGET UPDATE MANAGER"
    $padding = " " * [Math]::Max(0, [Math]::Floor(($width - 2 - $title.Length) / 2))
    Write-Host "$padding$title$padding" -NoNewline -ForegroundColor White
    if (($padding.Length * 2 + $title.Length) -lt ($width - 2)) {
        Write-Host " " -NoNewline
    }
    Write-Host "║" -ForegroundColor DarkGray
    Write-Host "╠$border╣" -ForegroundColor DarkGray

    # System information
    Write-Host "║ " -NoNewline -ForegroundColor DarkGray
    Write-Host "💻 System: " -NoNewline -ForegroundColor White
    Write-Host "$($systemInfo.ComputerName) - $($systemInfo.OSVersion)" -NoNewline -ForegroundColor Gray
    Write-Host (" " * ($width - 17 - $systemInfo.ComputerName.Length - $systemInfo.OSVersion.Length)) -NoNewline
    Write-Host "║" -ForegroundColor DarkGray
    
    Write-Host "║ " -NoNewline -ForegroundColor DarkGray
    Write-Host "📅 Date: " -NoNewline -ForegroundColor White
    Write-Host (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -NoNewline -ForegroundColor Gray
    Write-Host (" " * ($width - 31)) -NoNewline
    Write-Host "║" -ForegroundColor DarkGray
    
    Write-Host "║ " -NoNewline -ForegroundColor DarkGray
    Write-Host "🔧 Script Version: " -NoNewline -ForegroundColor White
    
    # Display version with separate colors if it's an update
    if ($toolVersionColors.Count -eq 3) {
        # Split the version string to display with different colors
        $versionParts = $toolVersionStatus -split " → "
        Write-Host $versionParts[0] -NoNewline -ForegroundColor $toolVersionColors[0]  # Old version (Red)
        Write-Host " → " -NoNewline -ForegroundColor $toolVersionColors[1]              # Arrow (White)
        Write-Host $versionParts[1] -NoNewline -ForegroundColor $toolVersionColors[2]  # New version (Green)
    } else {
        # Single color for current version
        Write-Host $toolVersionStatus -NoNewline -ForegroundColor $toolVersionColors[0]
    }
    
    Write-Host (" " * ($width - 22 - $toolVersionStatus.Length)) -NoNewline
    Write-Host "║" -ForegroundColor DarkGray

    Write-Host "╠$thinBorder╣" -ForegroundColor DarkGray

    # Update statistics
    $totalUpdates = $Updates.Count
    $excludedCount = @($Updates | Where-Object { ($ExcludeApps -contains $_.Id) -or ($PersistentExcludeApps -contains $_.Id) }).Count
    $activeUpdates = $totalUpdates - $excludedCount
    $securityUpdates = @($Updates | Where-Object { $_.SecurityUpdate -eq $true }).Count
    $highPriorityUpdates = @($Updates | Where-Object { $_.Priority -eq 1 }).Count

    Write-Host "║ " -NoNewline -ForegroundColor DarkGray
    Write-Host "📊 Update Statistics:" -NoNewline -ForegroundColor Yellow
    Write-Host (" " * ($width - 24)) -NoNewline
    Write-Host "║" -ForegroundColor DarkGray
    
    Write-Host "║   " -NoNewline -ForegroundColor DarkGray
    Write-Host "📦 Total Available: " -NoNewline -ForegroundColor White
    Write-Host "$totalUpdates packages" -NoNewline -ForegroundColor Green
    Write-Host (" " * ($width - 34 - $totalUpdates.ToString().Length)) -NoNewline
    Write-Host "║" -ForegroundColor DarkGray
    
    Write-Host "║   " -NoNewline -ForegroundColor DarkGray
    Write-Host "✅ Will Update: " -NoNewline -ForegroundColor White
    Write-Host "$activeUpdates packages" -NoNewline -ForegroundColor $(if ($activeUpdates -gt 0) { "Green" } else { "Gray" })
    Write-Host (" " * ($width - 30 - $activeUpdates.ToString().Length)) -NoNewline
    Write-Host "║" -ForegroundColor DarkGray
    
    Write-Host "║   " -NoNewline -ForegroundColor DarkGray
    Write-Host "🚫 Excluded: " -NoNewline -ForegroundColor White
    Write-Host "$excludedCount packages" -NoNewline -ForegroundColor $(if ($excludedCount -gt 0) { "Red" } else { "Gray" })
    Write-Host (" " * ($width - 27 - $excludedCount.ToString().Length)) -NoNewline
    Write-Host "║" -ForegroundColor DarkGray
    
    Write-Host "║   " -NoNewline -ForegroundColor DarkGray
    Write-Host "🔒 Security Updates: " -NoNewline -ForegroundColor White
    Write-Host "$securityUpdates packages" -NoNewline -ForegroundColor $(if ($securityUpdates -gt 0) { "Red" } else { "Gray" })
    Write-Host (" " * ($width - 35 - $securityUpdates.ToString().Length)) -NoNewline
    Write-Host "║" -ForegroundColor DarkGray
    
    Write-Host "║   " -NoNewline -ForegroundColor DarkGray
    Write-Host "🔖 High Priority: " -NoNewline -ForegroundColor White
    Write-Host "$highPriorityUpdates packages" -NoNewline -ForegroundColor $(if ($highPriorityUpdates -gt 0) { "Red" } else { "Gray" })
    Write-Host (" " * ($width - 32 - $highPriorityUpdates.ToString().Length)) -NoNewline
    Write-Host "║" -ForegroundColor DarkGray

    Write-Host "║   " -NoNewline -ForegroundColor DarkGray
    Write-Host "⚙️ Mode: " -NoNewline -ForegroundColor White
    Write-Host $Mode -NoNewline -ForegroundColor DarkMagenta
    if ($CustomParams) {
        Write-Host " (Custom: $CustomParams)" -NoNewline -ForegroundColor DarkGray
    }
    Write-Host (" " * ($width - 14 - $Mode.Length - $(if ($CustomParams) { $CustomParams.Length + 10 } else { 0 }))) -NoNewline
    Write-Host "║" -ForegroundColor DarkGray

    # 100% AUTOMATIC TABLE - ALL CALCULATED WITH FORMULAS
    if ($Updates.Count -gt 0) {
        # Automatic border calculation
        Write-Host "╠" -NoNewline -ForegroundColor DarkGray
        Write-Host ("─" * ($width - 2)) -NoNewline -ForegroundColor DarkGray
        Write-Host "╣" -ForegroundColor DarkGray
        
        # Automatic title line with perfect centering
        $titleContent = "📋 ALL AVAILABLE PACKAGES - COMPLETE LIST ($($Updates.Count) total):"
        $titlePadding = $width - $titleContent.Length - 4  # 4 for "║ " and " ║"
        $titleLeftPad = [Math]::Floor($titlePadding / 2)
        $titleRightPad = $titlePadding - $titleLeftPad
        
        Write-Host "║ " -NoNewline -ForegroundColor DarkGray
        for ($i = 0; $i -lt [math]::Max(0, $titleLeftPad); $i++) { Write-Host " " -NoNewline }
        Write-Host $titleContent -NoNewline -ForegroundColor Yellow
        for ($i = 0; $i -lt [math]::Max(0, $titleRightPad); $i++) { Write-Host " " -NoNewline }
        Write-Host " ║" -ForegroundColor DarkGray
        
        # Automatic border
        Write-Host "╠" -NoNewline -ForegroundColor DarkGray
        Write-Host ("─" * ($width - 2)) -NoNewline -ForegroundColor DarkGray
        Write-Host "╣" -ForegroundColor DarkGray

        # AUTOMATIC COLUMN WIDTH CALCULATION
        $totalTableWidth = $width - 4  # Remove "║ " and " ║"
        $separatorWidth = 9  # " │ " × 3 separators
        $availableForColumns = $totalTableWidth - $separatorWidth
        
        # Distribute width automatically based on content needs
        $statusW = 6    # Fixed for icons
        $remainingWidth = $availableForColumns - $statusW
        
        # Calculate proportional widths
        $nameW = [Math]::Floor($remainingWidth * 0.50)      # 45% for names
        $versionW = [Math]::Floor($remainingWidth * 0.40)   # 35% for versions  
        $sourceW = $remainingWidth - $nameW - $versionW     # Remaining for source
        
        # Automatic table header with calculated widths
        Write-Host "║ " -NoNewline -ForegroundColor DarkGray
        Write-Host "STATUS".PadRight($statusW) -NoNewline -ForegroundColor White
        Write-Host " │ " -NoNewline -ForegroundColor DarkGray
        Write-Host "APPLICATION NAME".PadRight($nameW) -NoNewline -ForegroundColor White
        Write-Host " │ " -NoNewline -ForegroundColor DarkGray
        Write-Host "VERSION UPGRADE".PadRight($versionW) -NoNewline -ForegroundColor White
        Write-Host " │ " -NoNewline -ForegroundColor DarkGray
        Write-Host "SOURCE".PadRight($sourceW) -NoNewline -ForegroundColor White
        Write-Host " ║" -ForegroundColor DarkGray
        
        # Automatic separator with calculated widths
        Write-Host "╠" -NoNewline -ForegroundColor DarkGray
        Write-Host ("─" * ($statusW + 1)) -NoNewline -ForegroundColor DarkGray
        Write-Host "─┼─" -NoNewline -ForegroundColor DarkGray
        Write-Host ("─" * ($nameW)) -NoNewline -ForegroundColor DarkGray
        Write-Host "─┼─" -NoNewline -ForegroundColor DarkGray
        Write-Host ("─" * $versionW) -NoNewline -ForegroundColor DarkGray
        Write-Host "─┼─" -NoNewline -ForegroundColor DarkGray
        Write-Host ("─" * $sourceW) -NoNewline -ForegroundColor DarkGray
        Write-Host "─╣" -ForegroundColor DarkGray

        # Display all packages with perfect alignment
        $sortedUpdates = $Updates | Sort-Object Priority, Name
        $counter = 0
        
        foreach ($update in $sortedUpdates) {
            $counter++
            $isExcluded = ($ExcludeApps -contains $update.Id) -or ($PersistentExcludeApps -contains $update.Id)
            
            # Status icons
            $priorityIcon = switch ($update.Priority) { 1 { "🔴" } 2 { "🟡" } 3 { "🟢" } }
            $excludeIcon = if ($isExcluded) { "🚫   " } else { "✅" }
            $statusText = "$priorityIcon$excludeIcon"
            
            # Clean app name
            $cleanId = ($update.Id -replace '<.*$', '').Trim()
            $appName = if ($update.Name -and $update.Name.Trim() -ne "") { $update.Name.Trim() } else { $cleanId }
            $displayName = if ($appName.Length -gt ($nameW - 1)) { 
                $appName.Substring(0, $nameW - 4) + "..." 
            } else { 
                $appName 
            }
            
            # Version handling
            $oldVer = $update.Version
            $newVer = $update.Available
            $fullVersion = "$oldVer → $newVer"
            
            # Smart version truncation if needed
            if ($fullVersion.Length -gt ($versionW - 1)) {
                $maxLen = [Math]::Floor(($versionW - 7) / 2)  # Account for " → " and padding
                if ($oldVer.Length -gt $maxLen) { $oldVer = $oldVer.Substring(0, $maxLen - 3) + "..." }
                if ($newVer.Length -gt $maxLen) { $newVer = $newVer.Substring(0, $maxLen - 3) + "..." }
            }
            
            # Source handling
            $sourceText = if ($update.Source.Length -gt ($sourceW - 1)) {
                $update.Source.Substring(0, $sourceW - 4) + "..."
            } else {
                $update.Source
            }
            
            # Row colors
            $rowColor = if ($counter % 2 -eq 0) { "White" } else { "Gray" }
            if ($isExcluded) { $rowColor = "DarkGray" }
            
            # Perfect row display
            Write-Host "║ " -NoNewline -ForegroundColor DarkGray
            Write-Host $statusText.PadRight($statusW) -NoNewline
            Write-Host "│ " -NoNewline -ForegroundColor DarkGray
            Write-Host $displayName.PadRight($nameW + 1) -NoNewline -ForegroundColor $rowColor
            Write-Host "│ " -NoNewline -ForegroundColor DarkGray
            
            # Colorful version display
            Write-Host $oldVer -NoNewline -ForegroundColor $(if ($isExcluded) { "DarkGray" } else { "Red" })
            Write-Host " → " -NoNewline -ForegroundColor $(if ($isExcluded) { "DarkGray" } else { "White" })
            Write-Host $newVer -NoNewline -ForegroundColor $(if ($isExcluded) { "DarkGray" } else { "Green" })
            
            # Pad version column perfectly
            $currentVersionLength = "$oldVer → $newVer".Length
            $versionPadding = $versionW - $currentVersionLength
            if ($versionPadding -gt 0) {
                for ($i = 0; $i -lt $versionPadding; $i++) { Write-Host " " -NoNewline }
                Write-Host " " -NoNewline
            }
            
            Write-Host "│ " -NoNewline -ForegroundColor DarkGray
            Write-Host $sourceText.PadRight($sourceW) -NoNewline -ForegroundColor $(if ($isExcluded) { "DarkGray" } else { "Cyan" })
            Write-Host " ║" -ForegroundColor DarkGray
        }
    }

    # Exclusions section
    if ($ExcludeApps.Count -gt 0 -or $PersistentExcludeApps.Count -gt 0) {
        Write-Host "╠$thinBorder╣" -ForegroundColor DarkGray
        Write-Host "║ " -NoNewline -ForegroundColor DarkGray
        Write-Host "🚫 Excluded Applications:" -NoNewline -ForegroundColor Yellow
        Write-Host (" " * ($width - 28)) -NoNewline
        Write-Host "║" -ForegroundColor DarkGray

        if ($PersistentExcludeApps.Count -gt 0) {
            Write-Host "║   " -NoNewline -ForegroundColor DarkGray
            Write-Host "📌 Persistent: " -NoNewline -ForegroundColor White
            $excludeList = $PersistentExcludeApps -join ", "
            if ($excludeList.Length -gt ($width - 20)) {
                $excludeList = $excludeList.Substring(0, $width - 23) + "..."
            }
            Write-Host $excludeList -NoNewline -ForegroundColor DarkMagenta
            Write-Host (" " * ($width - 20 - $excludeList.Length)) -NoNewline
            Write-Host "║" -ForegroundColor DarkGray
        }
    }

    Write-Host "╚$border╝" -ForegroundColor DarkGray
    Write-Host ""

    # Beautiful legend
    Write-Host "📋 LEGEND:" -ForegroundColor Yellow
    Write-Host "   🔴 High Priority  🟡 Medium Priority  🟢 Low Priority  🔒 Security Update" -ForegroundColor White
    Write-Host "   ✅ Will Update    🚫 Excluded" -ForegroundColor White
    Write-Host ""

    # Enhanced prompt
    if ($Updates.Count -gt 0 -and $activeUpdates -gt 0) {
        if ($Mode -eq "dry-run") {
            Write-Host "🔍 This is a DRY-RUN. No changes will be made." -ForegroundColor DarkGray
            Write-Host "Press any key to continue with simulation..." -ForegroundColor DarkGray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        } else {
            Write-Host "🚀 Ready to update $activeUpdates packages. Continue? (" -NoNewline
            Write-Host "Y" -NoNewline -ForegroundColor Green
            Write-Host "es/" -NoNewline
            Write-Host "N" -NoNewline -ForegroundColor Red
            Write-Host "o/" -NoNewline
            Write-Host "D" -NoNewline -ForegroundColor Yellow
            Write-Host "ry-run): " -NoNewline
        }
    } elseif ($Updates.Count -eq 0) {
        Write-Host "✅ No updates available. Your system is up to date!" -ForegroundColor Green
    } else {
        Write-Host "ℹ️  All available updates are excluded. Nothing to update." -ForegroundColor Yellow
    }
}

# Function to get basic system information
function Get-SystemInfo {
    try {
        return @{
            ComputerName = $env:COMPUTERNAME
            OSVersion = (Get-WmiObject -Class Win32_OperatingSystem).Caption
            PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        }
    } catch {
        return @{
            ComputerName = $env:COMPUTERNAME
            OSVersion = "Unknown"
            PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        }
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

# Global variables for dynamic interface
$script:interfaceInitialized = $false
$script:spinnerChars = @('⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏')
$script:spinnerIndex = 0
$script:lastSpinnerUpdate = Get-Date


# Removed unused function

function Update-ModernInterface {
    param(
        [string]$AppName = "Application",
        [string]$AppId = "",
        [string]$Version = "1.0.0",
        [string]$Phase = "Initialisation",
        [int]$Progress = 0,
        [int]$CurrentApp = 1,
        [int]$TotalApps = 1,
        [string]$Details = "",
        [string]$Size = "",
        [string]$Publisher = ""
    )
    
    if (-not $script:interfaceInitialized) {
        Initialize-ModernInterface -TotalApps $TotalApps
    }
    
    # Sanitize parameters
    $CurrentApp = [math]::Max(1, $CurrentApp)
    $TotalApps = [math]::Max(1, $TotalApps)
    
    # Update spinner animation
    $now = Get-Date
    if (($now - $script:lastSpinnerUpdate).TotalMilliseconds -gt 150) {
        $script:spinnerIndex = ($script:spinnerIndex + 1) % $script:spinnerChars.Length
        $script:lastSpinnerUpdate = $now
    }
    
    $spinner = $script:spinnerChars[$script:spinnerIndex]
    
    try {
        # Clear content area
        for ($line = 3; $line -lt ($script:boxHeight + 1); $line++) {
            [Console]::SetCursorPosition(0, $line)
            Write-Host "║" -NoNewline -ForegroundColor DarkCyan
            $spaces = ""
            for ($i = 0; $i -lt ($script:boxWidth - 2); $i++) {
                $spaces += " "
            }
            Write-Host $spaces -NoNewline
            Write-Host "║" -ForegroundColor DarkCyan
        }
        
        # ═══ APP INFORMATION ═══
        [Console]::SetCursorPosition(0, 4)
        $appCounter = "[$CurrentApp/$TotalApps]"
        $maxNameLength = $script:boxWidth - 15
        $displayName = if ($AppName.Length -gt $maxNameLength) { 
            $AppName.Substring(0, $maxNameLength - 3) + "..." 
        } else { 
            $AppName 
        }
        
        Write-Host "║  📦 " -NoNewline -ForegroundColor DarkCyan
        Write-Host $appCounter -NoNewline -ForegroundColor Gray
        Write-Host " " -NoNewline
        Write-Host $displayName -NoNewline -ForegroundColor White
        
        # Calculate remaining space and fill with spaces
        $usedSpace = 7 + $appCounter.Length + $displayName.Length
        $remainingSpace = $script:boxWidth - $usedSpace - 2
        if ($remainingSpace -gt 0) {
            for ($i = 0; $i -lt $remainingSpace; $i++) {
                Write-Host " " -NoNewline
            }
        }
        Write-Host " ║" -ForegroundColor DarkCyan
        
        # ═══ VERSION ═══
        if ($Version -and $Version.Trim() -ne "") {
            [Console]::SetCursorPosition(0, 5)
            Write-Host "║  🔄 Version: " -NoNewline -ForegroundColor DarkCyan
            Write-Host $Version -NoNewline -ForegroundColor Green
            
            $usedSpace = 15 + $Version.Length
            $remainingSpace = $script:boxWidth - $usedSpace - 2
            if ($remainingSpace -gt 0) {
                for ($i = 0; $i -lt $remainingSpace; $i++) {
                    Write-Host " " -NoNewline
                }
            }
            Write-Host " ║" -ForegroundColor DarkCyan
        }
        
        # ═══ SIZE ═══
        if ($Size -and $Size.Trim() -ne "") {
            [Console]::SetCursorPosition(0, 6)
            Write-Host "║  💾 Taille: " -NoNewline -ForegroundColor DarkCyan
            Write-Host $Size -NoNewline -ForegroundColor Yellow
            
            $usedSpace = 13 + $Size.Length
            $remainingSpace = $script:boxWidth - $usedSpace - 2
            if ($remainingSpace -gt 0) {
                for ($i = 0; $i -lt $remainingSpace; $i++) {
                    Write-Host " " -NoNewline
                }
            }
            Write-Host " ║" -ForegroundColor DarkCyan
        }
        
        # ═══ SEPARATOR ═══
        [Console]::SetCursorPosition(0, 8)
        Write-Host "║" -NoNewline -ForegroundColor DarkCyan
        for ($i = 0; $i -lt ($script:boxWidth - 2); $i++) {
            Write-Host "─" -NoNewline -ForegroundColor DarkGray
        }
        Write-Host "║" -ForegroundColor DarkCyan
        
        # ═══ STATUS WITH SPINNER ═══
        [Console]::SetCursorPosition(0, 10)
        $maxPhaseLength = $script:boxWidth - 15
        $phaseDisplay = if ($Phase.Length -gt $maxPhaseLength) { 
            $Phase.Substring(0, $maxPhaseLength - 3) + "..." 
        } else { 
            $Phase 
        }
        
        Write-Host "║  " -NoNewline -ForegroundColor DarkCyan
        Write-Host $spinner -NoNewline -ForegroundColor Cyan
        Write-Host " " -NoNewline
        Write-Host $phaseDisplay -NoNewline -ForegroundColor Magenta
        
        $usedSpace = 4 + $phaseDisplay.Length
        $remainingSpace = $script:boxWidth - $usedSpace - 2
        if ($remainingSpace -gt 0) {
            for ($i = 0; $i -lt $remainingSpace; $i++) {
                Write-Host " " -NoNewline
            }
        }
        Write-Host " ║" -ForegroundColor DarkCyan
        
        # ═══ DETAILS ═══
        if ($Details -and $Details.Trim() -ne "") {
            [Console]::SetCursorPosition(0, 11)
            $maxDetailsLength = $script:boxWidth - 10
            $detailsDisplay = if ($Details.Length -gt $maxDetailsLength) { 
                $Details.Substring(0, $maxDetailsLength - 3) + "..." 
            } else { 
                $Details 
            }
            
            Write-Host "║    ℹ️  " -NoNewline -ForegroundColor DarkCyan
            Write-Host $detailsDisplay -NoNewline -ForegroundColor DarkYellow
            
            $usedSpace = 8 + $detailsDisplay.Length
            $remainingSpace = $script:boxWidth - $usedSpace - 2
            if ($remainingSpace -gt 0) {
                for ($i = 0; $i -lt $remainingSpace; $i++) {
                    Write-Host " " -NoNewline
                }
            }
            Write-Host " ║" -ForegroundColor DarkCyan
        }
        
        # ═══ OVERALL PROGRESS ═══
        [Console]::SetCursorPosition(0, 13)
        Write-Host "║" -NoNewline -ForegroundColor DarkCyan
        for ($i = 0; $i -lt ($script:boxWidth - 2); $i++) {
            Write-Host "─" -NoNewline -ForegroundColor DarkGray
        }
        Write-Host "║" -ForegroundColor DarkCyan
        
        [Console]::SetCursorPosition(0, 14)
        $overallProgress = if ($TotalApps -gt 0) { [math]::Floor(($CurrentApp / $TotalApps) * 100) } else { 0 }
        $progressText = "🎯 Progression globale: $CurrentApp/$TotalApps ($overallProgress%)"
        
        Write-Host "║  " -NoNewline -ForegroundColor DarkCyan
        Write-Host $progressText -NoNewline -ForegroundColor Blue
        
        $usedSpace = 3 + $progressText.Length
        $remainingSpace = $script:boxWidth - $usedSpace - 2
        if ($remainingSpace -gt 0) {
            for ($i = 0; $i -lt $remainingSpace; $i++) {
                Write-Host " " -NoNewline
            }
        }
        Write-Host " ║" -ForegroundColor DarkCyan
        
    } catch {
        # Simple fallback
        [Console]::SetCursorPosition(0, 8)
        Write-Host "║  🔄 Mise à jour en cours..." -NoNewline -ForegroundColor DarkCyan
        $remainingSpace = $script:boxWidth - 25
        if ($remainingSpace -gt 0) {
            for ($i = 0; $i -lt $remainingSpace; $i++) {
                Write-Host " " -NoNewline
            }
        }
        Write-Host " ║" -ForegroundColor DarkCyan
    }
}

# Alias for backward compatibility
function Update-DynamicInterface {
    param(
        [string]$AppName = "Application",
        [string]$AppId = "",
        [string]$Version = "1.0.0",
        [string]$Phase = "Initialisation",
        [int]$Progress = 0,
        [int]$CurrentApp = 1,
        [int]$TotalApps = 1,
        [string]$Details = "",
        [string]$Size = "",
        [string]$Publisher = ""
    )
    Update-ModernInterface @PSBoundParameters
}

function Show-CompletionMessage {
    param(
        [string]$Message = "✅ Update completed successfully!",
        [string]$Color = "Green",
        [int]$Duration = 3000
    )
    
    if ($script:interfaceInitialized) {
        try {
            # Elegant completion separator
            [Console]::SetCursorPosition(0, 14)
            Write-Host "║" -NoNewline -ForegroundColor DarkCyan
            $separatorWidth = [math]::Max(0, $script:boxWidth - 2)
            for ($i = 0; $i -lt $separatorWidth; $i++) {
                Write-Host "═" -NoNewline -ForegroundColor Green
            }
            Write-Host "║" -ForegroundColor DarkCyan
            
            # Centered completion message with animation
            [Console]::SetCursorPosition(0, 15)
            $messagePadding = [math]::Max(0, $script:boxWidth - $Message.Length - 4)
            $leftPad = [math]::Max(0, [math]::Floor($messagePadding / 2))
            $rightPad = [math]::Max(0, $messagePadding - $leftPad)
            
            Write-Host "║" -NoNewline -ForegroundColor DarkCyan
            for ($i = 0; $i -lt $leftPad; $i++) {
                Write-Host " " -NoNewline
            }
            Write-Host $Message -NoNewline -ForegroundColor $Color
            for ($i = 0; $i -lt $rightPad; $i++) {
                Write-Host " " -NoNewline
            }
            Write-Host "║" -ForegroundColor DarkCyan
            
            # Professional bottom border
            [Console]::SetCursorPosition(0, 16)
            Write-Host "╚" -NoNewline -ForegroundColor DarkCyan
            # Remplacer Get-SafeChar par une boucle directe
            $borderWidth = [math]::Max(0, $script:boxWidth - 2)
            for ($i = 0; $i -lt $borderWidth; $i++) {
                Write-Host "═" -NoNewline -ForegroundColor DarkCyan
            }
            Write-Host "╝" -ForegroundColor DarkCyan
            
            # Brief pause for user to see the completion
            if ($Duration -gt 0) {
                Start-Sleep -Milliseconds $Duration
            }
            
        } catch {
            # Fallback simple message
            Write-Host "`n$Message" -ForegroundColor $Color
        }
    } else {
        Write-Host "`n$Message" -ForegroundColor $Color
    }
}

function Show-ErrorMessage {
    param(
        [string]$Message = "❌ An error occurred",
        [string]$Details = "",
        [int]$Duration = 2000
    )
    
    if ($script:interfaceInitialized) {
        try {
            # Error separator with red accent
            [Console]::SetCursorPosition(0, 14)
            Write-Host "║" -NoNewline -ForegroundColor DarkCyan
            $separatorWidth = [math]::Max(0, $script:boxWidth - 2)
            for ($i = 0; $i -lt $separatorWidth; $i++) {
                Write-Host "═" -NoNewline -ForegroundColor Red
            }
            Write-Host "║" -ForegroundColor DarkCyan
            
            # Centered error message
            [Console]::SetCursorPosition(0, 15)
            $messagePadding = [math]::Max(0, $script:boxWidth - $Message.Length - 4)
            $leftPad = [math]::Max(0, [math]::Floor($messagePadding / 2))
            $rightPad = [math]::Max(0, $messagePadding - $leftPad)
            
            Write-Host "║" -NoNewline -ForegroundColor DarkCyan
            for ($i = 0; $i -lt $leftPad; $i++) {
                Write-Host " " -NoNewline
            }
            Write-Host $Message -NoNewline -ForegroundColor Red
            for ($i = 0; $i -lt $rightPad; $i++) {
                Write-Host " " -NoNewline
            }
            Write-Host "║" -ForegroundColor DarkCyan
            
            # Optional details line
            if ($Details -and $Details.Trim() -ne "") {
                [Console]::SetCursorPosition(0, 16)
                $detailsDisplay = if ($Details.Length -gt ($script:boxWidth - 8)) { 
                    $Details.Substring(0, $script:boxWidth - 11) + "..." 
                } else { 
                    $Details 
                }
                $detailsPadding = [math]::Max(0, $script:boxWidth - $detailsDisplay.Length - 4)
                $leftPadDetails = [math]::Floor($detailsPadding / 2)
                $rightPadDetails = $detailsPadding - $leftPadDetails
                
                Write-Host "║" -NoNewline -ForegroundColor DarkCyan
                for ($i = 0; $i -lt [math]::Max(0, $leftPadDetails); $i++) { Write-Host " " -NoNewline }
                Write-Host $detailsDisplay -NoNewline -ForegroundColor Yellow
                for ($i = 0; $i -lt [math]::Max(0, $rightPadDetails); $i++) { Write-Host " " -NoNewline }
                Write-Host "║" -ForegroundColor DarkCyan
                
                # Bottom border for error with details
                [Console]::SetCursorPosition(0, 17)
                Write-Host "╚" -NoNewline -ForegroundColor DarkCyan
                Write-Host ("═" * [math]::Max(0, $script:boxWidth - 2)) -NoNewline -ForegroundColor DarkCyan
                Write-Host "╝" -ForegroundColor DarkCyan
            } else {
                # Bottom border for simple error
                [Console]::SetCursorPosition(0, 16)
                Write-Host "╚" -NoNewline -ForegroundColor DarkCyan
                Write-Host ("═" * [math]::Max(0, $script:boxWidth - 2)) -NoNewline -ForegroundColor DarkCyan
                Write-Host "╝" -ForegroundColor DarkCyan
            }
            
            # Brief pause for user to see the error
            if ($Duration -gt 0) {
                Start-Sleep -Milliseconds $Duration
            }
            
        } catch {
            # Fallback simple error message
            Write-Host "`n$Message" -ForegroundColor Red
            if ($Details) {
                Write-Host $Details -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "`n$Message" -ForegroundColor Red
        if ($Details) {
            Write-Host $Details -ForegroundColor Yellow
        }
    }
}

# Helper function to convert sizes to MB
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

# Dynamic Interface Functions for Professional UI - Redirects to Modern Interface
function Initialize-DynamicInterface {
    param(
        [int]$TotalApps
    )
    
    # Redirect to modern interface
    Initialize-ModernInterface -TotalApps $TotalApps
}

function Update-DynamicInterface {
    param(
        [int]$CurrentApp,
        [int]$TotalApps,
        [object]$Update,
        [hashtable]$Details,
        [string]$Phase = "Initializing",
        [int]$Progress = 0,
        [string]$Status = "",
        [string]$DownloadInfo = "",
        [string]$SpinnerChar = "⠋"
    )
    
    # Convert to modern interface parameters
    $appName = if ($Update.Name) { $Update.Name } else { "Application" }
    $version = if ($Update.Available) { "$($Update.Version) → $($Update.Available)" } else { $Update.Version }
    $size = if ($Details.Size) { $Details.Size } else { "" }
    $publisher = if ($Details.Publisher) { $Details.Publisher } else { "" }
    $details = if ($Status) { $Status } else { $DownloadInfo }
    
    # Call the modern interface
    Update-ModernInterface -AppName $appName -Version $version -Phase $Phase -CurrentApp $CurrentApp -TotalApps $TotalApps -Details $details -Size $size -Publisher $publisher
}

function Get-SpinnerChar {
    param([int]$Step)
    $spinnerChars = @("⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏")
    return $spinnerChars[$Step % $spinnerChars.Length]
}


# Modern function to display update interface with loading states
function Show-ModernUpdateInterface {
    param(
        [string]$AppName,
        [string]$AppId = "",
        [string]$CurrentVersion = "",
        [string]$NewVersion = "",
        [string]$Publisher = "",
        [string]$Size = "",
        [string]$Status = "Preparing",
        [int]$SpinnerStep = 0,
        [int]$CurrentApp = 1,
        [int]$TotalApps = 1
    )
    
    # Clear screen
    Clear-Host
    
    # Create spinner with more modern characters
    $spinnerChars = @('⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏')
    $spinner = $spinnerChars[$SpinnerStep % 10]
    
    # Modern header with improved design
    Write-Host ""
    Write-Host "  ╭─────────────────────────────────────────────────────────────╮" -ForegroundColor DarkCyan
    Write-Host "  │" -NoNewline -ForegroundColor DarkCyan
    Write-Host "                 Windows Package Manager                     " -NoNewline -ForegroundColor White
    Write-Host "│" -ForegroundColor DarkCyan
    Write-Host "  ╰─────────────────────────────────────────────────────────────╯" -ForegroundColor DarkCyan
    Write-Host ""
    
    # Application counter with modern design
    Write-Host "  📊 Progress: " -NoNewline -ForegroundColor Cyan
    Write-Host "$CurrentApp" -NoNewline -ForegroundColor Yellow
    Write-Host " / " -NoNewline -ForegroundColor DarkGray
    Write-Host "$TotalApps" -NoNewline -ForegroundColor Yellow
    Write-Host " applications" -ForegroundColor DarkGray
    Write-Host ""
    
    # Séparateur élégant
    Write-Host "  " -NoNewline
    Write-Host ("─" * 60) -ForegroundColor DarkGray
    Write-Host ""
    
    # Nom de l'application avec icône
    $displayName = if ($AppName.Length -gt 55) { $AppName.Substring(0, 52) + "..." } else { $AppName }
    Write-Host "  📦 " -NoNewline -ForegroundColor Blue
    Write-Host $displayName -ForegroundColor White
    
    # ID de l'application avec style
    if ($AppId) {
        Write-Host "  🆔 " -NoNewline -ForegroundColor DarkGray
        Write-Host $AppId -ForegroundColor DarkGray
    }
    
    Write-Host ""
    
    # Versions with colors and improved style
    if ($CurrentVersion -and $NewVersion) {
        Write-Host "  📊 " -NoNewline -ForegroundColor Magenta
        Write-Host "Version: " -NoNewline -ForegroundColor Gray
        Write-Host $CurrentVersion -NoNewline -ForegroundColor Red
        Write-Host "  ➜  " -NoNewline -ForegroundColor Yellow
        Write-Host $NewVersion -ForegroundColor Green
    }
    
    # Publisher information on separate line
    if ($Publisher) {
        Write-Host "  🏢 " -NoNewline -ForegroundColor Blue
        Write-Host "Publisher: " -NoNewline -ForegroundColor Gray
        Write-Host $Publisher -ForegroundColor DarkGray
    }
    
    # Size information on separate line
    if ($Size) {
        Write-Host "  💾 " -NoNewline -ForegroundColor Green
        Write-Host "Size: " -NoNewline -ForegroundColor Gray
        Write-Host $Size -ForegroundColor DarkGray
    }
    
    # Elegant separator
    Write-Host "  " -NoNewline
    Write-Host ("─" * 60) -ForegroundColor DarkGray
    Write-Host ""
    
    # Progress section with modern design
    Write-Host "  🔄 " -NoNewline -ForegroundColor Blue
    Write-Host "Progress Status" -ForegroundColor White
    Write-Host ""
    
    # Progress states with modern style
    $steps = @(
        @{ Name = "Preparation"; Key = "Preparing"; Icon = "🔧" },
        @{ Name = "Download"; Key = "Downloading"; Icon = "📥" },
        @{ Name = "Installation"; Key = "Installing"; Icon = "⚙️" },
        @{ Name = "Completion"; Key = "Completed"; Icon = "✨" }
    )
    
    foreach ($step in $steps) {
        Write-Host "     " -NoNewline
        
        if ($Status -eq $step.Key) {
            # Current step with modern spinner
            Write-Host "$spinner " -NoNewline -ForegroundColor Cyan
            Write-Host $step.Name -NoNewline -ForegroundColor White
            Write-Host " in progress..." -ForegroundColor Gray
        } elseif (($Status -eq "Downloading" -and $step.Key -eq "Preparing") -or
                  ($Status -eq "Installing" -and ($step.Key -eq "Preparing" -or $step.Key -eq "Downloading")) -or
                  ($Status -eq "Completed" -and $step.Key -ne "Completed")) {
            # Completed step with style
            Write-Host "✅ " -NoNewline -ForegroundColor Green
            Write-Host $step.Name -NoNewline -ForegroundColor Green
            Write-Host " completed" -ForegroundColor DarkGreen
        } elseif ($Status -eq "Failed") {
            # Failed step with style
            Write-Host "❌ " -NoNewline -ForegroundColor Red
            Write-Host $step.Name -NoNewline -ForegroundColor Red
            Write-Host " failed" -ForegroundColor DarkRed
        } else {
            # Pending step with style
            Write-Host "⏳ " -NoNewline -ForegroundColor DarkGray
            Write-Host $step.Name -NoNewline -ForegroundColor DarkGray
            Write-Host " pending" -ForegroundColor DarkGray
        }
    }
    
    Write-Host ""
    Write-Host "  " -NoNewline
    Write-Host ("─" * 60) -ForegroundColor DarkGray
    Write-Host ""
    
    # Status message with modern design
    $statusMessage = switch ($Status) {
        "Preparing" { "🔧 Preparing update..." }
        "Downloading" { "📥 Downloading files..." }
        "Installing" { "⚙️ Installing application..." }
        "Completed" { "✨ Update completed successfully!" }
        "Failed" { "❌ Update failed" }
        default { "🔄 Processing..." }
    }
    
    $statusColor = switch ($Status) {
        "Preparing" { "Yellow" }
        "Downloading" { "Cyan" }
        "Installing" { "Blue" }
        "Completed" { "Green" }
        "Failed" { "Red" }
        default { "Gray" }
    }
    
    # Display status message with style
    Write-Host "  " -NoNewline
    Write-Host $statusMessage -ForegroundColor $statusColor
    Write-Host ""
    
    # Pied de page élégant
    Write-Host "  " -NoNewline
    Write-Host ("─" * 60) -ForegroundColor DarkGray
}

# Function to interpret winget exit codes and provide meaningful messages
function Get-WingetErrorMessage {
    param([int]$ExitCode)
    
    switch ($ExitCode) {
        0 { return "Success" }
        -1978335189 { return "Package upgrade not supported by winget - use publisher's method" }
        -1978335188 { return "No installed package found matching criteria" }
        -1978335187 { return "No available upgrade found" }
        -1978335186 { return "Package installation failed" }
        -1978335185 { return "Package already installed" }
        -1978335184 { return "Package not found in configured sources" }
        1 { return "General installation failure" }
        2 { return "Installation cancelled by user" }
        3010 { return "Installation successful but restart required" }
        1603 { return "Fatal error during installation" }
        1618 { return "Another installation is already in progress" }
        1619 { return "Installation package could not be opened" }
        1620 { return "Installation package could not be opened (corrupt)" }
        1633 { return "This installation package is not supported on this platform" }
        default { return "Unknown error (Exit code: $ExitCode)" }
    }
}

# Function to simulate installation steps with modern interface
function Update-AppWithDetailedUI {
    param(
        [string]$AppId,
        [string]$AppName,
        [string]$CurrentVersion = "",
        [string]$NewVersion = "",
        [string]$Publisher = "",
        [string]$Size = "",
        [int]$AppNumber = 1,
        [int]$TotalApps = 1
    )
    
    try {
        $spinnerStep = 0
        
        # Étape 1: Préparation (plus fluide)
        for ($i = 0; $i -lt 6; $i++) {
            Show-ModernUpdateInterface -AppName $AppName -AppId $AppId -CurrentVersion $CurrentVersion -NewVersion $NewVersion -Publisher $Publisher -Size $Size -Status "Preparing" -SpinnerStep $spinnerStep -CurrentApp $AppNumber -TotalApps $TotalApps
            Start-Sleep -Milliseconds 200
            $spinnerStep++
        }
        
        # Étape 2: Téléchargement (plus fluide)
        for ($i = 0; $i -lt 10; $i++) {
            Show-ModernUpdateInterface -AppName $AppName -AppId $AppId -CurrentVersion $CurrentVersion -NewVersion $NewVersion -Publisher $Publisher -Size $Size -Status "Downloading" -SpinnerStep $spinnerStep -CurrentApp $AppNumber -TotalApps $TotalApps
            Start-Sleep -Milliseconds 150
            $spinnerStep++
        }
        
        # Étape 3: Installation (plus fluide)
        for ($i = 0; $i -lt 8; $i++) {
            Show-ModernUpdateInterface -AppName $AppName -AppId $AppId -CurrentVersion $CurrentVersion -NewVersion $NewVersion -Publisher $Publisher -Size $Size -Status "Installing" -SpinnerStep $spinnerStep -CurrentApp $AppNumber -TotalApps $TotalApps
            Start-Sleep -Milliseconds 200
            $spinnerStep++
        }
        
        # Exécution de la commande winget réelle
        $wingetArgs = @("upgrade", "--id", $AppId, "--silent", "--accept-package-agreements", "--accept-source-agreements")
        
        if ($CustomParams) {
            $wingetArgs += $CustomParams.Split(' ')
        }
        
        $result = & winget @wingetArgs 2>&1
        
        # Vérifier le résultat avec gestion d'erreur améliorée
        if ($LASTEXITCODE -eq 0) {
            Show-ModernUpdateInterface -AppName $AppName -AppId $AppId -CurrentVersion $CurrentVersion -NewVersion $NewVersion -Publisher $Publisher -Size $Size -Status "Completed" -SpinnerStep 0 -CurrentApp $AppNumber -TotalApps $TotalApps
            Start-Sleep -Milliseconds 2000
            return $true
        } else {
            # Afficher l'interface d'échec
            Show-ModernUpdateInterface -AppName $AppName -AppId $AppId -CurrentVersion $CurrentVersion -NewVersion $NewVersion -Publisher $Publisher -Size $Size -Status "Failed" -SpinnerStep 0 -CurrentApp $AppNumber -TotalApps $TotalApps
            
            # Afficher un message d'erreur détaillé
            $errorMessage = Get-WingetErrorMessage -ExitCode $LASTEXITCODE
            Write-Host ""
            Write-Host "  ❌ " -NoNewline -ForegroundColor Red
            Write-Host $errorMessage -ForegroundColor Red
            
            # Afficher des détails supplémentaires si disponibles
            if ($result -and $result.ToString().Trim() -ne "") {
                $cleanResult = ($result | Out-String).Trim()
                if ($cleanResult -notmatch "^\s*$" -and $cleanResult.Length -lt 200) {
                    Write-Host "  💡 " -NoNewline -ForegroundColor Yellow
                    Write-Host "Details: $cleanResult" -ForegroundColor DarkGray
                }
            }
            
            Write-Host ""
            Start-Sleep -Milliseconds 2000
            return $false
        }
        
    } catch {
        Show-ModernUpdateInterface -AppName $AppName -AppId $AppId -CurrentVersion $CurrentVersion -NewVersion $NewVersion -Publisher $Publisher -Size $Size -Status "Failed" -SpinnerStep 0 -CurrentApp $AppNumber -TotalApps $TotalApps
        
        # Afficher l'erreur d'exception
        Write-Host ""
        Write-Host "  ❌ " -NoNewline -ForegroundColor Red
        Write-Host "Exception: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Start-Sleep -Milliseconds 2000
        return $false
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
    
    # Interface simple - pas de variables complexes nécessaires
    
    # Create logs directory if it doesn't exist
    $logDir = Join-Path $PSScriptRoot "logs"
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir | Out-Null
    }

    # Create a new log file for this session
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $sessionLogFile = Join-Path $logDir "winget_update_$timestamp.log"
    
    # Function to write to log
    function Write-Log {
        param([string]$Message)
        $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
        Add-Content -Path $sessionLogFile -Value $logMessage
    }

    $baseCommand = "winget"
    $additionalParams = @()
    
    # Handle dry-run mode
    if ($Mode -eq "dry-run") {
        Write-Host "🔍 DRY-RUN MODE - Simulation Only (No Changes Will Be Made)" -ForegroundColor DarkGray
        Write-Host "═" * 80 -ForegroundColor DarkGray
        
        $currentApp = 0
        foreach ($update in $Updates) {
            if (($ExcludeApps -notcontains $update.Id) -and ($PersistentExcludeApps -notcontains $update.Id)) {
                $currentApp++
                
                # Fix ID display issue (remove < character and clean up)
                $cleanId = $update.Id -replace '<.*$', ''
                $cleanId = $cleanId.Trim()
                
                # Use clean name or ID as fallback
                $displayName = if ($update.Name -and $update.Name.Trim() -ne "") { 
                    $update.Name.Trim() 
                } else { 
                    $cleanId 
                }
                
                Write-Host "`n📦 Would update: " -NoNewline -ForegroundColor DarkGray
                Write-Host $displayName -ForegroundColor White
                Write-Host "   ID: " -NoNewline -ForegroundColor Gray
                Write-Host $cleanId -ForegroundColor DarkGray
                Write-Host "   Version: " -NoNewline -ForegroundColor Gray
                Write-Host $update.Version -NoNewline -ForegroundColor Red
                Write-Host " → " -NoNewline -ForegroundColor White
                Write-Host $update.Available -ForegroundColor Green
                Write-Host "   Category: " -NoNewline -ForegroundColor Gray
                Write-Host $update.Category -ForegroundColor DarkMagenta
                Write-Host "   Priority: " -NoNewline -ForegroundColor Gray
                Write-Host $(switch ($update.Priority) { 1 { "High" } 2 { "Medium" } 3 { "Low" } }) -ForegroundColor $(switch ($update.Priority) { 1 { "Red" } 2 { "Yellow" } 3 { "Green" } })
                if ($update.SecurityUpdate) {
                    Write-Host "   🔒 Security Update" -ForegroundColor Red
                }
            }
        }
        
        Write-Host "`n═" * 80 -ForegroundColor DarkGray
        Write-Host "✅ Dry-run completed. $currentApp packages would be updated." -ForegroundColor Green
        Write-Host "💡 Run without -Mode dry-run to perform actual updates." -ForegroundColor DarkGray
        return
    }
    
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
    
    # Interface simple - pas d'initialisation complexe nécessaire
    Clear-Host

    foreach ($update in $Updates) {
        if (($ExcludeApps -notcontains $update.Id) -and ($PersistentExcludeApps -notcontains $update.Id)) {
            $currentApp++
            $logPath = $null
            
            # Get package details first
            $details = @{
                'Publisher' = "Unknown"
                'License' = "Unknown"
                'Size' = "Unknown"
            }
            
            try {
                $packageInfo = winget show --id $update.Id --accept-source-agreements | Out-String
                
                $details = @{
                    'Publisher' = if ($packageInfo -match "Publisher:\s*(.+)") { $matches[1] } else { "Unknown" }
                    'License' = if ($packageInfo -match "License:\s*(.+)") { $matches[1] } else { "Unknown" }
                }

                if ($packageInfo -match "Download Size:\s*(.+)") {
                    $details.Size = Format-FileSize $matches[1]
                }
                elseif ($packageInfo -match "Installer Url:\s*(.+)") {
                    $details.Size = Get-RemoteFileSize $matches[1]
                }
                else {
                    $details.Size = "Unknown"
                }
            } catch {
                # Use defaults if package info retrieval fails
            }
            
            # Clean application name for display
            $displayName = if ($update.Name -and $update.Name.Trim() -ne "") { 
                $update.Name.Trim() 
            } else { 
                $update.Id.Trim() 
            }
            
            # Interface simple - pas de réinitialisation nécessaire

            try {
                # Use the new modern function for update
                $success = Update-AppWithDetailedUI -AppId $update.Id -AppName $displayName -CurrentVersion $update.Version -NewVersion $update.Available -Publisher $details.Publisher -Size $details.Size -AppNumber $currentApp -TotalApps $totalApps
                
                # Enregistrer le résultat
                $status = if ($success) { "Success" } else { "Failed" }
                $exitCode = if ($success) { 0 } else { 1 }
                
                Write-Log "Update $status for $displayName ($($update.Id)): $($update.Version) -> $($update.Available)"

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

                # Variables pour le suivi de progression
                $currentPhase = "Initializing"
                $phaseProgress = 0
                $downloadProgress = 0
                $overallProgress = [math]::Round((($currentApp - 1) / $totalApps) * 100)
                $script:lastProgressUpdate = Get-Date
                
                # Interface simple - pas de barre de progression complexe
                
                # Lire la sortie ligne par ligne avec buffer amélioré
                $reader = $p.StandardOutput
                $errorReader = $p.StandardError
                $buffer = ""
                $lastProgressUpdate = Get-Date
                
                # Fonction pour mettre à jour la progression de manière fluide
                function Update-RealTimeProgress {
                    param(
                        [string]$Phase,
                        [int]$PhasePercent = 0,
                        [string]$Details = ""
                    )
                    
                    $script:currentPhase = $Phase
                    $script:phaseProgress = $PhasePercent
                    
                    # Calculer la progression globale (chaque app = 100/totalApps %)
                    $appBaseProgress = [math]::Round((($currentApp - 1) / $totalApps) * 100)
                    $appCurrentProgress = [math]::Round(($PhasePercent / $totalApps))
                    $totalProgress = [math]::Min($appBaseProgress + $appCurrentProgress, 100)
                    
                    # Mettre à jour seulement si assez de temps s'est écoulé (évite le spam)
                    $now = Get-Date
                    if (($now - $script:lastProgressUpdate).TotalMilliseconds -gt 100) {
                        # Interface simple - pas de barre de progression
                        $script:lastProgressUpdate = $now
                    }
                }
                
                # Démarrer avec la phase d'initialisation
                Update-RealTimeProgress -Phase "Initializing" -PhasePercent 0 -Details "Starting update process"
                
                while (-not $reader.EndOfStream) {
                    $char = [char]$reader.Read()
                    if ($char -eq "`n") {
                        $line = $buffer.Trim()
                        Write-Log $line

                        # Analyse des différentes phases avec progression améliorée
                        if ($line -match "Downloading (.+)") {
                            $downloadUrl = $matches[1]
                            Update-RealTimeProgress -Phase "Downloading" -PhasePercent 10 -Details "From: $downloadUrl"
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
                                    $downloadPercent = [math]::Min([math]::Round(($currentBytes / $totalBytes) * 100), 100)
                                    $phasePercent = 10 + [math]::Round($downloadPercent * 0.4) # Download = 10-50%
                                    Update-RealTimeProgress -Phase "Downloading" -PhasePercent $phasePercent -Details "$currentSize / $totalSize ($downloadPercent%)"
                                    
                                    # Interface simple - pas de barre de progression
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
                            Update-RealTimeProgress -Phase "Verifying" -PhasePercent 60 -Details "Hash verification successful"
                            Write-Status -Status "Hash verification successful" -Icon "✅" -Color "Green"
                            $downloadStarted = $false
                        }
                        elseif ($line -match "Starting package install") {
                            Update-RealTimeProgress -Phase "Installing" -PhasePercent 70 -Details "Installing package"
                            Write-Status -Status "Installing package..." -Icon "🔧" -Color "Yellow"
                        }
                        elseif ($line -match "Successfully installed") {
                            Update-RealTimeProgress -Phase "Completed" -PhasePercent 100 -Details "Installation successful"
                            Write-Status -Status "Installation completed successfully!" -Icon "✅" -Color "Green"
                        }
                        elseif ($line -match "Installer failed with exit code: (.+)") {
                            Update-RealTimeProgress -Phase "Failed" -PhasePercent 100 -Details "Error code: $($matches[1])"
                            Write-Status -Status "Installation failed (Error code: $($matches[1]))" -Icon "❌" -Color "Red"
                        }
                        elseif ($line -match "Found (.+) \[(.+)\]") {
                            Update-RealTimeProgress -Phase "Resolving" -PhasePercent 5 -Details "Found package: $($matches[1])"
                        }
                        elseif ($line -match "Manifest download") {
                            Update-RealTimeProgress -Phase "Preparing" -PhasePercent 8 -Details "Downloading manifest"
                        }

                        $buffer = ""
                    }
                    else {
                        $buffer += $char
                    }
                }

                $p.WaitForExit()
                
                # Finaliser la barre de progression pour cette application
                if ($p.ExitCode -eq 0) {
                    $status = "Success"
                    # Installation completed successfully
                }
                else {
                    $status = "Failed"
                    # Installation failed
                    # Legacy error message removed - handled by modern interface
                }
                
                # Nettoyer la barre de progression console
                Write-Host "`r" + (" " * 100) + "`r" -NoNewline
                
                # Petite pause pour permettre à l'utilisateur de voir le résultat final
                Start-Sleep -Milliseconds 500

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
    Write-Host "UPDATE SUMMARY" -NoNewline -ForegroundColor DarkGray
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
            Write-Host $result.Name -NoNewline -ForegroundColor DarkGray
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
                    Write-Host "     ↳ Size: $($result.Size)" -ForegroundColor DarkMagenta
                }
            }
        }
    }

    # À la fin du script, afficher le chemin du log de session
    Write-Host "`n📝 Session log: " -NoNewline
    Write-Host $sessionLogFile -ForegroundColor DarkGray
}

# AUTOMATIC STARTUP BANNER - 100% CALCULATED SPACING
Clear-Host
Write-Host ""

# Define banner width and calculate all spacing automatically
$bannerWidth = 80
$borderChar = "═"
$sideChar = "║"

# Top border: ╔ + (width-2) × ═ + ╗
Write-Host "╔" -NoNewline -ForegroundColor DarkGray
Write-Host ($borderChar * ($bannerWidth - 2)) -NoNewline -ForegroundColor DarkGray
Write-Host "╗" -ForegroundColor DarkGray

# Empty line: ║ + (width-2) × space + ║
Write-Host $sideChar -NoNewline -ForegroundColor DarkGray
Write-Host (" " * ($bannerWidth - 2)) -NoNewline
Write-Host $sideChar -ForegroundColor DarkGray

# Title line: ║ + spaces + title + spaces + ║
$title = "WINGET UPDATE MANAGER v3.1.0"
$titlePadding = $bannerWidth - $title.Length - 2
$leftPadding = [Math]::Floor($titlePadding / 2)
$rightPadding = $titlePadding - $leftPadding

Write-Host $sideChar -NoNewline -ForegroundColor DarkGray
Write-Host (" " * $leftPadding) -NoNewline
Write-Host $title -NoNewline -ForegroundColor White
Write-Host (" " * $rightPadding) -NoNewline
Write-Host $sideChar -ForegroundColor DarkGray

# Empty line
Write-Host $sideChar -NoNewline -ForegroundColor DarkGray
Write-Host (" " * ($bannerWidth - 2)) -NoNewline
Write-Host $sideChar -ForegroundColor DarkGray

# Credits line: ║ + spaces + credits + spaces + ║
$credits = "Created by: Sterbweise"
$creditsPadding = $bannerWidth - $credits.Length - 2
$creditsLeftPad = [Math]::Floor($creditsPadding / 2)
$creditsRightPad = $creditsPadding - $creditsLeftPad

Write-Host $sideChar -NoNewline -ForegroundColor DarkGray
Write-Host (" " * $creditsLeftPad) -NoNewline
Write-Host "Created by: " -NoNewline -ForegroundColor Gray
Write-Host "Sterbweise" -NoNewline -ForegroundColor Red
Write-Host (" " * $creditsRightPad) -NoNewline
Write-Host $sideChar -ForegroundColor DarkGray
# Description line: ║ + spaces + description + spaces + ║
$description = "Professional Windows Package Management Solution"
$descPadding = $bannerWidth - $description.Length - 2
$descLeftPad = [Math]::Floor($descPadding / 2)
$descRightPad = $descPadding - $descLeftPad

Write-Host $sideChar -NoNewline -ForegroundColor DarkGray
Write-Host (" " * $descLeftPad) -NoNewline
Write-Host $description -NoNewline -ForegroundColor DarkMagenta
Write-Host (" " * $descRightPad) -NoNewline
Write-Host $sideChar -ForegroundColor DarkGray

# Empty line
Write-Host $sideChar -NoNewline -ForegroundColor DarkGray
Write-Host (" " * ($bannerWidth - 2)) -NoNewline
Write-Host $sideChar -ForegroundColor DarkGray

# Bottom border: ╚ + (width-2) × ═ + ╝
Write-Host "╚" -NoNewline -ForegroundColor DarkGray
Write-Host ($borderChar * ($bannerWidth - 2)) -NoNewline -ForegroundColor DarkGray
Write-Host "╝" -ForegroundColor DarkGray
Write-Host ""

# Get available updates from winget
$updates = @(Get-WingetUpdates)

# Handle potential null array and ensure it's always an array
$updates = @($updates)

if ($updates.Count -gt 0) {
    Show-UpdateSummary -Updates $updates -Mode $Mode -ExcludeApps $ExcludeApps -PersistentExcludeApps $CurrentPersistentExcludeApps -CustomParams $CustomParams
    
    # Handle different modes
    if ($Mode -eq "dry-run") {
        # Dry-run mode is handled in Update-Apps function
        Update-Apps -Updates $updates -Mode $Mode -CustomParams $CustomParams -ExcludeApps $ExcludeApps -PersistentExcludeApps $CurrentPersistentExcludeApps
    } else {
        # Interactive confirmation for other modes
        $activeUpdates = @($updates | Where-Object { 
            ($ExcludeApps -notcontains $_.Id) -and ($CurrentPersistentExcludeApps -notcontains $_.Id) 
        }).Count
        
        if ($activeUpdates -gt 0) {
            $confirmation = Read-Host
            
            switch ($confirmation.ToUpper()) {
                "Y" { 
                    Update-Apps -Updates $updates -Mode $Mode -CustomParams $CustomParams -ExcludeApps $ExcludeApps -PersistentExcludeApps $CurrentPersistentExcludeApps
                }
                "D" { 
                    Write-Host "`n🔍 Switching to dry-run mode..." -ForegroundColor Yellow
                    Update-Apps -Updates $updates -Mode "dry-run" -CustomParams $CustomParams -ExcludeApps $ExcludeApps -PersistentExcludeApps $CurrentPersistentExcludeApps
                }
                default { 
                    Write-Host "`n❌ Update cancelled by user." -ForegroundColor Red
                    Write-Host "💡 Tip: Use -Mode dry-run to simulate updates without making changes." -ForegroundColor DarkGray
                }
            }
        }
    }
} else {
    Write-Host "`n✅ No updates available. Your system is up to date!" -ForegroundColor Green
    Write-Host "💡 Tip: Updates are checked against the configured winget sources." -ForegroundColor DarkGray
}


