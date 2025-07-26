# Winget Update Manager - Installation Script
# Usage: iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/sterbweise/winget-update/main/install.ps1'))

Write-Host "🚀 Installing Winget Update Manager..." -ForegroundColor Cyan

try {
    # Check prerequisites
    Write-Host "📋 Checking prerequisites..." -ForegroundColor Yellow
    
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "PowerShell 5.1 or newer required. Current version: $($PSVersionTable.PSVersion)"
    }
    
    # Check winget
    try {
        $null = Get-Command winget -ErrorAction Stop
        Write-Host "✅ winget found" -ForegroundColor Green
    } catch {
        Write-Host "❌ winget not found. Install from Microsoft Store recommended." -ForegroundColor Red
        throw "winget is not installed"
    }
    
    # Create installation directory
    $installPath = "$env:USERPROFILE\Scripts\WingetUpdateManager"
    Write-Host "📁 Creating installation directory: $installPath" -ForegroundColor Yellow
    
    if (!(Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath -Force | Out-Null
    }
    
    # Download main script
    Write-Host "⬇️  Downloading script..." -ForegroundColor Yellow
    $scriptUrl = "https://raw.githubusercontent.com/sterbweise/winget-update/main/winget-update.ps1"
    $scriptPath = "$installPath\winget-update.ps1"
    
    try {
        Invoke-WebRequest -Uri $scriptUrl -OutFile $scriptPath -UseBasicParsing
        Write-Host "✅ Script downloaded" -ForegroundColor Green
    } catch {
        throw "Download error: $($_.Exception.Message)"
    }
    
    # Verify file was downloaded
    if (!(Test-Path $scriptPath)) {
        throw "Script could not be downloaded"
    }
    
    # Create global function in PowerShell profile
    Write-Host "🔧 Setting up global function..." -ForegroundColor Yellow
    
    # Create profile if it doesn't exist
    if (!(Test-Path $PROFILE)) {
        New-Item -ItemType File -Path $PROFILE -Force | Out-Null
        Write-Host "📝 PowerShell profile created" -ForegroundColor Green
    }
    
    # Check if function already exists
    $profileContent = Get-Content $PROFILE -Raw -ErrorAction SilentlyContinue
    if ($profileContent -notmatch "# Winget Update Manager - Global Function") {
        # Add function to profile with absolute path
        $functionContent = @"

# Winget Update Manager - Global Function
function winget-update {
    param([Parameter(ValueFromRemainingArguments)]`$args)
    & "$scriptPath" @args
}
"@
        
        Add-Content -Path $PROFILE -Value $functionContent
        Write-Host "✅ Global function added to PowerShell profile" -ForegroundColor Green
    } else {
        Write-Host "ℹ️  Global function already exists in profile" -ForegroundColor Blue
        
        # Update existing function with correct path
        $updatedFunction = @"
function winget-update {
    param([Parameter(ValueFromRemainingArguments)]`$args)
    & "$scriptPath" @args
}
"@
        $profileContent = $profileContent -replace 'function winget-update \{[\s\S]*?\n\}', $updatedFunction
        Set-Content -Path $PROFILE -Value $profileContent
        Write-Host "🔄 Updated existing function with correct path" -ForegroundColor Yellow
    }
    
    # Load function in current session (with error handling)
    try {
        # Define function directly in current session instead of sourcing profile
        Invoke-Expression @"
function winget-update {
    param([Parameter(ValueFromRemainingArguments)]`$args)
    & "$scriptPath" @args
}
"@
        Write-Host "✅ Function loaded in current session" -ForegroundColor Green
    } catch {
        Write-Host "⚠️  Function added to profile but couldn't load in current session" -ForegroundColor Yellow
        Write-Host "   Please restart PowerShell or run: . `$PROFILE" -ForegroundColor Gray
    }
    
    Write-Host ""
    Write-Host "🎉 Installation completed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "📖 Usage:" -ForegroundColor Cyan
    Write-Host "   winget-update                    # Normal update" -ForegroundColor White
    Write-Host "   winget-update -Mode dry-run      # Test without installing" -ForegroundColor White
    Write-Host "   winget-update -Help              # Complete help" -ForegroundColor White
    Write-Host ""
    Write-Host "📁 Script installed in: $installPath" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "💡 You can now use 'winget-update' from anywhere!" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "🔧 If the command doesn't work immediately:" -ForegroundColor Cyan
    Write-Host "   • Restart PowerShell, or" -ForegroundColor White
    Write-Host "   • Run: . `$PROFILE" -ForegroundColor White
    Write-Host ""
    
    # Test if function works in current session
    try {
        if (Get-Command winget-update -ErrorAction SilentlyContinue) {
            Write-Host "✅ Function is ready to use in current session!" -ForegroundColor Green
        }
    } catch {
        # Ignore errors in testing
    }
    
} catch {
    Write-Host ""
    Write-Host "❌ Installation error:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "🔧 Manual installation:" -ForegroundColor Yellow
    Write-Host "1. Download winget-update.ps1 from GitHub" -ForegroundColor White
    Write-Host "2. Place it in a folder of your choice" -ForegroundColor White
    Write-Host "3. Run: .\winget-update.ps1" -ForegroundColor White
    exit 1
}