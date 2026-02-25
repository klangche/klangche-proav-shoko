# proav-shoko.ps1
# Shōko Launcher – Windows (compatible with PowerShell 5.1 and 7+)
# One-liner: irm https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.ps1 | iex

$ErrorActionPreference = 'Stop'

$RepoOwner = "klangche"
$RepoName  = "klangche-proav-shoko"
$Branch    = "main"

$BaseUrl = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch"

Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "Shōko – USB + Display Diagnostic Tool" -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host ""

# Try to find PowerShell 7 first
$PSExe = $null
if (Get-Command "pwsh" -ErrorAction SilentlyContinue) {
    $PSExe = "pwsh.exe"
    Write-Host "PowerShell 7+ detected – will prefer pwsh.exe" -ForegroundColor Green
} else {
    $PSExe = "powershell.exe"
    Write-Host "Using Windows PowerShell (5.1) – some features may be limited" -ForegroundColor Yellow
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host ""
    Write-Host "Limited mode – full USB tree and display analytics require Administrator rights." -ForegroundColor Yellow
    
    $choice = Read-Host "Elevate now? (y/n)"
    if ($choice -match '^[Yy]$') {
        Write-Host "Requesting elevation..." -ForegroundColor Yellow
        
        $tempScript = "$env:TEMP\shoko-elevated.ps1"
        
        try {
            Write-Host "Downloading main script..." -ForegroundColor Gray
            Invoke-RestMethod -Uri "$BaseUrl/proav-shoko_powershell.ps1" -UseBasicParsing | 
                Out-File -FilePath $tempScript -Encoding UTF8 -Force
            
            Start-Process $PSExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tempScript`"" -Verb RunAs
            Write-Host "Elevation requested. Please allow UAC prompt if shown." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to download or launch elevated script: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "You can manually download and run:" -ForegroundColor Yellow
            Write-Host "  $BaseUrl/proav-shoko_powershell.ps1" -ForegroundColor Cyan
        }
        
        exit
    }
    else {
        Write-Host "Continuing in limited mode..." -ForegroundColor Yellow
    }
}

# ────────────────────────────────────────────────────────────────
# Run main logic (either directly or after elevation)
# ────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "Loading Shōko main script..." -ForegroundColor Gray

try {
    $mainScriptUrl = "$BaseUrl/proav-shoko_powershell.ps1"
    $scriptContent = Invoke-RestMethod -Uri $mainScriptUrl -UseBasicParsing
    
    # Execute the downloaded main script content
    Invoke-Expression $scriptContent
}
catch {
    Write-Host ""
    Write-Host "ERROR: Could not load or execute the main script." -ForegroundColor Red
    Write-Host "Message: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please try one of the following:" -ForegroundColor Yellow
    Write-Host " 1. Run as Administrator manually" -ForegroundColor White
    Write-Host " 2. Download directly:" -ForegroundColor White
    Write-Host "    $mainScriptUrl" -ForegroundColor Cyan
    Write-Host " 3. Check your internet connection / GitHub status" -ForegroundColor White
    Write-Host ""
    pause
}
