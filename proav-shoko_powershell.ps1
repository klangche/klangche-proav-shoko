<#
.SYNOPSIS
    Shōko Launcher - USB + Display Diagnostic Tool
.DESCRIPTION
    Launches the main Shōko diagnostic script with optional elevation.
    Downloads the latest version from GitHub and runs in memory.
.PARAMETER Verbose
    Show detailed debug information during launch
.EXAMPLE
    .\proav-shoko.ps1
    Run in normal mode
.EXAMPLE
    .\proav-shoko.ps1 -Verbose
    Run with debug output
#>

[CmdletBinding()]
param()

$Repo = "klangche/klangche-proav-shoko"
$Branch = "main"
$Base = "https://raw.githubusercontent.com/$Repo/$Branch"

Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "Shōko – USB + Display Diagnostic Tool Launcher" -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host ""

$isAdmin = try {
    [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
} catch {
    Write-Verbose "Failed to check admin status: $_"
    $null
}

$isAdmin = $isAdmin -and $isAdmin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Limited mode – full features require Administrator rights." -ForegroundColor Yellow
    $elevate = Read-Host "Run with administrator privileges? (y/n)"
    if ($elevate -match '^[Yy]') {
        $temp = "$env:TEMP\shoko-elevated.ps1"
        Write-Verbose "Downloading main script to: $temp"
        try {
            Invoke-RestMethod "$Base/proav-shoko_powershell.ps1" | Out-File $temp -Encoding UTF8
            Write-Verbose "Launching elevated PowerShell"
            Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$temp`" -Verbose:`$$($VerbosePreference -eq 'Continue')" -Verb RunAs
        } catch {
            Write-Host "Failed to download or launch: $($_.Exception.Message)" -ForegroundColor Red
        }
        exit
    }
}

try {
    Write-Host "Loading main script..." -ForegroundColor Gray
    Write-Verbose "Downloading from: $Base/proav-shoko_powershell.ps1"
    $script = Invoke-RestMethod "$Base/proav-shoko_powershell.ps1"
    
    if ($VerbosePreference -eq 'Continue') {
        Write-Verbose "Executing main script with verbose output"
        Invoke-Expression $script
    } else {
        Invoke-Expression $script
    }
} catch {
    Write-Host "Failed to load main script: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Try manually: irm $Base/proav-shoko_powershell.ps1 | iex" -ForegroundColor Yellow
    pause
}
