<#
.SYNOPSIS
    Shoko Launcher - USB + Display Diagnostic Tool
.DESCRIPTION
    Downloads and runs the main Shoko diagnostic script from GitHub.
    Supports both direct execution and "irm | iex" one-liner.
#>

# --- Configuration (change before running) ---
$CsvPath = ""
# --------------------------------------------

$Repo = "klangche/klangche-proav-shoko"
$Branch = "dev"
$Base = "https://raw.githubusercontent.com/$Repo/$Branch"

Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "Shoko - USB + Display Diagnostic Tool Launcher" -ForegroundColor Cyan
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
    Write-Host "Limited mode - full features require Administrator rights." -ForegroundColor Yellow
    $elevate = Read-Host "Run with administrator privileges? (y/n)"
    if ($elevate -match '^[Yy]') {
        $temp = "$env:TEMP\shoko-elevated.ps1"
        Write-Verbose "Downloading main script to: $temp"
        try {
            $ps1Args = @()
            if ($CsvPath) { $ps1Args += "-CsvPath", $CsvPath }
            $argumentString = $ps1Args -join ' '
            Invoke-RestMethod "$Base/proav-shoko_powershell.ps1" | Out-File $temp -Encoding UTF8
            Write-Verbose "Launching elevated PowerShell"
            Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `$temp $argumentString" -Verb RunAs
        } catch {
            Write-Host "Failed to download or launch: $($_.Exception.Message)" -ForegroundColor Red
        }
        exit
    } else {
        Write-Host "Continuing in basic mode..." -ForegroundColor Gray
    }
}

try {
    Write-Host "Loading main script..." -ForegroundColor Gray
    $script = @(Invoke-RestMethod "$Base/proav-shoko_powershell.ps1") -join "`n"
    Invoke-Expression $script
} catch {
    Write-Host "Failed to load main script: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Try manually: irm $Base/proav-shoko_powershell.ps1 | iex" -ForegroundColor Yellow
    pause
}
