<#
.SYNOPSIS
    ProAV Shoko Launcher - USB + Display Diagnostic Tool
.DESCRIPTION
    Downloads and runs the main ProAV Shoko diagnostic script from GitHub.
    Matches the behavior of python run.py --cli.
    Defaults to elevated (admin) mode for maximum data collection.
.EXAMPLE
    .\proav-shoko.ps1
    Run from main branch (default elevated)
.EXAMPLE
    .\proav-shoko.ps1 -Branch dev
    Run from dev branch
.EXAMPLE
    irm https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.ps1 | iex
    Run via one-liner
#>

# --- Configuration (change before running) ---
$Branch = "main"
$CsvPath = ""
# --------------------------------------------
# Parse command-line arguments (works with both iex and -File)
for ($i = 0; $i -lt $args.Count; $i += 2) {
    if ($args[$i] -eq "-Branch" -and ($i + 1) -lt $args.Count) { $Branch = $args[$i + 1] }
    if ($args[$i] -eq "-CsvPath" -and ($i + 1) -lt $args.Count) { $CsvPath = $args[$i + 1] }
}

$Repo = "klangche/klangche-proav-shoko"
$RefBase = "https://raw.githubusercontent.com/$Repo/refs/heads/$Branch"

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
    $elevate = Read-Host "Run with administrator privileges? [Y/n]"
    if ($elevate -eq '' -or $elevate -match '^[Yy]') {
        $temp = "$env:TEMP\shoko-elevated.ps1"
        Write-Verbose "Downloading main script to: $temp"
        try {
            $ps1Args = @("-Branch", $Branch)
            if ($CsvPath) { $ps1Args += "-CsvPath", $CsvPath }
            $argumentString = $ps1Args -join ' '
            Invoke-RestMethod "$RefBase/proav-shoko_powershell.ps1" | Out-File $temp -Encoding UTF8
            Write-Verbose "Launching elevated PowerShell"
            Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$temp`" $argumentString" -Verb RunAs
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
    $env:SHOKO_SKIP_ADMIN_CHECK = "1"
    $script = @(Invoke-RestMethod "$RefBase/proav-shoko_powershell.ps1") -join "`n"
    $script = $script.TrimStart([char]0xFEFF)
    Invoke-Expression $script
} catch {
    Write-Host "Failed to load main script: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Try manually: irm https://raw.githubusercontent.com/klangche/klangche-proav-shoko/refs/heads/$Branch/proav-shoko_powershell.ps1 | iex" -ForegroundColor Yellow
    pause
}
