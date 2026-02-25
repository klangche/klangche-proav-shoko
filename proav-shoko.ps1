# proav-shoko.ps1 - Shōko Launcher (works in PowerShell 5.1 and 7+)

$Repo = "klangche/klangche-proav-shoko"
$Branch = "main"
$Base = "https://raw.githubusercontent.com/$Repo/$Branch"

Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "Shōko – USB + Display Diagnostic Tool" -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host ""

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Limited mode – full features require Administrator rights." -ForegroundColor Yellow
    $elevate = Read-Host "Elevate now? (y/n)"
    if ($elevate -match '^[Yy]') {
        $temp = "$env:TEMP\shoko-elevated.ps1"
        Invoke-RestMethod "$Base/proav-shoko_powershell.ps1" | Out-File $temp -Encoding UTF8
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$temp`"" -Verb RunAs
        exit
    }
}

try {
    Write-Host "Loading main script..." -ForegroundColor Gray
    Invoke-Expression (Invoke-RestMethod "$Base/proav-shoko_powershell.ps1")
} catch {
    Write-Host "Failed to load main script: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Try manually: irm $Base/proav-shoko_powershell.ps1 | iex" -ForegroundColor Yellow
    pause
}
