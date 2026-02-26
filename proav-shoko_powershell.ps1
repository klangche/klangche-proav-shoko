# =============================================================================
# Shōko - USB + Display Diagnostic Tool
# Main PowerShell logic - updated to 0.0-10.0 scoring system
# =============================================================================

[CmdletBinding()]
param()

# =============================================================================
# CONFIG & HELPERS
# =============================================================================

function Get-Configuration {
    $jsonPath = "$PSScriptRoot\proav-shoko.json"
    if (Test-Path $jsonPath) {
        return Get-Content $jsonPath -Raw | ConvertFrom-Json
    } else {
        # Fallback - download from repo if local missing
        $url = "https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.json"
        return (Invoke-RestMethod $url)
    }
}

function Get-Color {
    param($ColorName)
    switch ($ColorName.ToLower()) {
        "cyan"   { [ConsoleColor]::Cyan }
        "magenta"{ [ConsoleColor]::Magenta }
        "yellow" { [ConsoleColor]::Yellow }
        "green"  { [ConsoleColor]::Green }
        "gray"   { [ConsoleColor]::Gray }
        "white"  { [ConsoleColor]::White }
        "red"    { [ConsoleColor]::Red }
        default  { [ConsoleColor]::White }
    }
}

function Format-Score {
    param([double]$Score)
    $clamped = [Math]::Max(0.0, [Math]::Min(10.0, $Score))
    return $clamped.ToString("N1")
}

function Format-Duration {
    param([TimeSpan]$ts)
    "{0:hh\:mm\:ss}" -f $ts
}

# =============================================================================
# PLATFORM STABILITY - updated to new 10.0 base
# =============================================================================

function Get-PlatformStability {
    param($Config, $Usb)

    Write-Verbose "Calculating platform stability - maxHops: $($Usb.MaxHops)"

    $referenceOutput = ""
    $additionalOutput = ""
    $worstReferenceScore = 10.0

    # Reference models (affect final score)
    foreach ($model in $Config.referenceModels) {
        $base = 10.0 - $Usb.MaxHops
        $score = [Math]::Max($Config.scoring.minScore, [Math]::Min($Config.scoring.maxScore, $base))

        $status = if ($score -ge $Config.scoring.thresholds.stable) { "STABLE" }
                  elseif ($score -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" }
                  else { "NOT STABLE" }

        $referenceOutput += "$($score.ToString('N1'))/10.0 $($model.name) $status`n"
        
        if ($score -lt $worstReferenceScore) { $worstReferenceScore = $score }
    }

    # Additional models (display only)
    foreach ($model in $Config.additionalModels) {
        $base = 10.0 - $Usb.MaxHops
        $score = [Math]::Max($Config.scoring.minScore, [Math]::Min($Config.scoring.maxScore, $base))

        $status = if ($score -ge $Config.scoring.thresholds.stable) { "STABLE" }
                  elseif ($score -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" }
                  else { "NOT STABLE" }

        $additionalOutput += "$($score.ToString('N1'))/10.0 $($model.name) $status`n"
    }

    $verdict = if ($worstReferenceScore -ge $Config.scoring.thresholds.stable) { "STABLE" }
               elseif ($worstReferenceScore -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" }
               else { "NOT STABLE" }

    return [PSCustomObject]@{
        ReferenceOutput  = $referenceOutput.Trim()
        AdditionalOutput = $additionalOutput.Trim()
        WorstScore       = $worstReferenceScore
        Verdict          = $verdict
    }
}

# =============================================================================
# MAIN REPORT DISPLAY (basic version)
# =============================================================================

function Show-Report {
    param($Config, $System, $Usb, $Display, $Stability)

    Clear-Host
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host "Shōko - USB + Display Diagnostic Tool v$($Config.version)" -ForegroundColor Cyan
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host $System.Mode -ForegroundColor Yellow
    Write-Host "Host: $($System.OSVersion) | PowerShell $($System.PSVersion)" -ForegroundColor Gray
    Write-Host "Arch: $($System.Architecture) | Current: $($System.CurrentPlatform)" -ForegroundColor Gray
    Write-Host ""

    Write-Host "USB TREE" -ForegroundColor Cyan
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host $Usb.Tree
    Write-Host "Max hops: $($Usb.MaxHops) | Tiers: $($Usb.Tiers) | Devices: $($Usb.Devices) | Hubs: $($Usb.Hubs)" -ForegroundColor Gray
    Write-Host ""

    Write-Host "DISPLAY TREE" -ForegroundColor Magenta
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host $Display

    Write-Host "STABILITY PER PLATFORM" -ForegroundColor Cyan
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host "Reference models (affects score):" -ForegroundColor White
    Write-Host $Stability.ReferenceOutput
    Write-Host "Additional models (reference only):" -ForegroundColor White
    Write-Host $Stability.AdditionalOutput
    Write-Host "==============================================================================" -ForegroundColor Cyan

    $color = Get-Color $Config.colors.($Stability.Verdict.ToLower() -replace ' ','')
    Write-Host ""
    Write-Host "HOST SUMMARY: $(Format-Score $Stability.WorstScore)/10.0 - $($Stability.Verdict)" -ForegroundColor $color
    Write-Host ""
}

# =============================================================================
# HTML REPORT (simplified - add your full logic here)
# =============================================================================

function Save-HtmlReport {
    param($Config, $System, $Usb, $Display, $Stability, $Analytics = $null)

    $date = Get-Date -Format "yyyyMMdd-HHmmss"
    $path = "$env:TEMP\shoko-report-$date.html"

    # Your existing HTML generation logic here...
    # For brevity: just a placeholder showing score format
    $html = @"
<html><body style='background:#000;color:#0f0;font-family:Consolas;'>
<pre>
Shōko Report - $date

HOST SUMMARY: $(Format-Score $Stability.WorstScore)/10.0 - $($Stability.Verdict)
</pre></body></html>
"@

    $html | Out-File $path -Encoding utf8
    Start-Process $path
}

# =============================================================================
# ANALYTICS SESSION (placeholder - insert your full monitoring code)
# =============================================================================

function Start-AnalyticsSession {
    param($Config, $System, $Usb, $Display, $Stability)

    # Your existing monitoring loop here...
    # For this example we simulate results

    $initialScore = $Stability.WorstScore
    $initialVerdict = $Stability.Verdict

    # Simulated deductions
    $deductions = 5.0
    $adjustedScore = [Math]::Max(0.0, $initialScore - $deductions)
    $adjustedVerdict = "NOT STABLE"   # calculate properly in real code

    return [PSCustomObject]@{
        InitialScore     = $initialScore
        InitialVerdict   = $initialVerdict
        AdjustedScore    = $adjustedScore
        AdjustedVerdict  = $adjustedVerdict
        Deductions       = $deductions
        Log              = @("16:22.004 – Started", "...")
    }
}

# =============================================================================
# FINAL REPORT AFTER ANALYTICS - with your exact SYSTEM STATUS layout
# =============================================================================

function Show-FinalReport {
    param($Config, $System, $InitialData, $Stability, $Analytics)

    Clear-Host

    # Reprint initial report trees etc. (your existing code)

    Write-Host "Points deducted: -$([math]::Round($Analytics.Deductions,1))" -ForegroundColor Yellow
    # ... your penalty breakdown here ...

    Write-Host ""
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host "SYSTEM STATUS" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "HOST SUMMARY (initial):  $(Format-Score $Analytics.InitialScore)/10.0 - $($Analytics.InitialVerdict)" -ForegroundColor Gray
    Write-Host "HOST SUMMARY (adjusted): $(Format-Score $Analytics.AdjustedScore)/10.0 - $($Analytics.AdjustedVerdict)" -ForegroundColor Red   # adjust color dynamically
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Analytics log excerpt:" -ForegroundColor Cyan
    $Analytics.Log | Select-Object -Last 8 | ForEach-Object { Write-Host $_ }

    Write-Host ""
    Write-Host "Open HTML report with full data? (y/n)" -ForegroundColor Green
}

# =============================================================================
# MAIN FLOW
# =============================================================================

$Config = Get-Configuration
$System = Get-SystemInfo               # your function
$Usb    = Get-UsbTree -Config $Config  # your function
$Display = Get-DisplayTree -Config $Config

$Stability = Get-PlatformStability -Config $Config -Usb $Usb

Show-Report -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability

$viewHtml = Read-Host "View result in browser? (y/n)"
if ($viewHtml -match '^[Yy]') { Save-HtmlReport ... }

$runAnalytics = Read-Host "Run Analytics (y/n)"
if ($runAnalytics -match '^[Yy]') {
    $Analytics = Start-AnalyticsSession -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability
    Show-FinalReport -Config $Config -System $System -InitialData $null -Stability $Stability -Analytics $Analytics
}

Read-Host "`nFinished. Press Enter to exit"
