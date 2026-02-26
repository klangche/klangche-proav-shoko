# =============================================================================
# Shōko - USB + Display Diagnostic Tool   (updated 2026-02-26)
# Scoring: 10.0 - maxHops, one decimal place, SYSTEM STATUS block
# =============================================================================

# =============================================================================
# HELPERS
# =============================================================================

function Get-Color {
    param($ColorName)
    switch ($ColorName.ToLower()) {
        "cyan"    { [ConsoleColor]::Cyan }
        "magenta" { [ConsoleColor]::Magenta }
        "yellow"  { [ConsoleColor]::Yellow }
        "green"   { [ConsoleColor]::Green }
        "gray"    { [ConsoleColor]::Gray }
        "white"   { [ConsoleColor]::White }
        "red"     { [ConsoleColor]::Red }
        default   { [ConsoleColor]::White }
    }
}

function Format-Score {
    param([double]$Score)
    $clamped = [Math]::Max(0.0, [Math]::Min(10.0, $Score))
    return $clamped.ToString("N1")
}

function Format-Duration {
    param([TimeSpan]$Duration)
    return $Duration.ToString("hh\:mm\:ss")
}

# =============================================================================
# PLATFORM STABILITY – updated to new scoring
# =============================================================================

function Get-PlatformStability {
    param($Config, $Usb)

    Write-Verbose "Calculating platform stability - maxHops: $($Usb.MaxHops)"

    $referenceOutput = ""
    $additionalOutput = ""
    $worstReferenceScore = 10.0

    # Reference models (affect final score)
    foreach ($model in $Config.referenceModels) {
        $baseScore = 10.0 - $Usb.MaxHops
        $score = [Math]::Max($Config.scoring.minScore, [Math]::Min($Config.scoring.maxScore, $baseScore))

        $status = if ($score -ge $Config.scoring.thresholds.stable) { "STABLE" }
                  elseif ($score -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" }
                  else { "NOT STABLE" }

        $referenceOutput += "$($score.ToString('N1'))/10.0 $($model.name) $status`n"

        if ($score -lt $worstReferenceScore) { $worstReferenceScore = $score }
    }

    # Additional models (display only)
    foreach ($model in $Config.additionalModels) {
        $baseScore = 10.0 - $Usb.MaxHops
        $score = [Math]::Max($Config.scoring.minScore, [Math]::Min($Config.scoring.maxScore, $baseScore))

        $status = if ($score -ge $Config.scoring.thresholds.stable) { "STABLE" }
                  elseif ($score -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" }
                  else { "NOT STABLE" }

        $additionalOutput += "$($score.ToString('N1'))/10.0 $($model.name) $status`n"
    }

    $verdict = if ($worstReferenceScore -ge $Config.scoring.thresholds.stable) { "STABLE" }
               elseif ($worstReferenceScore -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" }
               else { "NOT STABLE" }

    return [PSCustomObject]@{
        ReferenceOutput  = $referenceOutput.TrimEnd("`n")
        AdditionalOutput = $additionalOutput.TrimEnd("`n")
        WorstScore       = $worstReferenceScore
        Verdict          = $verdict
    }
}

# =============================================================================
# REPORTING
# =============================================================================

function Show-Report {
    param($Config, $System, $Usb, $Display, $Stability)

    Clear-Host
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Shōko - USB + Display Diagnostic Tool v$($Config.version)" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $System.Mode -ForegroundColor (Get-Color $Config.colors.yellow)
    Write-Host "Host: $($System.OSVersion) | PowerShell $($System.PSVersion)" -ForegroundColor (Get-Color $Config.colors.gray)
    Write-Host "Arch: $($System.Architecture) | Current: $($System.CurrentPlatform)" -ForegroundColor (Get-Color $Config.colors.gray)
    Write-Host ""

    Write-Host "USB TREE" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $Usb.Tree
    Write-Host "Max hops: $($Usb.MaxHops) | Tiers: $($Usb.Tiers) | Devices: $($Usb.Devices) | Hubs: $($Usb.Hubs)" -ForegroundColor (Get-Color $Config.colors.gray)
    Write-Host ""

    Write-Host "DISPLAY TREE" -ForegroundColor (Get-Color $Config.colors.magenta)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $Display

    Write-Host "STABILITY PER PLATFORM" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Reference models (affects score):" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $Stability.ReferenceOutput
    Write-Host "Additional models (reference only):" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $Stability.AdditionalOutput
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host ""

    $verdictColor = Get-Color $Config.colors.($Stability.Verdict.ToLower() -replace ' ','')
    Write-Host "HOST SUMMARY: $(Format-Score $Stability.WorstScore)/10.0 - $($Stability.Verdict)" -ForegroundColor $verdictColor
    Write-Host ""
}

function Save-HtmlReport {
    param($Config, $System, $Usb, $Display, $Stability, $Analytics = $null)

    $dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $outHtml = "$env:TEMP\shoko-report-$dateStamp.html"

    $analyticsSection = ""
    if ($Analytics) {
        $analyticsSection = @"

HOST SUMMARY (initial): $(Format-Score $Analytics.InitialScore)/10.0 - $($Analytics.InitialVerdict)
HOST SUMMARY (adjusted): $(Format-Score $Analytics.AdjustedScore)/10.0 - $($Analytics.AdjustedVerdict)

==============================================================================
Analytics Summary (during monitoring):
Total events logged: $($Analytics.Counters.total)
$($Analytics.SummaryText)
Points deducted: -$([math]::Round($Analytics.Deductions,1))

==============================================================================
Analytics Log (during monitoring):
$($Analytics.LogText)
"@
    } else {
        $analyticsSection = "HOST SUMMARY: $(Format-Score $Stability.WorstScore)/10.0 - $($Stability.Verdict)"
    }

    $htmlContent = @"
<html>
<head><title>Shōko Report $dateStamp</title></head>
<body style='background:#000;color:#0f0;font-family:Consolas;'>
<pre>
Shōko Report - $dateStamp

$($System.Mode)
Host: $($System.OSVersion) | PowerShell $($System.PSVersion)
Arch: $($System.Architecture) | Current: $($System.CurrentPlatform)

USB TREE
==============================================================================
$($Usb.Tree)
Max hops: $($Usb.MaxHops) | Tiers: $($Usb.Tiers) | Devices: $($Usb.Devices) | Hubs: $($Usb.Hubs)

DISPLAY TREE
==============================================================================
$Display
STABILITY PER PLATFORM
==============================================================================
Reference models (affects score):
$($Stability.ReferenceOutput)
Additional models (reference only):
$($Stability.AdditionalOutput)
==============================================================================

$analyticsSection
</pre></body></html>
"@

    try {
        $htmlContent | Out-File $outHtml -Encoding UTF8
        Start-Process $outHtml
    } catch {
        Write-Host "Failed to save/open HTML: $_" -ForegroundColor Red
    }
}

# =============================================================================
# ANALYTICS – waits for key press
# =============================================================================

function Start-AnalyticsSession {
    param($Config, $System, $Usb, $Display, $Stability)

    Clear-Host
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Shoko Analytics" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Monitoring connections... Press any key to stop" -ForegroundColor (Get-Color $Config.colors.gray)
    if ($System.IsAdmin) {
        Write-Host "Elevated mode - full device details" -ForegroundColor (Get-Color $Config.colors.green)
    } else {
        Write-Host "Basic mode - device IDs visible, friendly names unavailable" -ForegroundColor (Get-Color $Config.colors.yellow)
    }
    Write-Host ""

    $initialData = [PSCustomObject]@{
        Tree = $Usb.Tree
        Display = $Display
        ReferenceOutput = $Stability.ReferenceOutput
        AdditionalOutput = $Stability.AdditionalOutput
        Score = $Stability.WorstScore
        Verdict = $Stability.Verdict
        Devices = $Usb.Devices
        Hubs = $Usb.Hubs
        MaxHops = $Usb.MaxHops
        Tiers = $Usb.Tiers
    }

    $analyticsLog = @()
    $startTime = Get-Date
    $analyticsLog += "$($startTime.ToString('HH:mm:ss.fff')) - Logging started"

    $counters = @{
        total = 0
        connects = 0
        disconnects = 0
        rehandshakes = 0
        jitter = 0
        crcErrors = 0
        busResets = 0
        overcurrent = 0
        hotplugs = 0
        edidErrors = 0
        linkFailures = 0
        otherErrors = 0
    }

    while (-not $Host.UI.RawUI.KeyAvailable) {
        $elapsed = (Get-Date) - $startTime
        $duration = Format-Duration $elapsed

        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, 6
        Write-Host "Duration: $duration" -ForegroundColor (Get-Color $Config.colors.green) -NoNewline

        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, 7
        Write-Host "Total events logged: $($counters.total)" -ForegroundColor (Get-Color $Config.colors.white) -NoNewline

        # Placeholder – replace with real event detection logic
        Start-Sleep -Milliseconds 800
        $counters.total++
    }

    # Consume key press
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

    $stopTime = Get-Date
    $analyticsLog += "$($stopTime.ToString('HH:mm:ss.fff')) - Logging ended (total duration: $(Format-Duration ($stopTime - $startTime)))"

    # Calculate deductions (placeholder values – replace with real counter logic)
    $deductions = ($counters.rehandshakes * $Config.scoring.penalties.rehandshake) +
                  ($counters.jitter * $Config.scoring.penalties.jitter) +
                  ($counters.crcErrors * $Config.scoring.penalties.crc) +
                  ($counters.busResets * $Config.scoring.penalties.busReset) +
                  ($counters.overcurrent * $Config.scoring.penalties.overcurrent) +
                  ($counters.hotplugs * $Config.scoring.penalties.hotplug) +
                  ($counters.edidErrors * $Config.scoring.penalties.edidError) +
                  ($counters.linkFailures * $Config.scoring.penalties.linkFailure) +
                  ($counters.otherErrors * $Config.scoring.penalties.otherError)

    $adjustedScore = [Math]::Max(0.0, $initialData.Score - $deductions)
    $adjustedScore = [Math]::Min(10.0, $adjustedScore)

    $adjustedVerdict = if ($adjustedScore -ge $Config.scoring.thresholds.stable) { "STABLE" }
                       elseif ($adjustedScore -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" }
                       else { "NOT STABLE" }

    return [PSCustomObject]@{
        InitialData      = $initialData
        Counters         = $counters
        Deductions       = $deductions
        AdjustedScore    = $adjustedScore
        AdjustedVerdict  = $adjustedVerdict
        Log              = $analyticsLog
        InitialScore     = $initialData.Score
        InitialVerdict   = $initialData.Verdict
    }
}

# =============================================================================
# FINAL REPORT AFTER ANALYTICS
# =============================================================================

function Show-FinalReport {
    param($Config, $System, $InitialData, $Stability, $Analytics)

    Clear-Host

    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Shōko - USB + Display Diagnostic Tool v$($Config.version)" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $System.Mode -ForegroundColor (Get-Color $Config.colors.yellow)
    Write-Host "Host: $($System.OSVersion) | PowerShell $($System.PSVersion)" -ForegroundColor (Get-Color $Config.colors.gray)
    Write-Host "Arch: $($System.Architecture) | Current: $($System.CurrentPlatform)" -ForegroundColor (Get-Color $Config.colors.gray)
    Write-Host ""

    Write-Host "USB TREE" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $InitialData.Tree
    Write-Host "Max hops: $($InitialData.MaxHops) | Tiers: $($InitialData.Tiers) | Devices: $($InitialData.Devices) | Hubs: $($InitialData.Hubs)" -ForegroundColor Gray
    Write-Host ""

    Write-Host "DISPLAY TREE" -ForegroundColor (Get-Color $Config.colors.magenta)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $InitialData.Display

    Write-Host "STABILITY PER PLATFORM" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Reference models (affects score):" -ForegroundColor White
    Write-Host $InitialData.ReferenceOutput
    Write-Host "Additional models (reference only):" -ForegroundColor White
    Write-Host $InitialData.AdditionalOutput
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host ""

    $verdictColor = Get-Color $Config.colors.($Stability.Verdict.ToLower() -replace ' ','')
    $adjustedColor = Get-Color $Config.colors.($Analytics.AdjustedVerdict.ToLower() -replace ' ','')

    Write-Host "Points deducted: -$([math]::Round($Analytics.Deductions,1))" -ForegroundColor Yellow
    Write-Host ""

    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "SYSTEM STATUS" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host ""
    Write-Host "HOST SUMMARY (initial):  $(Format-Score $Analytics.InitialScore)/10.0 - $($Analytics.InitialVerdict)" -ForegroundColor Gray
    Write-Host "HOST SUMMARY (adjusted): $(Format-Score $Analytics.AdjustedScore)/10.0 - $($Analytics.AdjustedVerdict)" -ForegroundColor $adjustedColor
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host ""

    Write-Host "Analytics log excerpt:" -ForegroundColor (Get-Color $Config.colors.cyan)
    $Analytics.Log | Select-Object -Last 8 | ForEach-Object { Write-Host $_ }
}

# =============================================================================
# MAIN FLOW
# =============================================================================

function Main {
    $Config = Get-Configuration           # your existing function
    $System = Get-SystemInfo              # your existing function

    Write-Host "`nCollecting system data..." -ForegroundColor Gray

    $Usb = Get-UsbTree -Config $Config    # your existing function
    $Display = Get-DisplayTree -Config $Config

    $Stability = Get-PlatformStability -Config $Config -Usb $Usb

    Show-Report -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability

    $htmlChoice = Read-Host "Open HTML report? (y/n)"
    if ($htmlChoice -match '^[Yy]') {
        Save-HtmlReport -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability
    }

    $analyticsChoice = Read-Host "Run deep analytics session? (y/n)"
    if ($analyticsChoice -match '^[Yy]') {
        $Analytics = Start-AnalyticsSession -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability
        Show-FinalReport -Config $Config -System $System -InitialData $Analytics.InitialData -Stability $Stability -Analytics $Analytics

        $finalHtml = Read-Host "`nOpen HTML report with full data? (y/n)"
        if ($finalHtml -match '^[Yy]') {
            Save-HtmlReport -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability -Analytics $Analytics
        }
    }

    Write-Host "`nShōko finished. Press Enter to close." -ForegroundColor Green
    Read-Host
}

# Run
Main
