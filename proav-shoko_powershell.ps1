# =============================================================================
# PLATFORM STABILITY
# =============================================================================

function Get-PlatformStability {
    <#
    .SYNOPSIS
        Calculate stability for reference and additional models
    #>
    param($Config, $Tiers)
    
    Write-Verbose "Calculating platform stability for $Tiers tiers"
    
    $referenceOutput = ""
    $additionalOutput = ""
    $referenceScores = @()
    $worstReferenceScore = 10
    
    # Process reference models (affect score)
    foreach ($model in $Config.referenceModels) {
        $rec = $model.rec
        $max = $model.max
        $name = $model.name
        
        $status = if ($Tiers -le $rec) { "STABLE" } 
                  elseif ($Tiers -le $max) { "POTENTIALLY UNSTABLE" } 
                  else { "NOT STABLE" }
        
        $modelScore = 9 - ($Tiers - 1)
        if ($modelScore -lt $Config.scoring.minScore) { $modelScore = $Config.scoring.minScore }
        if ($modelScore -gt $Config.scoring.maxScore) { $modelScore = $Config.scoring.maxScore }
        
        $referenceOutput += Format-PlatformLine -tiers $Tiers -max $max -name $name -status $status
        $referenceOutput += "`n"
        
        $referenceScores += $modelScore
        if ($modelScore -lt $worstReferenceScore) { $worstReferenceScore = $modelScore }
    }
    
    # Process additional models (reference only)
    foreach ($model in $Config.additionalModels) {
        $rec = $model.rec
        $max = $model.max
        $name = $model.name
        
        $status = if ($Tiers -le $rec) { "STABLE" } 
                  elseif ($Tiers -le $max) { "POTENTIALLY UNSTABLE" } 
                  else { "NOT STABLE" }
        
        $additionalOutput += Format-PlatformLine -tiers $Tiers -max $max -name $name -status $status
        $additionalOutput += "`n"
    }
    
    if ($referenceScores.Count -eq 0) {
        $worstReferenceScore = 9 - ($Tiers - 1)
    }
    
    $verdict = if ($worstReferenceScore -ge $Config.scoring.thresholds.stable) { "STABLE" } 
               elseif ($worstReferenceScore -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" } 
               else { "NOT STABLE" }
    
    return [PSCustomObject]@{
        ReferenceOutput = $referenceOutput
        AdditionalOutput = $additionalOutput
        WorstScore = $worstReferenceScore
        Verdict = $verdict
    }
}

# =============================================================================
# REPORTING
# =============================================================================

function Show-Report {
    <#
    .SYNOPSIS
        Display the main report
    #>
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
    
    $verdictColor = Get-Color $Config.colors.($Stability.Verdict.ToLower().Replace(' ', ''))
    Write-Host "HOST SUMMARY: $(Format-Score $Stability.WorstScore)/10 - $($Stability.Verdict)" -ForegroundColor $verdictColor
    Write-Host ""
}

function Save-HtmlReport {
    <#
    .SYNOPSIS
        Save report as HTML file and open in browser
    #>
    param($Config, $System, $Usb, $Display, $Stability, $Analytics = $null)
    
    $dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $outHtml = "$env:TEMP\shoko-report-$dateStamp.html"
    
    $analyticsSection = ""
    if ($Analytics) {
        $analyticsSection = @"

HOST SUMMARY: $(Format-Score $Analytics.InitialScore)/10 - $($Analytics.InitialVerdict)
HOST SUMMARY: $(Format-Score $Analytics.AdjustedScore)/10 - $($Analytics.AdjustedVerdict) (adjusted)

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
        $analyticsSection = "HOST SUMMARY: $(Format-Score $Stability.WorstScore)/10 - $($Stability.Verdict)"
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
        Write-Verbose "HTML report saved to: $outHtml"
        Start-Process $outHtml
    } catch {
        Write-Verbose "Failed to save/open HTML: $_"
        Write-Host "Report saved to: $outHtml" -ForegroundColor Gray
    }
}

# =============================================================================
# ANALYTICS
# =============================================================================

function Start-AnalyticsSession {
    <#
    .SYNOPSIS
        Run real-time monitoring of USB/display events
    #>
    param($Config, $System, $Usb, $Display, $Stability)
    
    Write-Verbose "Starting analytics session"
    
    # Store initial data
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
    
    Clear-Host
    
    # Analytics header
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
    
    # Analytics counters
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
    
    # Store cursor positions for efficient updates
    $statsLine = 6
    $logStartLine = $statsLine + 12
    
    while (-not $Host.UI.RawUI.KeyAvailable) {
        $elapsed = (Get-Date) - $startTime
        $duration = Format-Duration $elapsed
        
        # Update only stats lines (non-aggressive)
        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $statsLine
        Write-Host "Duration: $duration" -ForegroundColor (Get-Color $Config.colors.green)
        
        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($statsLine + 1)
        Write-Host "Total events logged: $($counters.total)" -ForegroundColor (Get-Color $Config.colors.white)
        
        $line = $statsLine + 2
        if ($System.IsAdmin) {
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $line; Write-Host "USB RE-HANDSHAKES: $($counters.rehandshakes)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 1); Write-Host "USB JITTER: $($counters.jitter)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 2); Write-Host "USB CRC ERRORS: $($counters.crcErrors)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 3); Write-Host "USB BUS RESETS: $($counters.busResets)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 4); Write-Host "USB OVERCURRENT: $($counters.overcurrent)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 5); Write-Host "DISPLAY HOTPLUGS: $($counters.hotplugs)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 6); Write-Host "DISPLAY EDID ERRORS: $($counters.edidErrors)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 7); Write-Host "DISPLAY LINK FAILURES: $($counters.linkFailures)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 8); Write-Host "OTHER ERRORS: $($counters.otherErrors)          "
        } else {
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $line; Write-Host "USB CONNECTS: $($counters.connects)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 1); Write-Host "USB DISCONNECTS: $($counters.disconnects)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 2); Write-Host "USB RE-HANDSHAKES: $($counters.rehandshakes)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 3); Write-Host "USB JITTER: $($counters.jitter)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 4); Write-Host "USB ERRORS: $($counters.crcErrors + $counters.busResets + $counters.overcurrent + $counters.otherErrors)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 5); Write-Host "DISPLAY HOTPLUGS: $($counters.hotplugs)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 6); Write-Host "DISPLAY ERRORS: $($counters.edidErrors + $counters.linkFailures)          "
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    # Stop analytics
    $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    $stopTime = Get-Date
    $totalDuration = Format-Duration ($stopTime - $startTime)
    $analyticsLog += "$($stopTime.ToString('HH:mm:ss.fff')) - Logging ended (total duration: $totalDuration)"
    
    # Calculate deductions
    $deductions = 0
    if ($System.IsAdmin) {
        $deductions = ($counters.rehandshakes * $Config.scoring.penalties.rehandshake) +
                      ($counters.jitter * $Config.scoring.penalties.jitter) +
                      ($counters.crcErrors * $Config.scoring.penalties.crc) +
                      ($counters.busResets * $Config.scoring.penalties.busReset) +
                      ($counters.overcurrent * $Config.scoring.penalties.overcurrent) +
                      ($counters.hotplugs * $Config.scoring.penalties.hotplug) +
                      ($counters.edidErrors * $Config.scoring.penalties.edidError) +
                      ($counters.linkFailures * $Config.scoring.penalties.linkFailure) +
                      ($counters.otherErrors * $Config.scoring.penalties.otherError)
    } else {
        $deductions = ($counters.rehandshakes * $Config.scoring.penalties.rehandshake) +
                      ($counters.jitter * $Config.scoring.penalties.jitter) +
                      (($counters.crcErrors + $counters.busResets + $counters.overcurrent + $counters.otherErrors) * 0.5)
    }
    
    $adjustedScore = [Math]::Max(0, $initialData.Score - $deductions)
    $adjustedScore = [Math]::Min($adjustedScore, $Config.scoring.maxScore)
    
    $adjustedVerdict = if ($adjustedScore -ge $Config.scoring.thresholds.stable) { "STABLE" } 
                       elseif ($adjustedScore -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" } 
                       else { "NOT STABLE" }
    
    # Build summary text
    if ($System.IsAdmin) {
        $summaryText = @"
USB RE-HANDSHAKES: $($counters.rehandshakes)
USB JITTER: $($counters.jitter)
USB CRC ERRORS: $($counters.crcErrors)
USB BUS RESETS: $($counters.busResets)
USB OVERCURRENT: $($counters.overcurrent)
DISPLAY HOTPLUGS: $($counters.hotplugs)
DISPLAY EDID ERRORS: $($counters.edidErrors)
DISPLAY LINK FAILURES: $($counters.linkFailures)
OTHER ERRORS: $($counters.otherErrors)
"@
    } else {
        $summaryText = @"
USB CONNECTS: $($counters.connects)
USB DISCONNECTS: $($counters.disconnects)
USB RE-HANDSHAKES: $($counters.rehandshakes)
USB JITTER: $($counters.jitter)
USB ERRORS: $($counters.crcErrors + $counters.busResets + $counters.overcurrent + $counters.otherErrors)
DISPLAY HOTPLUGS: $($counters.hotplugs)
DISPLAY ERRORS: $($counters.edidErrors + $counters.linkFailures)
"@
    }
    
    return [PSCustomObject]@{
        InitialData = $initialData
        Counters = $counters
        Deductions = $deductions
        AdjustedScore = $adjustedScore
        AdjustedVerdict = $adjustedVerdict
        Log = $analyticsLog
        SummaryText = $summaryText
        LogText = $analyticsLog -join "`n"
        InitialScore = $initialData.Score
        InitialVerdict = $initialData.Verdict
    }
}

function Show-FinalReport {
    <#
    .SYNOPSIS
        Display final report after analytics session
    #>
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
    Write-Host "Max hops: $($InitialData.MaxHops) | Tiers: $($InitialData.Tiers) | Devices: $($InitialData.Devices) | Hubs: $($InitialData.Hubs)" -ForegroundColor (Get-Color $Config.colors.gray)
    Write-Host ""
    
    Write-Host "DISPLAY TREE" -ForegroundColor (Get-Color $Config.colors.magenta)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $InitialData.Display
    
    Write-Host "STABILITY PER PLATFORM" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Reference models (affects score):" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $InitialData.ReferenceOutput
    Write-Host "Additional models (reference only):" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $InitialData.AdditionalOutput
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host ""
    
    $verdictColor = Get-Color $Config.colors.($Stability.Verdict.ToLower().Replace(' ', ''))
    $adjustedColor = Get-Color $Config.colors.($Analytics.AdjustedVerdict.ToLower().Replace(' ', ''))
    
    Write-Host "HOST SUMMARY: $(Format-Score $InitialData.Score)/10 - $($InitialData.Verdict)" -ForegroundColor $verdictColor
    Write-Host "HOST SUMMARY: $(Format-Score $Analytics.AdjustedScore)/10 - $($Analytics.AdjustedVerdict) (adjusted)" -ForegroundColor $adjustedColor
    Write-Host ""
    
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Analytics Summary (during monitoring):" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Total events logged: $($Analytics.Counters.total)" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $Analytics.SummaryText
    Write-Host "Points deducted: -$([math]::Round($Analytics.Deductions,1))"
    Write-Host ""
    
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Analytics Log (during monitoring):" -ForegroundColor (Get-Color $Config.colors.cyan)
    foreach ($line in $Analytics.Log) {
        Write-Host $line
    }
}

# =============================================================================
# MAIN
# =============================================================================

function Main {
    <#
    .SYNOPSIS
        Main execution flow
    #>
    Write-Verbose "Starting Shōko main script"
    
    $Config = Get-Configuration
    $System = Get-SystemInfo
    
    Write-Host "`nCollecting system data..." -ForegroundColor Gray
    
    $Usb = Get-UsbTree -Config $Config
    $Display = Get-DisplayTree -Config $Config
    $Stability = Get-PlatformStability -Config $Config -Tiers $Usb.Tiers
    
    Show-Report -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability
    
    # Question 2 - HTML report
    $htmlChoice = Read-Host "Open HTML report? (y/n)"
    if ($htmlChoice -match '^[Yy]') {
        Save-HtmlReport -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability
    }
    
    # Question 3 - Analytics
    $analyticsChoice = Read-Host "Run deep analytics session? (y/n)"
    if ($analyticsChoice -match '^[Yy]') {
        $Analytics = Start-AnalyticsSession -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability
        
        Show-FinalReport -Config $Config -System $System -InitialData $Analytics.InitialData -Stability $Stability -Analytics $Analytics
        
        # Question 4 - HTML report with full data
        $finalHtmlChoice = Read-Host "`nOpen HTML report with full data? (y/n)"
        if ($finalHtmlChoice -match '^[Yy]') {
            Save-HtmlReport -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability -Analytics $Analytics
        }
    }
    
    Write-Host "`nShōko finished. Press Enter to close." -ForegroundColor (Get-Color $Config.colors.green)
    Read-Host
}

# Run main
Main
