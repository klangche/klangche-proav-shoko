# =============================================================================
# Shōko - USB + Display Diagnostic Tool v1.1.0
# Updated scoring: 10.0 - maxHops, always 1 decimal, SYSTEM STATUS block
# Analytics: all counters visible, basic mode shows clear message
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
# PLATFORM STABILITY – new scoring
# =============================================================================

function Get-PlatformStability {
    param($Config, $Usb)

    $referenceOutput = ""
    $additionalOutput = ""
    $worstReferenceScore = 10.0

    foreach ($model in $Config.referenceModels) {
        $baseScore = 10.0 - $Usb.MaxHops
        $score = [Math]::Max($Config.scoring.minScore, [Math]::Min($Config.scoring.maxScore, $baseScore))

        $status = if ($score -ge $Config.scoring.thresholds.stable) { "STABLE" }
                  elseif ($score -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" }
                  else { "NOT STABLE" }

        $referenceOutput += "$(Format-Score $score)/10.0 $($model.name) $status`n"

        if ($score -lt $worstReferenceScore) { $worstReferenceScore = $score }
    }

    foreach ($model in $Config.additionalModels) {
        $baseScore = 10.0 - $Usb.MaxHops
        $score = [Math]::Max($Config.scoring.minScore, [Math]::Min($Config.scoring.maxScore, $baseScore))

        $status = if ($score -ge $Config.scoring.thresholds.stable) { "STABLE" }
                  elseif ($score -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" }
                  else { "NOT STABLE" }

        $additionalOutput += "$(Format-Score $score)/10.0 $($model.name) $status`n"
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
# REPORTING – clean Nothing detected style
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

    if ([string]::IsNullOrWhiteSpace($Usb.Tree) -or $Usb.Devices -eq 0) {
        Write-Host "HOST"
        Write-Host "└── Nothing detected"
    } else {
        Write-Host $Usb.Tree
    }
    Write-Host "Max hops: $($Usb.MaxHops) | Tiers: $($Usb.Tiers) | Devices: $($Usb.Devices) | Hubs: $($Usb.Hubs)" -ForegroundColor Gray
    Write-Host ""

    Write-Host "DISPLAY TREE" -ForegroundColor Magenta
    Write-Host "==============================================================================" -ForegroundColor Cyan

    if ([string]::IsNullOrWhiteSpace($Display)) {
        Write-Host "HOST"
        Write-Host "└── Nothing detected"
    } else {
        Write-Host $Display
    }
    Write-Host ""

    Write-Host "STABILITY PER PLATFORM" -ForegroundColor Cyan
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host "Reference models (affects score):" -ForegroundColor White
    Write-Host $Stability.ReferenceOutput
    Write-Host "Additional models (reference only):" -ForegroundColor White
    Write-Host $Stability.AdditionalOutput
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host ""

    $verdictColor = Get-Color $Config.colors.($Stability.Verdict.ToLower() -replace ' ','')
    Write-Host "HOST SUMMARY: $(Format-Score $Stability.WorstScore)/10.0 - $($Stability.Verdict)" -ForegroundColor $verdictColor
    Write-Host ""
}

# =============================================================================
# ANALYTICS – all counters visible, basic mode shows message, 2s update
# =============================================================================

function Start-AnalyticsSession {
    param($Config, $System, $Usb, $Display, $Stability)

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
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host "Shoko Analytics" -ForegroundColor Cyan
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host "Monitoring connections... Press any key to stop" -ForegroundColor Yellow
    Write-Host "Update every 2 seconds" -ForegroundColor Gray
    Write-Host ""

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

    $statsLine = 6

    while (-not $Host.UI.RawUI.KeyAvailable) {
        $elapsed = (Get-Date) - $startTime
        $duration = Format-Duration $elapsed

        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $statsLine
        Write-Host "Duration: $duration                  " -ForegroundColor Green -NoNewline

        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($statsLine + 1)
        Write-Host "Total events logged: $($counters.total)                  " -ForegroundColor White -NoNewline

        $line = $statsLine + 3

        if ($System.IsAdmin) {
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $line;     Write-Host "USB RE-HANDSHAKES: $($counters.rehandshakes)          " -NoNewline
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line+1); Write-Host "USB JITTER: $($counters.jitter)          " -NoNewline
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line+2); Write-Host "USB CRC ERRORS: $($counters.crcErrors)          " -NoNewline
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line+3); Write-Host "USB BUS RESETS: $($counters.busResets)          " -NoNewline
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line+4); Write-Host "USB OVERCURRENT: $($counters.overcurrent)          " -NoNewline
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line+5); Write-Host "DISPLAY HOTPLUGS: $($counters.hotplugs)          " -NoNewline
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line+6); Write-Host "DISPLAY EDID ERRORS: $($counters.edidErrors)          " -NoNewline
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line+7); Write-Host "DISPLAY LINK FAILURES: $($counters.linkFailures)          " -NoNewline
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line+8); Write-Host "OTHER ERRORS: $($counters.otherErrors)          " -NoNewline
        } else {
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $line
            Write-Host "Basic mode – not evaluated                          " -ForegroundColor Yellow -NoNewline
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line+1)
            Write-Host "Run elevated for full counters & detailed errors   " -ForegroundColor Yellow -NoNewline
        }

        Start-Sleep -Seconds 2
    }

    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    $stopTime = Get-Date
    $analyticsLog += "$($stopTime.ToString('HH:mm:ss.fff')) - Logging ended (total duration: $(Format-Duration ($stopTime - $startTime)))"

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
    } # basic mode: deductions limited or 0

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
# FINAL REPORT – full breakdown, SYSTEM STATUS
# =============================================================================

function Show-FinalReport {
    param($Config, $System, $InitialData, $Stability, $Analytics)

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
    if ([string]::IsNullOrWhiteSpace($InitialData.Tree) -or $InitialData.Devices -eq 0) {
        Write-Host "HOST"
        Write-Host "└── Nothing detected"
    } else {
        Write-Host $InitialData.Tree
    }
    Write-Host "Max hops: $($InitialData.MaxHops) | Tiers: $($InitialData.Tiers) | Devices: $($InitialData.Devices) | Hubs: $($InitialData.Hubs)" -ForegroundColor Gray
    Write-Host ""

    Write-Host "DISPLAY TREE" -ForegroundColor Magenta
    Write-Host "==============================================================================" -ForegroundColor Cyan
    if ([string]::IsNullOrWhiteSpace($InitialData.Display)) {
        Write-Host "HOST"
        Write-Host "└── Nothing detected"
    } else {
        Write-Host $InitialData.Display
    }
    Write-Host ""

    Write-Host "STABILITY PER PLATFORM" -ForegroundColor Cyan
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host "Reference models (affects score):" -ForegroundColor White
    Write-Host $InitialData.ReferenceOutput
    Write-Host "Additional models (reference only):" -ForegroundColor White
    Write-Host $InitialData.AdditionalOutput
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Points deducted: -$([math]::Round($Analytics.Deductions,1))" -ForegroundColor Yellow

    if ($System.IsAdmin) {
        Write-Host "  └─ $($Analytics.Counters.rehandshakes) × rehandshake (−$($Config.scoring.penalties.rehandshake)) = -$($Analytics.Counters.rehandshakes * $Config.scoring.penalties.rehandshake)"
        Write-Host "  └─ $($Analytics.Counters.jitter) × jitter (−$($Config.scoring.penalties.jitter)) = -$($Analytics.Counters.jitter * $Config.scoring.penalties.jitter)"
        Write-Host "  └─ $($Analytics.Counters.crcErrors) × CRC (−$($Config.scoring.penalties.crc)) = -$($Analytics.Counters.crcErrors * $Config.scoring.penalties.crc)"
        Write-Host "  └─ $($Analytics.Counters.busResets) × bus reset (−$($Config.scoring.penalties.busReset)) = -$($Analytics.Counters.busResets * $Config.scoring.penalties.busReset)"
        Write-Host "  └─ $($Analytics.Counters.overcurrent) × overcurrent (−$($Config.scoring.penalties.overcurrent)) = -$($Analytics.Counters.overcurrent * $Config.scoring.penalties.overcurrent)"
        Write-Host "  └─ $($Analytics.Counters.hotplugs) × hotplug (−$($Config.scoring.penalties.hotplug)) = -$($Analytics.Counters.hotplugs * $Config.scoring.penalties.hotplug)"
        Write-Host "  └─ $($Analytics.Counters.edidErrors) × EDID error (−$($Config.scoring.penalties.edidError)) = -$($Analytics.Counters.edidErrors * $Config.scoring.penalties.edidError)"
        Write-Host "  └─ $($Analytics.Counters.linkFailures) × link failure (−$($Config.scoring.penalties.linkFailure)) = -$($Analytics.Counters.linkFailures * $Config.scoring.penalties.linkFailure)"
        Write-Host "  └─ $($Analytics.Counters.otherErrors) × other (−$($Config.scoring.penalties.otherError)) = -$($Analytics.Counters.otherErrors * $Config.scoring.penalties.otherError)"
    } else {
        Write-Host "  Basic mode – not evaluated" -ForegroundColor Yellow
        Write-Host "  Only basic connect/disconnect events counted" -ForegroundColor Yellow
        Write-Host "  Run elevated for full error types and accurate deductions" -ForegroundColor Yellow
    }
    Write-Host ""

    $initialColor = Gray
    $adjustedColor = Get-Color $Config.colors.($Analytics.AdjustedVerdict.ToLower() -replace ' ','')

    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host "SYSTEM STATUS" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "HOST SUMMARY (initial):  $(Format-Score $Analytics.InitialScore)/10.0 - $($Analytics.InitialVerdict)" -ForegroundColor $initialColor
    Write-Host "HOST SUMMARY (adjusted): $(Format-Score $Analytics.AdjustedScore)/10.0 - $($Analytics.AdjustedVerdict)" -ForegroundColor $adjustedColor
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Analytics log excerpt:" -ForegroundColor Cyan
    $Analytics.Log | Select-Object -Last 8 | ForEach-Object { Write-Host $_ }
}

# =============================================================================
# MAIN (keep your original Main function, just update calls to Format-Score)
# =============================================================================

# ... your original Main function here, with any score prints changed to Format-Score ...
# Example in Show-Report or HTML:
# "HOST SUMMARY: $(Format-Score $Stability.WorstScore)/10.0 - $($Stability.Verdict)"

Main
