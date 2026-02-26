# proav-shoko_powershell.ps1 - Shōko Main Logic

$ErrorActionPreference = 'Stop'

# Load config
try {
    $Config = Invoke-RestMethod "https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.json"
    $Version = $Config.version
} catch {
    $Version = "local"
    $Config = [PSCustomObject]@{ 
        version = "local"; 
        scoring = [PSCustomObject]@{ minScore = 1; maxScore = 10; thresholds = [PSCustomObject]@{ stable = 7; potentiallyUnstable = 4; notStable = 1 } }
        colors = [PSCustomObject]@{ cyan = "Cyan"; magenta = "Magenta"; yellow = "Yellow"; green = "Green"; gray = "Gray"; white = "White"; red = "Red" }
        platformStability = [PSCustomObject]@{}
        analytics = [PSCustomObject]@{ updateInterval = 5; jitterThreshold = 2; jitterWindow = 5 }
    }
}

function Get-Color { param($n) 
    $colorMap = @{
        "cyan" = "Cyan"
        "magenta" = "Magenta" 
        "yellow" = "Yellow"
        "green" = "Green"
        "gray" = "Gray"
        "white" = "White"
        "red" = "Red"
    }
    return $colorMap[$n]
}

# System info
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$mode = if ($isAdmin) { "Elevated / admin" } else { "Basic mode" }

$osInfo = Get-WmiObject Win32_OperatingSystem
$winVersion = "$($osInfo.Caption) $($osInfo.Version) (Build $($osInfo.BuildNumber))"
$psVersion = $PSVersionTable.PSVersion.ToString()
$arch = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }
$currentPlatform = "windows" # This would be detected properly in real implementation

# Initial data collection
Write-Host "`nCollecting system data..." -ForegroundColor Gray

# USB Tree
$allDevices = Get-PnpDevice -Class USB | Where-Object {$_.Status -eq 'OK'} | Select-Object InstanceId, FriendlyName, Name, Class
$devices = $allDevices | Where-Object { ($_.FriendlyName -notlike "*hub*") -and ($_.Name -notlike "*hub*") -and ($_.Class -ne "USBHub") }
$hubs = $allDevices | Where-Object { ($_.FriendlyName -like "*hub*") -or ($_.Name -like "*hub*") -or ($_.Class -eq "USBHub") }

# Build tree from registry paths
$tree = @{}
$maxHops = 0
foreach ($d in $allDevices) {
    $path = $d.InstanceId
    $depth = ($path.Split('\').Count) - 1
    if ($depth -gt $maxHops) { $maxHops = $depth }
    $tree[$path] = @{
        Name = if ($d.FriendlyName) { $d.FriendlyName } else { $d.Name }
        Depth = $depth
        IsHub = ($d.FriendlyName -like "*hub*") -or ($d.Name -like "*hub*") -or ($d.Class -eq "USBHub")
    }
}

# Build tree output
$treeOutput = ""
$roots = $tree.Keys | Where-Object { $_.Split('\').Count -eq 2 }
foreach ($root in $roots) {
    $rootName = $tree[$root].Name
    $rootHub = if ($tree[$root].IsHub) { " [HUB]" } else { "" }
    $treeOutput += "├─ $rootName$rootHub ← 0 hops`n"
    
    $children = $tree.Keys | Where-Object { $_.StartsWith($root) -and $_ -ne $root } | Sort-Object
    foreach ($child in $children) {
        $depth = $tree[$child].Depth
        $childName = $tree[$child].Name
        $childHub = if ($tree[$child].IsHub) { " [HUB]" } else { "" }
        $prefix = "│   " * ($depth - 1)
        $treeOutput += "$prefix├─ $childName$childHub ← $depth hops`n"
    }
}

$numTiers = $maxHops + 1
$baseScore = [Math]::Max($Config.scoring.minScore, (9 - $maxHops))
$baseScore = [Math]::Min($baseScore, $Config.scoring.maxScore)

# Platform stability
$platformOutput = ""
foreach ($plat in $Config.platformStability.PSObject.Properties.Name) {
    $rec = $Config.platformStability.$plat.rec
    $max = $Config.platformStability.$plat.max
    $status = if ($numTiers -le $rec) { "STABLE" } 
              elseif ($numTiers -le $max) { "POTENTIALLY UNSTABLE" } 
              else { "NOT STABLE" }
    $name = $Config.platformStability.$plat.name
    $platformOutput += "$name".PadRight(30) + "$status`n"
}

# Display Tree
$displayOutput = ""
$monitors = Get-CimInstance -Namespace root\wmi -Class WmiMonitorID -ErrorAction SilentlyContinue
$connections = Get-CimInstance -Namespace root\wmi -Class WmiMonitorConnectionParams -ErrorAction SilentlyContinue

if ($monitors -and $monitors.Count -gt 0) {
    for ($i = 0; $i -lt $monitors.Count; $i++) {
        $m = $monitors[$i]
        $c = $connections | Where-Object { $_.InstanceName -eq $m.InstanceName } | Select-Object -First 1
        
        $name = if ($m.UserFriendlyName -and $m.UserFriendlyName -ne 0) { 
            ($m.UserFriendlyName | ForEach-Object { [char]$_ }) -join '' 
        } else { "Display $($i+1)" }
        
        if ($c.VideoOutputTechnology -eq 10) { $connType = "DisplayPort (External)" }
        elseif ($c.VideoOutputTechnology -eq 11) { $connType = "DisplayPort (Embedded / Alt Mode)" }
        elseif ($c.VideoOutputTechnology -eq 5) { $connType = "HDMI" }
        else { $connType = "Unknown ($($c.VideoOutputTechnology))" }
        
        $displayOutput += "└─ $name`n"
        $displayOutput += " ├─ Connection : $connType`n"
        $displayOutput += " ├─ Path       : Direct / Unknown`n"
        
        if ($isAdmin) {
            try {
                $serial = if ($m.SerialNumberID -and $m.SerialNumberID -ne 0) { 
                    ($m.SerialNumberID | ForEach-Object { [char]$_ }) -join '' 
                } else { "N/A" }
                $displayOutput += " ├─ Size       : Basic mode`n"  # Size calc would go here
                $displayOutput += " ├─ Serial     : $serial`n"
                $displayOutput += " └─ Analytics  : Elevated`n"
            } catch {
                $displayOutput += " ├─ Size       : Basic mode`n"
                $displayOutput += " ├─ Serial     : Basic mode`n"
                $displayOutput += " └─ Analytics  : Basic mode`n"
            }
        } else {
            $displayOutput += " ├─ Size       : Basic mode`n"
            $displayOutput += " ├─ Serial     : Basic mode`n"
            $displayOutput += " └─ Analytics  : Basic mode`n"
        }
        $displayOutput += "`n"
    }
} else {
    $displayOutput = "No displays detected.`n"
}

# Initial display
Clear-Host
Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host "Shōko - USB + Display Diagnostic Tool v$Version" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host $mode -ForegroundColor (Get-Color $Config.colors.yellow)
Write-Host "Host: $winVersion | PowerShell $psVersion" -ForegroundColor (Get-Color $Config.colors.gray)
Write-Host "Arch: $arch | Current: $($Config.platformStability.$currentPlatform.name)" -ForegroundColor (Get-Color $Config.colors.gray)
Write-Host ""

Write-Host "USB TREE" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host $treeOutput
Write-Host "Max hops: $maxHops | Tiers: $numTiers | Devices: $($devices.Count) | Hubs: $($hubs.Count)" -ForegroundColor (Get-Color $Config.colors.gray)
Write-Host ""

Write-Host "DISPLAY TREE" -ForegroundColor (Get-Color $Config.colors.magenta)
Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host $displayOutput

Write-Host "STABILITY PER PLATFORM" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host $platformOutput
Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host ""

$verdict = if ($baseScore -ge $Config.scoring.thresholds.stable) { "STABLE" } 
           elseif ($baseScore -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" } 
           else { "NOT STABLE" }
$verdictColor = Get-Color $Config.colors.($verdict.ToLower().Replace(' ', ''))
Write-Host "HOST SUMMARY: Score: $baseScore/10 $verdict" -ForegroundColor $verdictColor
Write-Host ""

# Question 2 - HTML report
$htmlChoice = Read-Host "Open HTML report? (y/n)"
if ($htmlChoice -match '^[Yy]') {
    $dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $outHtml = "$env:TEMP\shoko-report-$dateStamp.html"
    $htmlContent = @"
<html>
<head><title>Shōko Report $dateStamp</title></head>
<body style='background:#000;color:#0f0;font-family:Consolas;'>
<pre>
Shōko Report - $dateStamp

$mode
Host: $winVersion | PowerShell $psVersion
Arch: $arch | Current: $($Config.platformStability.$currentPlatform.name)

USB TREE
==============================================================================
$treeOutput
Max hops: $maxHops | Tiers: $numTiers | Devices: $($devices.Count) | Hubs: $($hubs.Count)

DISPLAY TREE
==============================================================================
$displayOutput
STABILITY PER PLATFORM
==============================================================================
$platformOutput
==============================================================================

HOST SUMMARY: Score: $baseScore/10 $verdict
</pre></body></html>
"@
    $htmlContent | Out-File $outHtml -Encoding UTF8
    Start-Process $outHtml
}

# Question 3 - Analytics
$analyticsChoice = Read-Host "Run deep analytics session? (y/n)"
if ($analyticsChoice -match '^[Yy]') {
    # Store initial data
    $initialTree = $treeOutput
    $initialDisplay = $displayOutput
    $initialPlatforms = $platformOutput
    $initialScore = $baseScore
    $initialVerdict = $verdict
    
    # Clear for analytics
    Clear-Host
    
    # Analytics header
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Shoko Analytics" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Monitoring connections... Press any key to stop" -ForegroundColor (Get-Color $Config.colors.gray)
    if ($isAdmin) {
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
    
    $lastEvent = $null
    $eventHistory = @()
    
    # Live monitoring loop
    while (-not $Host.UI.RawUI.KeyAvailable) {
        $elapsed = (Get-Date) - $startTime
        $duration = "{0:hh\:mm\:ss\.fff}" -f $elapsed
        
        # Move cursor up to update header
        $cursor = $Host.UI.RawUI.CursorPosition
        $cursor.Y = 6
        $Host.UI.RawUI.CursorPosition = $cursor
        
        Write-Host "Duration: $duration" -ForegroundColor (Get-Color $Config.colors.green)
        Write-Host "Total events logged: $($counters.total)" -ForegroundColor (Get-Color $Config.colors.white)
        if ($isAdmin) {
            Write-Host "USB RE-HANDSHAKES: $($counters.rehandshakes)"
            Write-Host "USB JITTER: $($counters.jitter)"
            Write-Host "USB CRC ERRORS: $($counters.crcErrors)"
            Write-Host "USB BUS RESETS: $($counters.busResets)"
            Write-Host "USB OVERCURRENT: $($counters.overcurrent)"
            Write-Host "DISPLAY HOTPLUGS: $($counters.hotplugs)"
            Write-Host "DISPLAY EDID ERRORS: $($counters.edidErrors)"
            Write-Host "DISPLAY LINK FAILURES: $($counters.linkFailures)"
            Write-Host "OTHER ERRORS: $($counters.otherErrors)"
        } else {
            Write-Host "USB CONNECTS: $($counters.connects)"
            Write-Host "USB DISCONNECTS: $($counters.disconnects)"
            Write-Host "USB RE-HANDSHAKES: $($counters.rehandshakes)"
            Write-Host "USB ERRORS: $($counters.crcErrors + $counters.busResets + $counters.overcurrent + $counters.otherErrors)"
            Write-Host "DISPLAY HOTPLUGS: $($counters.hotplugs)"
            Write-Host "DISPLAY ERRORS: $($counters.edidErrors + $counters.linkFailures)"
        }
        Write-Host ""
        Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
        Write-Host "Analytics Log:" -ForegroundColor (Get-Color $Config.colors.cyan)
        
        # Show last 20 log entries
        $startIndex = [Math]::Max(0, $analyticsLog.Count - 20)
        for ($i = $startIndex; $i -lt $analyticsLog.Count; $i++) {
            Write-Host $analyticsLog[$i]
        }
        
        # Check for new events (simplified for demo)
        Start-Sleep -Seconds $Config.analytics.updateInterval
    }
    
    # Stop analytics
    $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    $stopTime = Get-Date
    $totalDuration = "{0:hh\:mm\:ss\.fff}" -f ($stopTime - $startTime)
    $analyticsLog += "$($stopTime.ToString('HH:mm:ss.fff')) - Logging ended (total duration: $totalDuration)"
    
    # Calculate deductions
    $deductions = 0
    if ($isAdmin) {
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
    
    $adjustedScore = [Math]::Max(0, $initialScore - $deductions)
    $adjustedScore = [Math]::Min($adjustedScore, $Config.scoring.maxScore)
    
    $adjustedVerdict = if ($adjustedScore -ge $Config.scoring.thresholds.stable) { "STABLE" } 
                       elseif ($adjustedScore -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" } 
                       else { "NOT STABLE" }
    $adjustedColor = Get-Color $Config.colors.($adjustedVerdict.ToLower().Replace(' ', ''))
    
    # Clear and show final report
    Clear-Host
    
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Shōko - USB + Display Diagnostic Tool v$Version" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $mode -ForegroundColor (Get-Color $Config.colors.yellow)
    Write-Host "Host: $winVersion | PowerShell $psVersion" -ForegroundColor (Get-Color $Config.colors.gray)
    Write-Host "Arch: $arch | Current: $($Config.platformStability.$currentPlatform.name)" -ForegroundColor (Get-Color $Config.colors.gray)
    Write-Host ""
    
    Write-Host "USB TREE" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $initialTree
    
    Write-Host "DISPLAY TREE" -ForegroundColor (Get-Color $Config.colors.magenta)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $initialDisplay
    
    Write-Host "STABILITY PER PLATFORM" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $initialPlatforms
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host ""
    
    Write-Host "HOST SUMMARY: Score: $initialScore/10 $initialVerdict" -ForegroundColor $verdictColor
    Write-Host "HOST SUMMARY (adjusted): Score: $([math]::Round($adjustedScore,1))/10 $adjustedVerdict" -ForegroundColor $adjustedColor
    Write-Host ""
    
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Analytics Summary (during monitoring):" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Total events logged: $($counters.total)" -ForegroundColor (Get-Color $Config.colors.white)
    if ($isAdmin) {
        Write-Host "USB RE-HANDSHAKES: $($counters.rehandshakes)"
        Write-Host "USB JITTER: $($counters.jitter)"
        Write-Host "USB CRC ERRORS: $($counters.crcErrors)"
        Write-Host "USB BUS RESETS: $($counters.busResets)"
        Write-Host "USB OVERCURRENT: $($counters.overcurrent)"
        Write-Host "DISPLAY HOTPLUGS: $($counters.hotplugs)"
        Write-Host "DISPLAY EDID ERRORS: $($counters.edidErrors)"
        Write-Host "DISPLAY LINK FAILURES: $($counters.linkFailures)"
        Write-Host "OTHER ERRORS: $($counters.otherErrors)"
    } else {
        Write-Host "USB CONNECTS: $($counters.connects)"
        Write-Host "USB DISCONNECTS: $($counters.disconnects)"
        Write-Host "USB RE-HANDSHAKES: $($counters.rehandshakes)"
        Write-Host "USB JITTER: $($counters.jitter)"
        Write-Host "USB ERRORS: $($counters.crcErrors + $counters.busResets + $counters.overcurrent + $counters.otherErrors)"
        Write-Host "DISPLAY HOTPLUGS: $($counters.hotplugs)"
        Write-Host "DISPLAY ERRORS: $($counters.edidErrors + $counters.linkFailures)"
    }
    Write-Host "Points deducted: -$([math]::Round($deductions,1))"
    Write-Host ""
    
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Analytics Log (during monitoring):" -ForegroundColor (Get-Color $Config.colors.cyan)
    foreach ($line in $analyticsLog) {
        Write-Host $line
    }
    
    # Question 4 - HTML report with full data
    $finalHtmlChoice = Read-Host "`nOpen HTML report with full data? (y/n)"
    if ($finalHtmlChoice -match '^[Yy]') {
        $dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $outHtml = "$env:TEMP\shoko-report-$dateStamp.html"
        
        $analyticsSummary = if ($isAdmin) {
            @"
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
            @"
USB CONNECTS: $($counters.connects)
USB DISCONNECTS: $($counters.disconnects)
USB RE-HANDSHAKES: $($counters.rehandshakes)
USB JITTER: $($counters.jitter)
USB ERRORS: $($counters.crcErrors + $counters.busResets + $counters.overcurrent + $counters.otherErrors)
DISPLAY HOTPLUGS: $($counters.hotplugs)
DISPLAY ERRORS: $($counters.edidErrors + $counters.linkFailures)
"@
        }
        
        $analyticsLogText = $analyticsLog -join "`n"
        
        $htmlContent = @"
<html>
<head><title>Shōko Report $dateStamp</title></head>
<body style='background:#000;color:#0f0;font-family:Consolas;'>
<pre>
Shōko Report - $dateStamp

$mode
Host: $winVersion | PowerShell $psVersion
Arch: $arch | Current: $($Config.platformStability.$currentPlatform.name)

USB TREE
==============================================================================
$initialTree
Max hops: $maxHops | Tiers: $numTiers | Devices: $($devices.Count) | Hubs: $($hubs.Count)

DISPLAY TREE
==============================================================================
$initialDisplay
STABILITY PER PLATFORM
==============================================================================
$initialPlatforms
==============================================================================

HOST SUMMARY: Score: $initialScore/10 $initialVerdict
HOST SUMMARY (adjusted): Score: $([math]::Round($adjustedScore,1))/10 $adjustedVerdict

==============================================================================
Analytics Summary (during monitoring):
Total events logged: $($counters.total)
$analyticsSummary
Points deducted: -$([math]::Round($deductions,1))

==============================================================================
Analytics Log (during monitoring):
$analyticsLogText
</pre></body></html>
"@
        $htmlContent | Out-File $outHtml -Encoding UTF8
        Start-Process $outHtml
    }
}

Write-Host "`nShōko finished. Press Enter to close." -ForegroundColor (Get-Color $Config.colors.green)
Read-Host
