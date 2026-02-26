# proav-shoko_powershell.ps1 - Shōko Main Logic

$ErrorActionPreference = 'Stop'

# Load config
try {
    $Config = Invoke-RestMethod "https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.json"
    $Version = $Config.version
} catch {
    $Version = "local"
    $Config = [PSCustomObject]@{ 
        version = "local"
        scoring = [PSCustomObject]@{ 
            minScore = 1
            maxScore = 10
            penalties = [PSCustomObject]@{
                rehandshake = 0.5
                jitter = 1.5
                crc = 1.0
                busReset = 1.0
                overcurrent = 2.0
                hotplug = 0.5
                edidError = 1.0
                linkFailure = 1.0
                otherError = 0.25
            }
            thresholds = [PSCustomObject]@{ 
                stable = 7
                potentiallyUnstable = 4
                notStable = 1
            }
        }
        colors = [PSCustomObject]@{ 
            cyan = "Cyan"
            magenta = "Magenta"
            yellow = "Yellow"
            green = "Green"
            gray = "Gray"
            white = "White"
            red = "Red"
        }
        referenceModels = @(
            [PSCustomObject]@{ id = "windows"; name = "Windows x86"; rec = 5; max = 7 }
            [PSCustomObject]@{ id = "windowsArm"; name = "Windows ARM"; rec = 3; max = 5 }
            [PSCustomObject]@{ id = "macAppleSilicon"; name = "Mac Apple Silicon"; rec = 3; max = 5 }
        )
        additionalModels = @(
            [PSCustomObject]@{ id = "ipad"; name = "iPad USB-C (M-series)"; rec = 2; max = 4 }
            [PSCustomObject]@{ id = "iphone"; name = "iPhone USB-C"; rec = 2; max = 4 }
        )
        analytics = [PSCustomObject]@{ 
            updateInterval = 2
            jitterThreshold = 2
            jitterWindow = 5
        }
    }
}

function Get-Color { 
    param($n) 
    $colorMap = @{
        "cyan" = "Cyan"
        "magenta" = "Magenta" 
        "yellow" = "Yellow"
        "green" = "Green"
        "gray" = "Gray"
        "white" = "White"
        "red" = "Red"
    }
    if ($colorMap.ContainsKey($n)) { return $colorMap[$n] } else { return "White" }
}

# System info
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$mode = if ($isAdmin) { "Elevated / admin" } else { "Basic mode" }

$osInfo = Get-WmiObject Win32_OperatingSystem -ErrorAction SilentlyContinue
if (-not $osInfo) {
    $osInfo = [PSCustomObject]@{ Caption = "Windows"; Version = "10.0"; BuildNumber = "0000" }
}
$winVersion = "$($osInfo.Caption) $($osInfo.Version) (Build $($osInfo.BuildNumber))"
$psVersion = $PSVersionTable.PSVersion.ToString()
$arch = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }

# Detect current platform
$currentPlatform = "Windows x86"
if ([Environment]::Is64BitOperatingSystem -and [Environment]::ProcessorArchitecture -eq "Arm64") {
    $currentPlatform = "Windows ARM"
}

Write-Host "`nCollecting system data..." -ForegroundColor Gray

# =============================================================================
# USB TREE - SIMPLIFIED VERSION
# =============================================================================
$allDevices = Get-PnpDevice -Class USB -ErrorAction SilentlyContinue | Where-Object {$_.Status -eq 'OK'}
if (-not $allDevices) { $allDevices = @() }

$devices = @()
$hubs = @()
$treeOutput = "HOST`n"
$maxHops = 0

if ($allDevices.Count -eq 0) {
    $treeOutput += "├── USB Root Hub (Host Controller) [HUB] ← 1 hops`n"
    $treeOutput += "│   └── No USB devices connected`n"
    $maxHops = 1
} else {
    $deviceMap = @{}
    
    foreach ($d in $allDevices) {
        $isHub = ($d.FriendlyName -like "*hub*") -or ($d.Name -like "*hub*") -or ($d.Class -eq "USBHub")
        if ($isHub) {
            $hubs += $d
        } else {
            $devices += $d
        }
        
        $depth = ($d.InstanceId.ToCharArray() | Where-Object {$_ -eq '\'} | Measure-Object).Count
        if ($depth -gt $maxHops) { $maxHops = $depth }
        
        $lastSlash = $d.InstanceId.LastIndexOf('\')
        $parentId = if ($lastSlash -gt 0) { $d.InstanceId.Substring(0, $lastSlash) } else { "" }
        
        $deviceMap[$d.InstanceId] = @{
            Name = if ($d.FriendlyName) { $d.FriendlyName } else { $d.Name }
            Depth = $depth
            IsHub = $isHub
            Parent = $parentId
            Children = @()
        }
    }
    
    foreach ($id in $deviceMap.Keys) {
        $parent = $deviceMap[$id].Parent
        if ($parent -and $deviceMap.ContainsKey($parent)) {
            $deviceMap[$parent].Children += $id
        }
    }
    
    $roots = @()
    foreach ($id in $deviceMap.Keys) {
        $parent = $deviceMap[$id].Parent
        if (-not $parent -or -not $deviceMap.ContainsKey($parent)) {
            $roots += $id
        }
    }
    
    if ($roots.Count -eq 0) {
        $roots = $deviceMap.Keys | Where-Object { $deviceMap[$_].Depth -eq 1 }
    }
    
    $roots = $roots | Sort-Object { $deviceMap[$_].Name }
    
    function Print-Device {
        param($id, $level, $isLast)
        
        $node = $deviceMap[$id]
        if (-not $node) { return }
        
        $prefix = if ($level -eq 0) { "" } else { "│   " * ($level - 1) }
        if ($level -gt 0) {
            $prefix += if ($isLast) { "└── " } else { "├── " }
        } else {
            $prefix = "├── "
        }
        
        $hubTag = if ($node.IsHub) { " [HUB]" } else { "" }
        $script:treeOutput += "$prefix$($node.Name)$hubTag ← $($node.Depth) hops`n"
        
        $children = $node.Children | Sort-Object { $deviceMap[$_].Name }
        for ($i = 0; $i -lt $children.Count; $i++) {
            $childIsLast = ($i -eq $children.Count - 1)
            Print-Device $children[$i] ($level + 1) $childIsLast
        }
    }
    
    for ($i = 0; $i -lt $roots.Count; $i++) {
        $isLastRoot = ($i -eq $roots.Count - 1)
        Print-Device $roots[$i] 0 $isLastRoot
    }
}

$numTiers = $maxHops + 1
$baseScore = [Math]::Max($Config.scoring.minScore, (9 - $maxHops))
$baseScore = [Math]::Min($baseScore, $Config.scoring.maxScore)

# =============================================================================
# STABILITY PER PLATFORM - REFERENCE VS ADDITIONAL
# =============================================================================
$referenceOutput = ""
$additionalOutput = ""
$referenceScores = @()
$worstReferenceScore = 10

# Process reference models (affect score)
foreach ($model in $Config.referenceModels) {
    $rec = $model.rec
    $max = $model.max
    $name = $model.name
    
    $status = if ($numTiers -le $rec) { "STABLE" } 
              elseif ($numTiers -le $max) { "POTENTIALLY UNSTABLE" } 
              else { "NOT STABLE" }
    
    $modelScore = 9 - $maxHops
    if ($modelScore -lt $Config.scoring.minScore) { $modelScore = $Config.scoring.minScore }
    if ($modelScore -gt $Config.scoring.maxScore) { $modelScore = $Config.scoring.maxScore }
    
    $line = "{0,-4} {1,-40} {2}" -f "$numTiers/$max", $name, $status
    $referenceOutput += "$line`n"
    
    $referenceScores += $modelScore
    if ($modelScore -lt $worstReferenceScore) { $worstReferenceScore = $modelScore }
}

# Process additional models (reference only)
foreach ($model in $Config.additionalModels) {
    $rec = $model.rec
    $max = $model.max
    $name = $model.name
    
    $status = if ($numTiers -le $rec) { "STABLE" } 
              elseif ($numTiers -le $max) { "POTENTIALLY UNSTABLE" } 
              else { "NOT STABLE" }
    
    $line = "{0,-4} {1,-40} {2}" -f "$numTiers/$max", $name, $status
    $additionalOutput += "$line`n"
}

if ($referenceScores.Count -eq 0) {
    $worstReferenceScore = $baseScore
}

# =============================================================================
# DISPLAY TREE
# =============================================================================
$displayOutput = "HOST`n"
$monitors = Get-CimInstance -Namespace root\wmi -Class WmiMonitorID -ErrorAction SilentlyContinue

if ($monitors -and $monitors.Count -gt 0) {
    for ($i = 0; $i -lt $monitors.Count; $i++) {
        $m = $monitors[$i]
        
        $name = "Display $($i+1)"
        if ($m.UserFriendlyName -and $m.UserFriendlyName -ne 0) { 
            $name = ($m.UserFriendlyName | ForEach-Object { [char]$_ }) -join '' 
        }
        
        $displayOutput += "└── $name`n"
        $displayOutput += " ├─ Connection : "
        
        $conn = Get-CimInstance -Namespace root\wmi -Class WmiMonitorConnectionParams -Filter "InstanceName = '$($m.InstanceName)'" -ErrorAction SilentlyContinue
        if ($conn -and $conn.VideoOutputTechnology) {
            if ($conn.VideoOutputTechnology -eq 5) { $displayOutput += "HDMI`n" }
            elseif ($conn.VideoOutputTechnology -eq 10) { $displayOutput += "DisplayPort (External)`n" }
            elseif ($conn.VideoOutputTechnology -eq 11) { $displayOutput += "DisplayPort (Embedded / Alt Mode)`n" }
            else { $displayOutput += "Unknown ($($conn.VideoOutputTechnology))`n" }
        } else {
            $displayOutput += "Basic mode`n"
        }
        
        $displayOutput += " ├─ Path       : "
        if ($m.InstanceName -match "DISPLAYPORT") { $displayOutput += "DP / DP Alt Mode`n" }
        elseif ($m.InstanceName -match "USB") { $displayOutput += "USB-C Dock / Alt Mode`n" }
        elseif ($m.InstanceName -match "TBT|THUNDER") { $displayOutput += "Thunderbolt`n" }
        else { $displayOutput += "Direct / Unknown`n" }
        
        if ($isAdmin) {
            $serial = "Basic mode"
            if ($m.SerialNumberID -and $m.SerialNumberID -ne 0) { 
                $serial = ($m.SerialNumberID | ForEach-Object { [char]$_ }) -join '' 
            }
            $displayOutput += " ├─ Size       : Basic mode`n"
            $displayOutput += " ├─ Serial     : $serial`n"
            $displayOutput += " └─ Analytics  : Elevated`n"
        } else {
            $displayOutput += " ├─ Size       : Basic mode`n"
            $displayOutput += " ├─ Serial     : Basic mode`n"
            $displayOutput += " └─ Analytics  : Basic mode`n"
        }
        $displayOutput += "`n"
    }
} else {
    $displayOutput += "└── No displays detected`n"
}

# =============================================================================
# INITIAL DISPLAY
# =============================================================================
Clear-Host
Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host "Shōko - USB + Display Diagnostic Tool v$Version" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host $mode -ForegroundColor (Get-Color $Config.colors.yellow)
Write-Host "Host: $winVersion | PowerShell $psVersion" -ForegroundColor (Get-Color $Config.colors.gray)
Write-Host "Arch: $arch | Current: $currentPlatform" -ForegroundColor (Get-Color $Config.colors.gray)
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
Write-Host "Reference models (affects score):" -ForegroundColor (Get-Color $Config.colors.white)
Write-Host $referenceOutput
Write-Host "Additional models (reference only):" -ForegroundColor (Get-Color $Config.colors.white)
Write-Host $additionalOutput
Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
Write-Host ""

$verdict = if ($worstReferenceScore -ge $Config.scoring.thresholds.stable) { "STABLE" } 
           elseif ($worstReferenceScore -ge $Config.scoring.thresholds.potentiallyUnstable) { "POTENTIALLY UNSTABLE" } 
           else { "NOT STABLE" }
$verdictColor = Get-Color $Config.colors.($verdict.ToLower().Replace(' ', ''))

Write-Host "HOST SUMMARY: $("{0:D2}" -f $worstReferenceScore)/10 - $verdict" -ForegroundColor $verdictColor
Write-Host ""

# =============================================================================
# QUESTION 2 - HTML REPORT
# =============================================================================
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
Arch: $arch | Current: $currentPlatform

USB TREE
==============================================================================
$treeOutput
Max hops: $maxHops | Tiers: $numTiers | Devices: $($devices.Count) | Hubs: $($hubs.Count)

DISPLAY TREE
==============================================================================
$displayOutput
STABILITY PER PLATFORM
==============================================================================
Reference models (affects score):
$referenceOutput
Additional models (reference only):
$additionalOutput
==============================================================================

HOST SUMMARY: $("{0:D2}" -f $worstReferenceScore)/10 - $verdict
</pre></body></html>
"@
    $htmlContent | Out-File $outHtml -Encoding UTF8
    Start-Process $outHtml
}

# =============================================================================
# QUESTION 3 - ANALYTICS
# =============================================================================
$analyticsChoice = Read-Host "Run deep analytics session? (y/n)"
if ($analyticsChoice -match '^[Yy]') {
    # STORE INITIAL DATA
    $initialTree = $treeOutput
    $initialDisplay = $displayOutput
    $initialReferenceOutput = $referenceOutput
    $initialAdditionalOutput = $additionalOutput
    $initialScore = $worstReferenceScore
    $initialVerdict = $verdict
    $initialDevices = $devices.Count
    $initialHubs = $hubs.Count
    $initialMaxHops = $maxHops
    $initialTiers = $numTiers
    
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
    
    while (-not $Host.UI.RawUI.KeyAvailable) {
        $elapsed = (Get-Date) - $startTime
        $duration = "{0:hh\:mm\:ss\.fff}" -f $elapsed
        
        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, 6
        
        $width = $Host.UI.RawUI.WindowSize.Width
        $height = $Host.UI.RawUI.WindowSize.Height
        $currentLine = $Host.UI.RawUI.CursorPosition.Y
        
        for ($y = $currentLine; $y -lt $height; $y++) {
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $y
            Write-Host (" " * $width) -NoNewline
        }
        
        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, 6
        
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
            Write-Host "USB JITTER: $($counters.jitter)"
            Write-Host "USB ERRORS: $($counters.crcErrors + $counters.busResets + $counters.overcurrent + $counters.otherErrors)"
            Write-Host "DISPLAY HOTPLUGS: $($counters.hotplugs)"
            Write-Host "DISPLAY ERRORS: $($counters.edidErrors + $counters.linkFailures)"
        }
        
        Write-Host ""
        Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
        Write-Host "Analytics Log:" -ForegroundColor (Get-Color $Config.colors.cyan)
        
        $startIndex = [Math]::Max(0, $analyticsLog.Count - 20)
        for ($i = $startIndex; $i -lt $analyticsLog.Count; $i++) {
            Write-Host $analyticsLog[$i]
        }
        

        
        Start-Sleep -Milliseconds 500
    }
    
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
    
    Clear-Host
    
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Shōko - USB + Display Diagnostic Tool v$Version" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $mode -ForegroundColor (Get-Color $Config.colors.yellow)
    Write-Host "Host: $winVersion | PowerShell $psVersion" -ForegroundColor (Get-Color $Config.colors.gray)
    Write-Host "Arch: $arch | Current: $currentPlatform" -ForegroundColor (Get-Color $Config.colors.gray)
    Write-Host ""
    
    Write-Host "USB TREE" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $initialTree
    Write-Host "Max hops: $initialMaxHops | Tiers: $initialTiers | Devices: $initialDevices | Hubs: $initialHubs" -ForegroundColor (Get-Color $Config.colors.gray)
    Write-Host ""
    
    Write-Host "DISPLAY TREE" -ForegroundColor (Get-Color $Config.colors.magenta)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host $initialDisplay
    
    Write-Host "STABILITY PER PLATFORM" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Reference models (affects score):" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $initialReferenceOutput
    Write-Host "Additional models (reference only):" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $initialAdditionalOutput
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host ""
    
    Write-Host "HOST SUMMARY: $("{0:D2}" -f $initialScore)/10 - $initialVerdict" -ForegroundColor $verdictColor
    Write-Host "HOST SUMMARY: $("{0:D2}" -f $adjustedScore)/10 - $adjustedVerdict (adjusted)" -ForegroundColor $adjustedColor
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
    
    $finalHtmlChoice = Read-Host "`nOpen HTML report with full data? (y/n)"
    if ($finalHtmlChoice -match '^[Yy]') {
        $dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $outHtml = "$env:TEMP\shoko-report-$dateStamp.html"
        
        if ($isAdmin) {
            $analyticsSummary = @"
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
            $analyticsSummary = @"
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
Arch: $arch | Current: $currentPlatform

USB TREE
==============================================================================
$initialTree
Max hops: $initialMaxHops | Tiers: $initialTiers | Devices: $initialDevices | Hubs: $initialHubs

DISPLAY TREE
==============================================================================
$initialDisplay
STABILITY PER PLATFORM
==============================================================================
Reference models (affects score):
$initialReferenceOutput
Additional models (reference only):
$initialAdditionalOutput
==============================================================================

HOST SUMMARY: $("{0:D2}" -f $initialScore)/10 - $initialVerdict
HOST SUMMARY: $("{0:D2}" -f $adjustedScore)/10 - $adjustedVerdict (adjusted)

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

