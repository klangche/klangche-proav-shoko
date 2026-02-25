# Shōko - USB + Display Diagnostic Tool - Windows PowerShell Edition
# =============================================================================
# Uses centralized configuration
# Now includes Display Tree & Analytics after USB tree
# =============================================================================

# Load configuration (your current remote config)
try {
    $global:Config = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/klangche/usb-script/main/usb-tree-config.json"
    $scriptVersion = $Config.version
} catch {
    Write-Host "Failed to load configuration. Using defaults." -ForegroundColor Yellow
    $scriptVersion = "1.0.0"
}

# Helper to map config colors to console colors
function Get-Color {
    param($ColorName)
    $hex = $Config.colors.$ColorName
    switch ($hex) {
        "#00ffff" { return "Cyan" }
        "#ff00ff" { return "Magenta" }
        "#ffff00" { return "Yellow" }
        "#00ff00" { return "Green" }
        "#c0c0c0" { return "Gray" }
        default   { return "Gray" }
    }
}

Write-Host "==============================================================================" -ForegroundColor (Get-Color "cyan")
Write-Host "Shōko - USB + DISPLAY DIAGNOSTIC TOOL - WINDOWS EDITION v$scriptVersion" -ForegroundColor (Get-Color "cyan")
Write-Host "==============================================================================" -ForegroundColor (Get-Color "cyan")
Write-Host "$($Config.messages.en.platform): Windows $([System.Environment]::OSVersion.VersionString)" -ForegroundColor (Get-Color "gray")
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# Admin check + elevation
# ─────────────────────────────────────────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    $adminChoice = Read-Host $Config.messages.en.adminPrompt
    if ($adminChoice -match '^[Yy]') {
        Write-Host $Config.messages.en.adminYes -ForegroundColor Yellow
        
        $scriptPath = $MyInvocation.MyCommand.Path
        if (-not $scriptPath) {
            $scriptPath = "$env:TEMP\shoko-temp.ps1"
            $selfContent = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko_powershell.ps1"
            $selfContent | Out-File -FilePath $scriptPath -Encoding UTF8
        }
        
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
        exit
    } else {
        Write-Host "$($Config.messages.en.adminNo) (basic mode for USB & display)" -ForegroundColor Yellow
    }
} else {
    Write-Host $Config.messages.en.adminAlready -ForegroundColor Green
}
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# PART 1: USB TREE ENUMERATION AND REPORTING (your original logic, unchanged)
# ─────────────────────────────────────────────────────────────────────────────
$dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outTxt = "$env:TEMP\shoko-usb-report-$dateStamp.txt"
$outHtml = "$env:TEMP\shoko-usb-report-$dateStamp.html"

Write-Host "$($Config.messages.en.enumerating)..." -ForegroundColor (Get-Color "gray")

$allDevices = Get-PnpDevice -Class USB | Where-Object {$_.Status -eq 'OK'} | Select-Object InstanceId, FriendlyName, Name, Class, @{n='IsHub';e={
    ($_.FriendlyName -like "*hub*") -or ($_.Name -like "*hub*") -or ($_.Class -eq "USBHub") -or ($_.InstanceId -like "*HUB*")
}}

if ($allDevices.Count -eq 0) {
    Write-Host $Config.messages.en.noDevices -ForegroundColor Yellow
    exit
}

$devices = $allDevices | Where-Object { -not $_.IsHub }
$hubs = $allDevices | Where-Object { $_.IsHub }

Write-Host "$($Config.messages.en.found) $($devices.Count) $($Config.messages.en.devices) $($Config.messages.en.and) $($hubs.Count) $($Config.messages.en.hubs)" -ForegroundColor (Get-Color "gray")

# Build map for hierarchy (your original mapping logic)
$map = @{}
foreach ($d in $allDevices) {
    try {
        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\" + $d.InstanceId.Replace('\','\\')
        $reg = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        $parent = $reg.ParentIdPrefix
        $map[$d.InstanceId] = @{ 
            Name = if ($d.FriendlyName) { $d.FriendlyName } else { $d.Name }
            Parent = $parent
            Children = @()
            InstanceId = $d.InstanceId
            IsHub = $d.IsHub
        }
    } catch {
        $map[$d.InstanceId] = @{ 
            Name = if ($d.FriendlyName) { $d.FriendlyName } else { $d.Name }
            Parent = $null
            Children = @()
            InstanceId = $d.InstanceId
            IsHub = $d.IsHub
        }
    }
}

# Build hierarchy + tree output (your original recursive function)
$roots = @()
foreach ($id in $map.Keys) {
    $node = $map[$id]
    if (-not $node.Parent) { $roots += $id }
    else {
        foreach ($parentId in $map.Keys) {
            if ($map[$parentId].Name -like "*$($node.Parent)*" -or $map[$parentId].InstanceId -like "*$($node.Parent)*") {
                $map[$parentId].Children += $id
                break
            }
        }
    }
}

$treeOutput = ""
$maxHops = 0

function Print-Tree {
    param($id, $level, $prefix, $isLast)
    $node = $map[$id]
    if (-not $node) { return }
    
    $branch = if ($level -eq 0) { "├── " } else { $prefix + $(if ($isLast) { "└── " } else { "├── " }) }
    $displayName = if ($node.IsHub) { "$($node.Name) [HUB]" } else { $node.Name }
    $script:treeOutput += "$branch$displayName ← $level hops`n"
    $script:maxHops = [Math]::Max($script:maxHops, $level)
    
    $newPrefix = $prefix + $(if ($isLast) { "    " } else { "│   " })
    $children = $node.Children
    
    for ($i = 0; $i -lt $children.Count; $i++) {
        Print-Tree -id $children[$i] -level ($level + 1) -prefix $newPrefix -isLast ($i -eq $children.Count - 1)
    }
}

foreach ($root in $roots) {
    Print-Tree -id $root -level 0 -prefix "" -isLast $true
}

$numTiers = $maxHops + 1
$totalHubs = $hubs.Count
$totalDevices = $devices.Count
$baseStabilityScore = [Math]::Max($Config.scoring.minScore, (9 - $maxHops))

# Platform stability from config
$statusLines = @()
foreach ($plat in $Config.platformStability.PSObject.Properties.Name) {
    $rec = $Config.platformStability.$plat.rec
    $max = $Config.platformStability.$plat.max
    $status = if ($numTiers -le $rec) { "STABLE" } 
              elseif ($numTiers -le $max) { "POTENTIALLY UNSTABLE" } 
              else { "NOT STABLE" }
    $statusLines += [PSCustomObject]@{ Platform = $Config.platformStability.$plat.name; Status = $status }
}

$maxPlatLen = ($statusLines.Platform | Measure-Object Length -Maximum).Maximum
$statusSummaryTerminal = ""
foreach ($line in $statusLines) {
    $pad = " " * ($maxPlatLen - $line.Platform.Length + 4)
    $statusSummaryTerminal += "$($line.Platform)$pad$($line.Status)`n"
}

$appleSiliconStatus = ($statusLines | Where-Object { $_.Platform -eq "Mac Apple Silicon" }).Status
$hostStatus = $appleSiliconStatus
$hostColor = if ($hostStatus -eq "STABLE") { (Get-Color "green") } 
             elseif ($hostStatus -eq "POTENTIALLY UNSTABLE") { (Get-Color "yellow") } 
             else { (Get-Color "magenta") }

# USB console output
Write-Host ""
Write-Host "==============================================================================" -ForegroundColor (Get-Color "cyan")
Write-Host "USB TREE" -ForegroundColor (Get-Color "cyan")
Write-Host "==============================================================================" -ForegroundColor (Get-Color "cyan")
Write-Host $treeOutput
Write-Host ""
Write-Host "$($Config.messages.en.furthestJumps): $maxHops" -ForegroundColor (Get-Color "gray")
Write-Host "$($Config.messages.en.numberOfTiers): $numTiers" -ForegroundColor (Get-Color "gray")
Write-Host "$($Config.messages.en.totalDevices): $totalDevices" -ForegroundColor (Get-Color "gray")
Write-Host "$($Config.messages.en.totalHubs): $totalHubs" -ForegroundColor (Get-Color "gray")
Write-Host ""
Write-Host "==============================================================================" -ForegroundColor (Get-Color "cyan")
Write-Host "STABILITY PER PLATFORM (based on $maxHops hops)" -ForegroundColor (Get-Color "cyan")
Write-Host "==============================================================================" -ForegroundColor (Get-Color "cyan")
Write-Host $statusSummaryTerminal
Write-Host ""
Write-Host "==============================================================================" -ForegroundColor (Get-Color "cyan")
Write-Host "HOST SUMMARY" -ForegroundColor (Get-Color "cyan")
Write-Host "==============================================================================" -ForegroundColor (Get-Color "cyan")
Write-Host "$($Config.messages.en.hostStatus): " -NoNewline
Write-Host "$hostStatus" -ForegroundColor $hostColor
Write-Host "$($Config.messages.en.stabilityScore): $baseStabilityScore/10" -ForegroundColor (Get-Color "gray")
Write-Host ""

# Save text report (USB part)
$txtReport = @"
Shōko USB REPORT - $dateStamp

$treeOutput

$($Config.messages.en.furthestJumps): $maxHops
$($Config.messages.en.numberOfTiers): $numTiers
$($Config.messages.en.totalDevices): $totalDevices
$($Config.messages.en.totalHubs): $totalHubs

STABILITY SUMMARY
$statusSummaryTerminal

$($Config.messages.en.hostStatus): $hostStatus ($($Config.messages.en.stabilityScore): $baseStabilityScore/10)
"@
$txtReport | Out-File $outTxt
Write-Host "$($Config.messages.en.reportSaved): $outTxt" -ForegroundColor (Get-Color "gray")

# Basic HTML (USB part only for now)
$htmlContent = @"
<html><body style="background:#000;color:#0f0;font-family:Consolas;">
<pre>
Shōko USB Report - $dateStamp

==============================================================================
USB TREE
==============================================================================
$treeOutput

$($Config.messages.en.furthestJumps): $maxHops
$($Config.messages.en.numberOfTiers): $numTiers
$($Config.messages.en.totalDevices): $totalDevices
$($Config.messages.en.totalHubs): $totalHubs

==============================================================================
STABILITY PER PLATFORM
==============================================================================
$statusSummaryTerminal

==============================================================================
HOST SUMMARY
==============================================================================
$($Config.messages.en.hostStatus): $hostStatus
$($Config.messages.en.stabilityScore): $baseStabilityScore/10
</pre>
</body></html>
"@
$htmlContent | Out-File $outHtml -Encoding UTF8
Write-Host "$($Config.messages.en.htmlSaved): $outHtml" -ForegroundColor (Get-Color "gray")

# ─────────────────────────────────────────────────────────────────────────────
# PART 1.5: DISPLAY TREE & ANALYTICS (new)
# ─────────────────────────────────────────────────────────────────────────────

function Get-DisplayTree {
    function Decode-Connection { param([int]$v)
        $b = $v -band 0x7FFFFFFF
        switch ($b) {
            -2 {"Uninitialized"} -1 {"Other/Unknown"} 0 {"VGA (HD15)"} 1 {"S-Video"}
            2 {"Composite"} 3 {"Component"} 4 {"DVI"} 5 {"HDMI"} 6 {"LVDS (Internal)"}
            10 {"DisplayPort (External)"} 11 {"DisplayPort (Embedded / Alt Mode)"}
            15 {"Miracast (Wireless)"} default {"Unknown ($v)"}
        }
    }

    function Get-ConnColor { param([string]$t)
        switch -Wildcard ($t) {
            "*HDMI*"     { "Yellow" }
            "*DisplayPort*" { "Cyan" }
            "*VGA*"      { "Red" }
            "*Internal*" { "DarkGray" }
            "*Miracast*" { "Magenta" }
            "*Unknown*"  { "DarkGray" }
            default      { "Gray" }
        }
    }

    function Detect-Transport { param([string]$inst, [string]$adapt)
        if ($adapt -match "DisplayLink") { return "USB Graphics (DisplayLink)" }
        if ($inst -match "DISPLAYPORT") { "DP / DP Alt Mode" }
        elseif ($inst -match "USB")     { "USB-C Dock / Alt Mode" }
        elseif ($inst -match "TBT|THUNDER") { "Thunderbolt" }
        elseif ($inst -match "HDMI")    { "HDMI" }
        else { "Direct / Unknown" }
    }

    function Detect-MST { param([string]$inst) if ($inst -match "&MI_") { $true } else { $false } }

    function Get-HealthHint { param([string]$ct, [bool]$mst, [string]$tr, [int]$dc)
        if ($dc -gt 1) { return "Magenta", "Potential issue (disconnects detected)" }
        if ($mst -or $tr -match "DisplayLink") { return "Yellow", "Watch (MST chain or software path)" }
        if ($ct -match "DisplayPort|HDMI") { return "Green", "Stable" }
        return "Gray", "N/A"
    }

    $monitors    = Get-CimInstance -Namespace root\wmi -Class WmiMonitorID -EA SilentlyContinue
    $connections = Get-CimInstance -Namespace root\wmi -Class WmiMonitorConnectionParams -EA SilentlyContinue
    $controllers = Get-CimInstance Win32_VideoController -EA SilentlyContinue

    if (!$monitors -or $monitors.Count -eq 0) {
        Write-Host "No displays detected beyond primary/internal." -ForegroundColor (Get-Color "gray")
        return
    }

    Write-Host ""
    Write-Host "==============================================================================" -ForegroundColor (Get-Color "cyan")
    Write-Host "DISPLAY TREE & ANALYTICS" -ForegroundColor (Get-Color "magenta")
    Write-Host "==============================================================================" -ForegroundColor (Get-Color "cyan")

    foreach ($i in 0..($monitors.Count-1)) {
        $m = $monitors[$i]
        $c = $connections | Where-Object { $_.InstanceName -eq $m.InstanceName } | Select-Object -First 1

        $name = if ($m.UserFriendlyName -and $m.UserFriendlyName -ne 0) { ($m.UserFriendlyName | ForEach-Object {[char]$_}) -join '' } else { "Display $($i+1)" }
        $serial = if ($m.SerialNumberID -and $m.SerialNumberID -ne 0) { ($m.SerialNumberID | ForEach-Object {[char]$_}) -join '' .Trim() } else { "N/A" }

        $connType = Decode-Connection $c.VideoOutputTechnology
        $connColor = Get-ConnColor $connType

        $adapter = $controllers | Where-Object { $m.InstanceName -like "*$($_.PNPDeviceID)*" } | Select-Object -First 1
        $transport = Detect-Transport $m.InstanceName ($adapter ? $adapter.Name : "")
        $isMST = Detect-MST $m.InstanceName

        $since = (Get-Date).AddHours(-1)
        $events = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Kernel-PnP/Configuration'; StartTime=$since} -MaxEvents 100 -EA SilentlyContinue |
                  Where-Object { $_.Message -match "(?i)(display|monitor|connect|disconnect|hotplug|EDID)" }

        $disconnectCount = ($events | Where-Object { $_.Message -match "(?i)disconnect" }).Count

        $fgColor, $hintText = Get-HealthHint $connType $isMST $transport $disconnectCount

        Write-Host "└─ $name" -ForegroundColor White
        Write-Host " ├─ Connection : $connType" -ForegroundColor $connColor
        Write-Host " ├─ Path       : $transport" -ForegroundColor DarkCyan
        if ($isMST) { Write-Host " ├─ MST Chain  : YES" -ForegroundColor Yellow }
        if ($adapter -and $adapter.Name -match "DisplayLink") {
            Write-Host " ├─ Adapter    : DisplayLink (software)" -ForegroundColor Yellow
        }

        if ($isAdmin) {
            try {
                $params = Get-CimInstance -Namespace root\wmi -Class WmiMonitorBasicDisplayParams |
                          Where-Object { $_.InstanceName -eq $m.InstanceName } | Select-Object -First 1
                $size = if ($params) { "$($params.MaxHorizontalImageSize)×$($params.MaxVerticalImageSize) cm" } else { "N/A" }

                Write-Host " ├─ Size       : $size" -ForegroundColor (Get-Color "gray")
                Write-Host " ├─ Serial     : $serial" -ForegroundColor (Get-Color "gray")

                if ($events.Count -gt 0) {
                    Write-Host " ├─ Recent Events ($($events.Count) in last hour):" -ForegroundColor Green
                    $events | Select-Object -Last 5 | ForEach-Object {
                        Write-Host " │   $($_.TimeCreated): $($_.Message.Trim())" -ForegroundColor (Get-Color "gray")
                    }
                    if ($disconnectCount -gt 0) {
                        Write-Host " └─ Issues     : $disconnectCount disconnects detected" -ForegroundColor Magenta
                    } else {
                        Write-Host " └─ Analytics  : Deep Mode - No major issues" -ForegroundColor Green
                    }
                } else {
                    Write-Host " └─ Analytics  : Deep Mode - No recent events" -ForegroundColor Green
                }
            } catch {
                Write-Host " └─ Analytics  : Deep Mode Partial (query error)" -ForegroundColor Yellow
            }
        } else {
            $eventSum = if ($events.Count -gt 0) { "$($events.Count) recent events (admin → details)" } else { "No recent events captured" }
            Write-Host " └─ Analytics  : Basic Mode - $eventSum" -ForegroundColor (Get-Color "gray")
        }
        Write-Host ""
    }

    Write-Host "Health Hints: Green = Stable | Yellow = Watch | Magenta = Potential Issue" -ForegroundColor (Get-Color "gray")
    Write-Host ""
}

# Prompt for display section
Write-Host ""
$showDisplay = Read-Host "Show display information and analytics? (y/n)"
if ($showDisplay -match '^[Yy]$') {
    Get-DisplayTree
}

# ─────────────────────────────────────────────────────────────────────────────
# PART 2: DEEP USB ANALYTICS (your original continues here – unchanged)
# ─────────────────────────────────────────────────────────────────────────────

Write-Host ""
$wantDeep = Read-Host $Config.messages.en.deepPrompt
if ($wantDeep -notmatch '^[Yy]') {
    Write-Host "$($Config.messages.en.deepPrompt) skipped." -ForegroundColor (Get-Color "gray")
    Write-Host $Config.messages.en.exitPrompt -ForegroundColor (Get-Color "gray")
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# ... (your original deep analytics loop code goes here – counters, monitoring until Ctrl+C, etc.)
# For brevity I stopped here; paste your existing deep analytics part after this comment.

# Example placeholder for end of script
Write-Host "Shōko complete. Reports saved." -ForegroundColor Green
Read-Host "Press Enter to exit"
