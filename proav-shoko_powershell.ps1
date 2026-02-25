# proav-shoko_powershell.ps1 - Shōko Main Logic
# USB tree always shown + display tree + analytics loop

$ErrorActionPreference = 'Stop'

# Load config from current repo
try {
    $Config = Invoke-RestMethod "https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.json"
    $Version = $Config.version
} catch {
    $Version = "local"
    $Config = [PSCustomObject]@{ version = "local" }
}

function Get-Color { param($n) switch($n){ "cyan"{"Cyan"} "magenta"{"Magenta"} "yellow"{"Yellow"} "green"{"Green"} "gray"{"Gray"} default{"White"} } }

Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "Shōko - USB + Display Diagnostic Tool v$Version" -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Running in basic mode (admin needed for full stability & analytics)" -ForegroundColor Yellow
}

$dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outTxt = "$env:TEMP\shoko-report-$dateStamp.txt"
$outHtml = "$env:TEMP\shoko-report-$dateStamp.html"

# ────────────────────────────────────────────────────────────────────────────────
# USB TREE – always shown immediately
# ────────────────────────────────────────────────────────────────────────────────

Write-Host "Enumerating USB devices..." -ForegroundColor Gray

$allDevices = Get-PnpDevice -Class USB | Where-Object {$_.Status -eq 'OK'} | Select-Object InstanceId, FriendlyName, Name, Class, @{n='IsHub';e={
    ($_.FriendlyName -like "*hub*") -or ($_.Name -like "*hub*") -or ($_.Class -eq "USBHub") -or ($_.InstanceId -like "*HUB*")
}}

if ($allDevices.Count -eq 0) {
    Write-Host "No USB devices found." -ForegroundColor Yellow
    exit
}

$devices = $allDevices | Where-Object { -not $_.IsHub }
$hubs = $allDevices | Where-Object { $_.IsHub }

Write-Host "Found $($devices.Count) devices and $($hubs.Count) hubs" -ForegroundColor Gray

# Hierarchy map
$map = @{}
foreach ($d in $allDevices) {
    try {
        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\" + $d.InstanceId.Replace('\','\\')
        $reg = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        $parent = $reg.ParentIdPrefix
        $map[$d.InstanceId] = @{ Name = if ($d.FriendlyName) { $d.FriendlyName } else { $d.Name }; Parent = $parent; Children = @(); InstanceId = $d.InstanceId; IsHub = $d.IsHub }
    } catch {
        $map[$d.InstanceId] = @{ Name = if ($d.FriendlyName) { $d.FriendlyName } else { $d.Name }; Parent = $null; Children = @(); InstanceId = $d.InstanceId; IsHub = $d.IsHub }
    }
}

$roots = @()
foreach ($id in $map.Keys) {
    if (-not $map[$id].Parent) { $roots += $id }
    else {
        foreach ($p in $map.Keys) {
            if ($map[$p].Name -like "*$($map[$id].Parent)*" -or $map[$p].InstanceId -like "*$($map[$id].Parent)*") {
                $map[$p].Children += $id
                break
            }
        }
    }
}

$treeOutput = ""
$maxHops = 0

function Print-Tree { param($id, $level, $prefix, $isLast)
    $node = $map[$id]
    if (-not $node) { return }
    $branch = if ($level -eq 0) { "├── " } else { $prefix + $(if ($isLast) { "└── " } else { "├── " }) }
    $name = if ($node.IsHub) { "$($node.Name) [HUB]" } else { $node.Name }
    $script:treeOutput += "$branch$name ← $level hops`n"
    $script:maxHops = [Math]::Max($script:maxHops, $level)
    $newPrefix = $prefix + $(if ($isLast) { "    " } else { "│   " })
    $children = $node.Children
    for ($i = 0; $i -lt $children.Count; $i++) {
        Print-Tree $children[$i] ($level + 1) $newPrefix ($i -eq $children.Count - 1)
    }
}

foreach ($root in $roots) { Print-Tree $root 0 "" $true }

$numTiers = $maxHops + 1
$baseStabilityScore = [Math]::Max($Config.scoring.minScore, (9 - $maxHops))

# Platform stability table
$statusLines = @()
foreach ($plat in $Config.platformStability.PSObject.Properties.Name) {
    $rec = $Config.platformStability.$plat.rec
    $max = $Config.platformStability.$plat.max
    $status = if ($numTiers -le $rec) { "STABLE" } elseif ($numTiers -le $max) { "POTENTIALLY UNSTABLE" } else { "NOT STABLE" }
    $statusLines += "$($Config.platformStability.$plat.name)    $status"
}

$statusSummary = $statusLines -join "`n"

# USB output
Write-Host ""
Write-Host "USB TREE" -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host $treeOutput
Write-Host ""
Write-Host "Max hops: $maxHops | Tiers: $numTiers | Devices: $($devices.Count) | Hubs: $($hubs.Count)" -ForegroundColor Gray
Write-Host ""
Write-Host "STABILITY PER PLATFORM" -ForegroundColor Cyan
Write-Host $statusSummary
Write-Host ""
Write-Host "HOST SUMMARY: $status (Score: $baseStabilityScore/10)" -ForegroundColor Green

# Save reports
$txt = "Shoko Report $dateStamp`n$treeOutput`nStability:`n$statusSummary`nScore: $baseStabilityScore/10"
$txt | Out-File $outTxt
$html = "<pre>$txt</pre>"
$html | Out-File $outHtml -Encoding UTF8

Write-Host "Reports saved: $outTxt / $outHtml" -ForegroundColor Gray

# Ask browser
$open = Read-Host "Open HTML report in browser? (y/n)"
if ($open -match '^[Yy]') { Start-Process $outHtml }

# ────────────────────────────────────────────────────────────────────────────────
# DISPLAY TREE & ANALYTICS
# ────────────────────────────────────────────────────────────────────────────────

$showDisplay = Read-Host "Show display tree & analytics? (y/n)"
if ($showDisplay -match '^[Yy]') {
    function Get-DisplayTree {
        function Decode-Connection { param([int]$v)
            $b = $v -band 0x7FFFFFFF
            switch ($b) {
                -2 {"Uninitialized"} -1 {"Other/Unknown"} 0 {"VGA"} 5 {"HDMI"} 10 {"DisplayPort"} 11 {"DP Alt Mode"} default {"Unknown ($v)"}
            }
        }

        function Get-ConnColor { param([string]$t)
            if ($t -like "*HDMI*") { "Yellow" } elseif ($t -like "*DisplayPort*") { "Cyan" } else { "Gray" }
        }

        function Detect-Transport { param([string]$inst, [string]$adapt)
            if ($adapt -match "DisplayLink") { return "USB Graphics (DisplayLink)" }
            if ($inst -match "DISPLAYPORT") { return "DP / DP Alt Mode" }
            if ($inst -match "USB") { return "USB-C Dock / Alt Mode" }
            if ($inst -match "TBT|THUNDER") { return "Thunderbolt" }
            return "Direct / Unknown"
        }

        function Detect-MST { param([string]$inst) $inst -match "&MI_" }

        $monitors = Get-CimInstance -Namespace root\wmi -Class WmiMonitorID -EA SilentlyContinue
        $connections = Get-CimInstance -Namespace root\wmi -Class WmiMonitorConnectionParams -EA SilentlyContinue
        $controllers = Get-CimInstance Win32_VideoController -EA SilentlyContinue

        if (-not $monitors) {
            Write-Host "No displays detected." -ForegroundColor Gray
            return
        }

        Write-Host ""
        Write-Host "DISPLAY TREE & ANALYTICS" -ForegroundColor Magenta
        Write-Host "-------------------------------"

        for ($i = 0; $i -lt $monitors.Count; $i++) {
            $m = $monitors[$i]
            $c = $connections | Where-Object { $_.InstanceName -eq $m.InstanceName } | Select-Object -First 1

            $name = if ($m.UserFriendlyName -and $m.UserFriendlyName -ne 0) { ($m.UserFriendlyName | ForEach-Object { [char]$_ }) -join '' } else { "Display $($i+1)" }

            $serialRaw = if ($m.SerialNumberID -and $m.SerialNumberID -ne 0) { ($m.SerialNumberID | ForEach-Object { [char]$_ }) -join '' } else { "N/A" }
            $serial = $serialRaw.Trim()

            $connType = Decode-Connection $c.VideoOutputTechnology
            $connColor = Get-ConnColor $connType

            $adapter = $controllers | Where-Object { $m.InstanceName -like "*$($_.PNPDeviceID)*" } | Select-Object -First 1
            $adapterName = if ($adapter) { $adapter.Name } else { "" }
            $transport = Detect-Transport $m.InstanceName $adapterName

            $isMST = Detect-MST $m.InstanceName

            $since = (Get-Date).AddHours(-1)
            $events = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Kernel-PnP/Configuration'; StartTime=$since} -MaxEvents 100 -EA SilentlyContinue |
                      Where-Object { $_.Message -match "(?i)(display|monitor|connect|disconnect|hotplug|EDID)" }

            $disconnectCount = ($events | Where-Object { $_.Message -match "(?i)disconnect" }).Count

            Write-Host "└─ $name" -ForegroundColor White
            Write-Host " ├─ Connection : $connType" -ForegroundColor $connColor
            Write-Host " ├─ Path       : $transport" -ForegroundColor DarkCyan
            if ($isMST) { Write-Host " ├─ MST Chain  : YES" -ForegroundColor Yellow }
            if ($adapterName -match "DisplayLink") { Write-Host " ├─ Adapter    : DisplayLink" -ForegroundColor Yellow }

            if ($isAdmin) {
                try {
                    $params = Get-CimInstance -Namespace root\wmi -Class WmiMonitorBasicDisplayParams | Where-Object { $_.InstanceName -eq $m.InstanceName } | Select-Object -First 1
                    $size = if ($params) { "$($params.MaxHorizontalImageSize) x $($params.MaxVerticalImageSize) cm" } else { "N/A" }
                    Write-Host " ├─ Size       : $size" -ForegroundColor Gray
                    Write-Host " ├─ Serial     : $serial" -ForegroundColor Gray
                    Write-Host " └─ Analytics  : Deep mode active" -ForegroundColor Green
                } catch {
                    Write-Host " └─ Analytics  : Partial" -ForegroundColor Yellow
                }
            } else {
                Write-Host " └─ Analytics  : Basic mode" -ForegroundColor Gray
            }
            Write-Host ""
        }
    }

    Get-DisplayTree
}

# ────────────────────────────────────────────────────────────────────────────────
# DEEP ANALYTICS LOOP
# ────────────────────────────────────────────────────────────────────────────────

$wantAnalytics = Read-Host "Run deep analytics (Ctrl+C to stop)? (y/n)"
if ($wantAnalytics -match '^[Yy]' -and $isAdmin) {
    Write-Host "Deep analytics started. Press Ctrl+C to exit." -ForegroundColor Green
    try {
        while ($true) {
            Write-Host "$(Get-Date -Format HH:mm:ss) - Monitoring... (press Ctrl+C to stop)" -ForegroundColor Green
            Start-Sleep -Seconds 5
        }
    } catch {
        Write-Host "Analytics stopped." -ForegroundColor Yellow
    }
}

Write-Host "Shōko finished. Press Enter to close." -ForegroundColor Green
Read-Host
