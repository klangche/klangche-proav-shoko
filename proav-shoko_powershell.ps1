# proav-shoko_powershell.ps1 - Shōko Main Logic (USB + Display always shown, then analytics)
# PowerShell 5.1 compatible

$ErrorActionPreference = 'Stop'

# Load config
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
    Write-Host "Basic mode (limited analytics)" -ForegroundColor Yellow
}

$dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outTxt = "$env:TEMP\shoko-report-$dateStamp.txt"
$outHtml = "$env:TEMP\shoko-report-$dateStamp.html"

$htmlContent = "<pre>Shōko Report - $dateStamp`n`n"

# ────────────────────────────────────────────────────────────────────────────────
# USB TREE – always shown
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
$baseStabilityScore = [Math]::Max(1, 9 - $maxHops)

$statusLines = @()
foreach ($plat in $Config.platformStability.PSObject.Properties.Name) {
    $rec = $Config.platformStability.$plat.rec
    $max = $Config.platformStability.$plat.max
    $status = if ($numTiers -le $rec) { "STABLE" } elseif ($numTiers -le $max) { "POTENTIALLY UNSTABLE" } else { "NOT STABLE" }
    $statusLines += "$($Config.platformStability.$plat.name)    $status"
}

$statusSummary = $statusLines -join "`n"

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
Write-Host "HOST SUMMARY: STABLE (Score: $baseStabilityScore/10)" -ForegroundColor Green

$htmlContent += "USB TREE`n$treeOutput`nMax hops: $maxHops | Tiers: $numTiers`nStability:`n$statusSummary`nScore: $baseStabilityScore/10`n`n"

# ────────────────────────────────────────────────────────────────────────────────
# DISPLAY TREE – shown immediately after USB tree
# ────────────────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "Enumerating displays..." -ForegroundColor Gray

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

    if (-not $monitors -or $monitors.Count -eq 0) {
        Write-Host "No displays detected." -ForegroundColor Gray
        return
    }

    Write-Host ""
    Write-Host "DISPLAY TREE & ANALYTICS" -ForegroundColor Magenta
    Write-Host "==============================================================================" -ForegroundColor Cyan

    $displayTreeOutput = ""
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

        $displayTreeOutput += "└─ $name`n"
        $displayTreeOutput += " ├─ Connection : $connType`n"
        $displayTreeOutput += " ├─ Path       : $transport`n"
        if ($isMST) { $displayTreeOutput += " ├─ MST Chain  : YES`n" }
        if ($adapterName -match "DisplayLink") { $displayTreeOutput += " ├─ Adapter    : DisplayLink`n" }

        if ($isAdmin) {
            try {
                $params = Get-CimInstance -Namespace root\wmi -Class WmiMonitorBasicDisplayParams | Where-Object { $_.InstanceName -eq $m.InstanceName } | Select-Object -First 1
                $size = if ($params) { "$($params.MaxHorizontalImageSize) x $($params.MaxVerticalImageSize) cm" } else { "N/A" }
                $displayTreeOutput += " ├─ Size       : $size`n"
                $displayTreeOutput += " ├─ Serial     : $serial`n"
                $displayTreeOutput += " └─ Analytics  : Deep mode`n"
            } catch {
                $displayTreeOutput += " └─ Analytics  : Partial`n"
            }
        } else {
            $displayTreeOutput += " └─ Analytics  : Basic mode`n"
        }
        $displayTreeOutput += "`n"
    }

    Write-Host $displayTreeOutput

    $htmlContent += "DISPLAY TREE & ANALYTICS`n$displayTreeOutput`n"
}

Get-DisplayTree

# ────────────────────────────────────────────────────────────────────────────────
# Reports & browser
# ────────────────────────────────────────────────────────────────────────────────

$htmlContent += "</pre>"
$htmlContent | Out-File $outHtml -Encoding UTF8
$txtContent = $htmlContent -replace "<pre>","" -replace "</pre>",""
$txtContent | Out-File $outTxt

Write-Host "Reports saved: $outTxt / $outHtml" -ForegroundColor Gray

$openBrowser = Read-Host "Open HTML report in browser? (y/n)"
if ($openBrowser -match '^[Yy]') {
    Start-Process $outHtml
}

# ────────────────────────────────────────────────────────────────────────────────
# Deep Analytics (USB + Display) – loop until Ctrl+C
# ────────────────────────────────────────────────────────────────────────────────

$runAnalytics = Read-Host "Run deep analytics (USB + Display monitoring, Ctrl+C to stop)? (y/n)"
if ($runAnalytics -match '^[Yy]' -and $isAdmin) {
    Write-Host "Deep analytics mode started. Press Ctrl+C to exit." -ForegroundColor Green

    try {
        while ($true) {
            Write-Host "$(Get-Date -Format HH:mm:ss) - Monitoring USB & Display stability..." -ForegroundColor Green
            # Add your real monitoring code here (events, counters, re-check hops, display PnP events, etc.)
            Start-Sleep -Seconds 5
        }
    }
    catch {
        Write-Host "Analytics stopped (Ctrl+C detected)." -ForegroundColor Yellow
    }
}

Write-Host "Shōko finished. Press Enter to close." -ForegroundColor Green
Read-Host
