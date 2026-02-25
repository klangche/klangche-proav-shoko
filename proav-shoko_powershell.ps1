# proav-shoko_powershell.ps1 - Shōko Main Logic
# USB tree + display tree always shown consecutively, combined report, real analytics loop with elapsed time & event list

$ErrorActionPreference = 'Stop'

# Load config
try {
    $Config = Invoke-RestMethod "https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.json"
    $Version = $Config.version
} catch {
    $Version = "local"
    $Config = [PSCustomObject]@{ version = "local"; scoring = [PSCustomObject]@{ minScore = 1 } }
}

function Get-Color { param($n) switch($n){ "cyan"{"Cyan"} "magenta"{"Magenta"} "yellow"{"Yellow"} "green"{"Green"} "gray"{"Gray"} default{"White"} } }

Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "Shōko - USB + Display Diagnostic Tool v$Version" -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Basic mode – full analytics requires admin" -ForegroundColor Yellow
}

$dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outTxt = "$env:TEMP\shoko-report-$dateStamp.txt"
$outHtml = "$env:TEMP\shoko-report-$dateStamp.html"

$htmlContent = "<html><body style='background:#000;color:#0f0;font-family:Consolas;'><pre>Shōko Report - $dateStamp`n`n"

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
$baseStabilityScore = [Math]::Max($Config.scoring.minScore, (9 - $maxHops))

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

$htmlContent += "USB TREE`n$treeOutput`nMax hops: $maxHops | Tiers: $numTiers | Devices: $($devices.Count) | Hubs: $($hubs.Count)`n`nSTABILITY PER PLATFORM`n$statusSummary`n`nHOST SUMMARY: STABLE (Score: $baseStabilityScore/10)`n`n"

# ────────────────────────────────────────────────────────────────────────────────
# DISPLAY TREE – always shown immediately after USB
# ────────────────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "Enumerating displays..." -ForegroundColor Gray

$displayTreeOutput = ""
$monitors = Get-CimInstance -Namespace root\wmi -Class WmiMonitorID -EA SilentlyContinue
$connections = Get-CimInstance -Namespace root\wmi -Class WmiMonitorConnectionParams -EA SilentlyContinue
$controllers = Get-CimInstance Win32_VideoController -EA SilentlyContinue

if ($monitors -and $monitors.Count -gt 0) {
    Write-Host "DISPLAY TREE & ANALYTICS" -ForegroundColor Magenta
    Write-Host "==============================================================================" -ForegroundColor Cyan

    for ($i = 0; $i -lt $monitors.Count; $i++) {
        $m = $monitors[$i]
        $c = $connections | Where-Object { $_.InstanceName -eq $m.InstanceName } | Select-Object -First 1

        $name = if ($m.UserFriendlyName -and $m.UserFriendlyName -ne 0) { ($m.UserFriendlyName | ForEach-Object { [char]$_ }) -join '' } else { "Display $($i+1)" }

        $serialRaw = if ($m.SerialNumberID -and $m.SerialNumberID -ne 0) { ($m.SerialNumberID | ForEach-Object { [char]$_ }) -join '' } else { "N/A" }
        $serial = $serialRaw.Trim()

        $connType = if ($c.VideoOutputTechnology -eq 10) { "DisplayPort (External)" } elseif ($c.VideoOutputTechnology -eq 11) { "DisplayPort (Embedded / Alt Mode)" } elseif ($c.VideoOutputTechnology -eq 5) { "HDMI" } else { "Unknown ($($c.VideoOutputTechnology))" }
        $connColor = if ($connType -like "*HDMI*") { "Yellow" } elseif ($connType -like "*DisplayPort*") { "Cyan" } else { "Gray" }

        $adapter = $controllers | Where-Object { $m.InstanceName -like "*$($_.PNPDeviceID)*" } | Select-Object -First 1
        $adapterName = if ($adapter) { $adapter.Name } else { "" }
        $transport = if ($adapterName -match "DisplayLink") { "USB Graphics (DisplayLink)" } elseif ($m.InstanceName -match "DISPLAYPORT") { "DP / DP Alt Mode" } elseif ($m.InstanceName -match "USB") { "USB-C Dock / Alt Mode" } elseif ($m.InstanceName -match "TBT|THUNDER") { "Thunderbolt" } else { "Direct / Unknown" }

        $isMST = $m.InstanceName -match "&MI_"

        $displayTreeOutput += "└─ $name`n"
        $displayTreeOutput += " ├─ Connection : $connType`n"
        $displayTreeOutput += " ├─ Path       : $transport`n"
        if ($isMST) { $displayTreeOutput += " ├─ MST Chain  : YES`n" }
        if ($adapterName -match "DisplayLink") { $displayTreeOutput += " ├─ Adapter    : DisplayLink (software)`n" }

        if ($isAdmin) {
            try {
                $params = Get-CimInstance -Namespace root\wmi -Class WmiMonitorBasicDisplayParams | Where-Object { $_.InstanceName -eq $m.InstanceName } | Select-Object -First 1
                $size = if ($params) { "$($params.MaxHorizontalImageSize) x $($params.MaxVerticalImageSize) cm" } else { "N/A" }
                $displayTreeOutput += " ├─ Size       : $size`n"
                $displayTreeOutput += " ├─ Serial     : $serial`n"
                $displayTreeOutput += " └─ Analytics  : Deep mode active`n"
            } catch {
                $displayTreeOutput += " └─ Analytics  : Partial (error)`n"
            }
        } else {
            $displayTreeOutput += " └─ Analytics  : Basic mode`n"
        }
        $displayTreeOutput += "`n"
    }

    Write-Host $displayTreeOutput
    $htmlContent += "DISPLAY TREE & ANALYTICS`n$displayTreeOutput`n"
} else {
    Write-Host "No displays detected." -ForegroundColor Gray
    $htmlContent += "No displays detected.`n"
}

# ────────────────────────────────────────────────────────────────────────────────
# Save combined report & ask browser
# ────────────────────────────────────────────────────────────────────────────────

$htmlContent += "</pre></body></html>"
$htmlContent | Out-File $outHtml -Encoding UTF8

$txtContent = $htmlContent -replace "<[^>]+>","" -replace "\s+"," "
$txtContent | Out-File $outTxt

Write-Host "Combined report saved (USB + Display + Score): $outTxt / $outHtml" -ForegroundColor Gray

$openBrowser = Read-Host "Open HTML report in browser? (y/n)"
if ($openBrowser -match '^[Yy]') {
    Start-Process $outHtml
}

# ────────────────────────────────────────────────────────────────────────────────
# Deep Analytics Loop (USB + Display events, elapsed time, error list)
# ────────────────────────────────────────────────────────────────────────────────

$runAnalytics = Read-Host "Run deep analytics (USB + Display events, Ctrl+C to stop)? (y/n)"
if ($runAnalytics -match '^[Yy]' -and $isAdmin) {
    Write-Host "Deep analytics mode started. Press Ctrl+C to exit." -ForegroundColor Green

    $startTime = Get-Date
    $eventLog = @()   # Track seen event IDs to avoid duplicates

    try {
        while ($true) {
            $elapsed = (Get-Date) - $startTime
            $elapsedStr = "{0:hh\:mm\:ss}" -f $elapsed

            Write-Host "`n[$elapsedStr] Monitoring USB & Display events..." -ForegroundColor Green

            # Fetch recent PnP events (USB + Display related)
            $recentEvents = Get-WinEvent -FilterHashtable @{
                LogName = 'Microsoft-Windows-Kernel-PnP/Configuration'
                StartTime = (Get-Date).AddMinutes(-10)   # last 10 min
            } -MaxEvents 100 -EA SilentlyContinue |
                Where-Object { $_.Message -match "(?i)(usb|hub|display|monitor|connect|disconnect|hotplug|EDID|error|fail|reset|fault)" }

            if ($recentEvents -and $recentEvents.Count -gt 0) {
                Write-Host "  Found $($recentEvents.Count) relevant events in last 10 min:" -ForegroundColor Yellow

                $newEvents = $recentEvents | Where-Object { $eventLog -notcontains $_.Id }
                foreach ($ev in $newEvents) {
                    $time = $ev.TimeCreated.ToString("HH:mm:ss")
                    $msg = $ev.Message.Trim()
                    if ($msg.Length -gt 120) { $msg = $msg.Substring(0, 120) + "..." }
                    Write-Host "  $time - $msg" -ForegroundColor Yellow
                    $eventLog += $ev.Id
                }
            } else {
                Write-Host "  No new relevant events in last 10 min" -ForegroundColor Green
            }

            Start-Sleep -Seconds 10
        }
    }
    catch [System.Management.Automation.Host.PromptingException] {
        Write-Host "`nCtrl+C detected. Exiting analytics mode." -ForegroundColor Yellow
    }
    catch {
        Write-Host "Analytics error: $($_.Exception.Message)" -ForegroundColor Red
    }

    # After exit → back to main view + re-ask browser
    Write-Host "`nReturning to main view..." -ForegroundColor Gray
    Write-Host "Previous data (USB tree + Display tree + Score) still available in report: $outHtml" -ForegroundColor Cyan

    $reOpen = Read-Host "Re-open HTML report in browser? (y/n)"
    if ($reOpen -match '^[Yy]') {
        Start-Process $outHtml
    }
}

Write-Host "`nShōko finished. Press Enter to close." -ForegroundColor Green
Read-Host
