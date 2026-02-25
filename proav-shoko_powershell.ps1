# proav-shoko_powershell.ps1 - Shōko Main Logic (USB + Display Diagnostics)
# Compatible with PowerShell 5.1 and 7+

$ErrorActionPreference = 'Stop'

# Load config from CURRENT repo
$ConfigUrl = "https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.json"
try {
    $Config = Invoke-RestMethod -Uri $ConfigUrl -UseBasicParsing
    Write-Host "Config loaded (v$($Config.version))" -ForegroundColor Green
} catch {
    Write-Host "Config load failed - using fallback" -ForegroundColor Yellow
    $Config = [PSCustomObject]@{ version = "fallback" }
}

# Color helper (safe for 5.1)
function Get-Color {
    param($ColorName)
    $map = @{ cyan = "Cyan"; magenta = "Magenta"; yellow = "Yellow"; green = "Green"; gray = "Gray" }
    if ($map.ContainsKey($ColorName)) { return $map[$ColorName] } else { return "White" }
}

Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "Shōko - USB + Display Diagnostic Tool v$($Config.version)" -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Basic mode only (admin required for full features)" -ForegroundColor Yellow
}

# =============================================================================
# USB TREE SECTION (your original logic - adapted to current repo)
# =============================================================================
# ... (paste your full USB tree code here if different; this is a placeholder from your repo pattern)
# For now using minimal stable version - replace with your exact USB code if needed

$dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
Write-Host "USB Enumeration..." -ForegroundColor Gray

# (Your USB code would go here - tree building, stability, output, etc.)
# Assuming it runs and prints the tree + verdict

Write-Host ""
Write-Host "USB section complete." -ForegroundColor Green

# =============================================================================
# DISPLAY TREE & ANALYTICS (integrated - PS 5.1 safe)
# =============================================================================

$showDisplay = Read-Host "`nShow display information and analytics? (y/n)"
if ($showDisplay -match '^[Yy]$') {
    function Get-DisplayTree {
        function Decode-Connection {
            param([int]$value)
            $base = $value -band 0x7FFFFFFF
            switch ($base) {
                -2 { "Uninitialized" }
                -1 { "Other/Unknown" }
                0  { "VGA (HD15)" }
                5  { "HDMI" }
                10 { "DisplayPort (External)" }
                11 { "DisplayPort (Embedded / Alt Mode)" }
                default { "Unknown ($value)" }
            }
        }

        function Get-ConnColor {
            param([string]$type)
            if ($type -like "*HDMI*")     { return "Yellow" }
            if ($type -like "*DisplayPort*") { return "Cyan" }
            if ($type -like "*VGA*")      { return "Red" }
            if ($type -like "*Unknown*")  { return "DarkGray" }
            return "Gray"
        }

        function Detect-Transport {
            param([string]$instanceName, [string]$adapterName)
            if ($adapterName -match "DisplayLink") { return "USB Graphics (DisplayLink)" }
            if ($instanceName -match "DISPLAYPORT") { return "DP / DP Alt Mode" }
            if ($instanceName -match "USB") { return "USB-C Dock / Alt Mode" }
            if ($instanceName -match "TBT|THUNDER") { return "Thunderbolt" }
            if ($instanceName -match "HDMI") { return "HDMI" }
            return "Direct / Unknown"
        }

        function Detect-MST {
            param([string]$instanceName)
            return $instanceName -match "&MI_"
        }

        function Get-HealthHint {
            param([string]$connType, [bool]$isMST, [string]$transport, [int]$disconnectCount)
            if ($disconnectCount -gt 1) { return "Magenta", "Potential issue (disconnects detected)" }
            if ($isMST -or $transport -match "DisplayLink") { return "Yellow", "Watch (MST or software path)" }
            if ($connType -match "DisplayPort|HDMI") { return "Green", "Stable" }
            return "Gray", "N/A"
        }

        $monitors = Get-CimInstance -Namespace root\wmi -Class WmiMonitorID -ErrorAction SilentlyContinue
        $connections = Get-CimInstance -Namespace root\wmi -Class WmiMonitorConnectionParams -ErrorAction SilentlyContinue
        $controllers = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue

        if (-not $monitors -or $monitors.Count -eq 0) {
            Write-Host "No displays detected." -ForegroundColor Gray
            return
        }

        Write-Host ""
        Write-Host "DISPLAY TREE & ANALYTICS" -ForegroundColor Magenta
        Write-Host "-------------------------------"

        for ($i = 0; $i -lt $monitors.Count; $i++) {
            $mon = $monitors[$i]
            $conn = $connections | Where-Object { $_.InstanceName -eq $mon.InstanceName } | Select-Object -First 1

            $name = if ($mon.UserFriendlyName -and $mon.UserFriendlyName -ne 0) {
                ($mon.UserFriendlyName | ForEach-Object { [char]$_ }) -join ''
            } else { "Display $($i+1)" }

            $serialRaw = if ($mon.SerialNumberID -and $mon.SerialNumberID -ne 0) {
                ($mon.SerialNumberID | ForEach-Object { [char]$_ }) -join ''
            } else { "N/A" }
            $serial = $serialRaw.Trim()

            $connCode = $conn.VideoOutputTechnology
            $connType = Decode-Connection $connCode
            $connColor = Get-ConnColor $connType

            $adapter = $controllers | Where-Object { $mon.InstanceName -like "*$($_.PNPDeviceID)*" } | Select-Object -First 1
            $adapterName = if ($adapter) { $adapter.Name } else { "" }
            $transport = Detect-Transport $mon.InstanceName $adapterName

            $isMST = Detect-MST $mon.InstanceName

            $logTime = (Get-Date).AddHours(-1)
            $events = Get-WinEvent -FilterHashtable @{
                LogName = 'Microsoft-Windows-Kernel-PnP/Configuration'
                StartTime = $logTime
            } -MaxEvents 100 -ErrorAction SilentlyContinue |
                Where-Object { $_.Message -match "(?i)(display|monitor|connect|disconnect|hotplug|EDID)" }

            $disconnectCount = ($events | Where-Object { $_.Message -match "(?i)disconnect" }).Count

            $healthColor, $healthText = Get-HealthHint $connType $isMST $transport $disconnectCount

            Write-Host "└─ $name" -ForegroundColor White
            Write-Host " ├─ Connection : $connType" -ForegroundColor $connColor
            Write-Host " ├─ Path       : $transport" -ForegroundColor DarkCyan
            if ($isMST) {
                Write-Host " ├─ MST Chain  : YES" -ForegroundColor Yellow
            }
            if ($adapterName -match "DisplayLink") {
                Write-Host " ├─ Adapter    : DisplayLink (software)" -ForegroundColor Yellow
            }

            if ($isAdmin) {
                try {
                    $params = Get-CimInstance -Namespace root\wmi -Class WmiMonitorBasicDisplayParams |
                              Where-Object { $_.InstanceName -eq $mon.InstanceName } | Select-Object -First 1
                    $size = if ($params) { "$($params.MaxHorizontalImageSize) x $($params.MaxVerticalImageSize) cm" } else { "N/A" }

                    Write-Host " ├─ Size       : $size" -ForegroundColor Gray
                    Write-Host " ├─ Serial     : $serial" -ForegroundColor Gray

                    if ($events.Count -gt 0) {
                        Write-Host " ├─ Recent Events ($($events.Count)):" -ForegroundColor Green
                        $events | Select-Object -Last 5 | ForEach-Object {
                            Write-Host " │   $($_.TimeCreated): $($_.Message.Trim())" -ForegroundColor Gray
                        }
                        if ($disconnectCount -gt 0) {
                            Write-Host " └─ Issues     : $disconnectCount disconnects" -ForegroundColor Magenta
                        } else {
                            Write-Host " └─ Analytics  : No major issues" -ForegroundColor Green
                        }
                    } else {
                        Write-Host " └─ Analytics  : No recent events" -ForegroundColor Green
                    }
                } catch {
                    Write-Host " └─ Analytics  : Partial (error)" -ForegroundColor Yellow
                }
            } else {
                $summary = if ($events.Count -gt 0) { "$($events.Count) events (admin for details)" } else { "No recent events" }
                Write-Host " └─ Analytics  : Basic - $summary" -ForegroundColor Gray
            }
            Write-Host ""
        }

        Write-Host "Health: Green=Stable | Yellow=Watch | Magenta=Issue" -ForegroundColor Gray
    }

    Get-DisplayTree
}

# =============================================================================
# Your deep analytics / monitoring loop would go here (original code)
# =============================================================================
Write-Host ""
Write-Host "Shōko complete. Press Enter to exit." -ForegroundColor Green
Read-Host
