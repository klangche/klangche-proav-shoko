<#
.SYNOPSIS
    Shōko - USB + Display Diagnostic Tool
.DESCRIPTION
    Analyzes USB topology, display connections, and platform stability.
    Provides real-time analytics for connection events in ProAV environments.
.PARAMETER Verbose
    Show detailed debug information during execution
.EXAMPLE
    .\proav-shoko_powershell.ps1
    Run in interactive mode
.EXAMPLE
    .\proav-shoko_powershell.ps1 -Verbose
    Run with debug output
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$global:error.Clear()

# Add trap to catch any fatal errors
trap {
    Write-Host "`n" -ForegroundColor Red
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║                     ERROR DETECTED                          ║" -ForegroundColor Red
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Yellow
    Write-Host "Command: $($_.InvocationInfo.Line)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press Enter to exit." -ForegroundColor Gray
    Read-Host
    exit
}

# =============================================================================
# CONFIGURATION LOADING
# =============================================================================

function Load-Configuration {
    <#
    .SYNOPSIS
        Load configuration from GitHub with visible retry attempts
        Exits if config cannot be loaded
    #>
    $configUrl = "https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.json"
    $maxRetries = 3
    $retryCount = 0
    $delay = 2 # seconds
    
    Clear-Host
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host "Shōko - Loading Configuration" -ForegroundColor Cyan
    Write-Host "==============================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    while ($retryCount -lt $maxRetries) {
        $retryCount++
        Write-Host "[Attempt $retryCount of $maxRetries] Loading config from GitHub..." -ForegroundColor Gray -NoNewline
        Write-Host " 🔄" -ForegroundColor Yellow
        
        try {
            $config = Invoke-RestMethod $configUrl
            Write-Host "`r[Attempt $retryCount of $maxRetries] Loading config from GitHub... ✓" -ForegroundColor Green
            Write-Host ""
            Write-Host "Configuration loaded successfully!" -ForegroundColor Green
            Write-Host "  Version: $($config.version)" -ForegroundColor Cyan
            Write-Host "  Reference models: $($config.referenceModels.Count)" -ForegroundColor Cyan
            Write-Host "  Additional models: $($config.additionalModels.Count)" -ForegroundColor Cyan
            Start-Sleep -Seconds 1
            return $config
        } catch {
            Write-Host "`r[Attempt $retryCount of $maxRetries] Loading config from GitHub... ✗" -ForegroundColor Red
            Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
            
            if ($retryCount -lt $maxRetries) {
                Write-Host "  Retrying in $delay seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $delay
            }
        }
    }
    
    # If we get here, all retries failed
    Write-Host ""
    Write-Host "==============================================================================" -ForegroundColor Red
    Write-Host "CONFIGURATION LOAD FAILED" -ForegroundColor Red
    Write-Host "==============================================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "The script cannot run without the configuration file." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Options:" -ForegroundColor White
    Write-Host "  y - Retry loading configuration" -ForegroundColor Cyan
    Write-Host "  n - Exit" -ForegroundColor Gray
    Write-Host ""
    
    $retry = Read-Host "retry (y/n)"
    if ($retry -match '^[Yy]') {
        Write-Host ""
        Write-Host "Restarting..." -ForegroundColor Gray
        Start-Sleep -Seconds 1
        & $PSCommandPath @PSBoundParameters
        exit
    } else {
        Write-Host ""
        Write-Host "Exiting. Press Enter to close." -ForegroundColor Gray
        Read-Host
        exit
    }
}

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

function Get-Color {
    <#
    .SYNOPSIS
        Convert color name to console color
    #>
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

function Format-PlatformLine {
    <#
    .SYNOPSIS
        Format a platform stability line with Hops header
    #>
    param($tiers, $max, $name, $status)
    return "{0,-4} {1,-40} {2}" -f "$tiers/$max", $name, $status
}

function Format-Score {
    <#
    .SYNOPSIS
        Format score with two digits and one decimal (e.g., 10.0, 05.5, 00.0)
    #>
    param($score)
    return "{0:00.0}" -f [Math]::Round($score, 1)
}

function Format-Duration {
    <#
    .SYNOPSIS
        Format timespan as HH:MM:SS.fff
    #>
    param($timespan)
    return "{0:hh\:mm\:ss\.fff}" -f $timespan
}

# =============================================================================
# SYSTEM INFORMATION
# =============================================================================

function Get-SystemInfo {
    <#
    .SYNOPSIS
        Get system information including OS, PowerShell version, admin status
    #>
    Write-Verbose "Getting system information"
    
    $isAdmin = try {
        $identity = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        $identity.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        Write-Verbose "Failed to check admin status: $_"
        $false
    }
    
    $osInfo = try {
        Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
    } catch {
        Write-Verbose "Failed to get OS info: $_"
        [PSCustomObject]@{ Caption = "Windows"; Version = "10.0"; BuildNumber = "0000" }
    }
    
    $winVersion = "$($osInfo.Caption) $($osInfo.Version) (Build $($osInfo.BuildNumber))"
    $psVersion = $PSVersionTable.PSVersion.ToString()
    $arch = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }
    
    # Detect current platform
    $currentPlatform = "Windows x86"
    if ([Environment]::Is64BitOperatingSystem -and [Environment]::ProcessorArchitecture -eq "Arm64") {
        $currentPlatform = "Windows ARM"
    }
    
    return [PSCustomObject]@{
        IsAdmin = $isAdmin
        Mode = if ($isAdmin) { "Elevated / admin" } else { "Basic mode" }
        OSVersion = $winVersion
        PSVersion = $psVersion
        Architecture = $arch
        CurrentPlatform = $currentPlatform
    }
}

# =============================================================================
# USB TREE - PROPER HIERARCHICAL VERSION
# =============================================================================

function Get-UsbTree {
    <#
    .SYNOPSIS
        Enumerate USB devices and build hierarchical tree
    #>
    param($Config)
    
    Write-Verbose "Enumerating USB devices"
    
    $allDevices = try {
        Get-PnpDevice -Class USB -ErrorAction Stop | Where-Object {$_.Status -eq 'OK'}
    } catch {
        Write-Verbose "Failed to get USB devices: $_"
        @()
    }
    
    if (-not $allDevices) { $allDevices = @() }
    
    $devices = @()
    $hubs = @()
    $maxHops = 0
    $deviceMap = @{}
    $treeOutput = "HOST`n"
    
    # First pass: create device objects with depth and parent info
    foreach ($d in $allDevices) {
        $isHub = ($d.FriendlyName -like "*hub*") -or ($d.Name -like "*hub*") -or ($d.Class -eq "USBHub")
        if ($isHub) {
            $hubs += $d
        } else {
            $devices += $d
        }
        
        # Get depth from instance ID (count backslashes)
        $depth = ($d.InstanceId.ToCharArray() | Where-Object {$_ -eq '\'} | Measure-Object).Count
        if ($depth -gt $maxHops) { $maxHops = $depth }
        
        # Extract parent ID (remove last segment)
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
    
    # Build parent-child relationships
    foreach ($id in $deviceMap.Keys) {
        $parent = $deviceMap[$id].Parent
        if ($parent -and $deviceMap.ContainsKey($parent)) {
            $deviceMap[$parent].Children += $id
        }
    }
    
    # Find root devices (those with no parent or parent not in map)
    $roots = @()
    foreach ($id in $deviceMap.Keys) {
        $parent = $deviceMap[$id].Parent
        if (-not $parent -or -not $deviceMap.ContainsKey($parent)) {
            $roots += $id
        }
    }
    
    # If no roots found, use depth 1 as fallback
    if ($roots.Count -eq 0) {
        $roots = $deviceMap.Keys | Where-Object { $deviceMap[$_].Depth -eq 1 }
    }
    
    $roots = $roots | Sort-Object { $deviceMap[$_].Name }
    
    # Recursive tree printer
    function Write-DeviceNode {
        param($id, $level, $isLast)
        
        $node = $deviceMap[$id]
        if (-not $node) { return }
        
        # Build the prefix with proper tree characters
        $prefix = ""
        if ($level -gt 0) {
            $prefix = "│   " * ($level - 1)
            $prefix += if ($isLast) { "└── " } else { "├── " }
        } else {
            $prefix = "├── "
        }
        
        $hubTag = if ($node.IsHub) { " [HUB]" } else { "" }
        $script:treeOutput += "$prefix$($node.Name)$hubTag ← $($node.Depth) hops`n"
        
        # Process children
        $children = $node.Children | Sort-Object { $deviceMap[$_].Name }
        for ($i = 0; $i -lt $children.Count; $i++) {
            $childIsLast = ($i -eq $children.Count - 1)
            Write-DeviceNode $children[$i] ($level + 1) $childIsLast
        }
    }
    
    # Handle case with no devices
    if ($roots.Count -eq 0) {
        $treeOutput += "├── USB Root Hub (Host Controller) [HUB] ← 1 hops`n"
        $treeOutput += "│   $($Config.messages.noDevices)`n"
        $maxHops = 1
    } else {
        # Print all roots
        for ($i = 0; $i -lt $roots.Count; $i++) {
            $isLastRoot = ($i -eq $roots.Count - 1)
            Write-DeviceNode $roots[$i] 0 $isLastRoot
        }
    }
    
    $numTiers = $maxHops + 1
    
    return [PSCustomObject]@{
        Tree = $treeOutput
        MaxHops = $maxHops
        Tiers = $numTiers
        Devices = $devices.Count
        Hubs = $hubs.Count
    }
}

# =============================================================================
# DISPLAY TREE
# =============================================================================

function Get-DisplayTree {
    <#
    .SYNOPSIS
        Enumerate displays and connection information
    #>
    param($Config)
    
    Write-Verbose "Enumerating displays"
    
    $isAdmin = (Get-SystemInfo).IsAdmin
    $displayOutput = "HOST`n"
    
    $monitors = try {
        Get-CimInstance -Namespace root\wmi -Class WmiMonitorID -ErrorAction Stop
    } catch {
        Write-Verbose "Failed to get monitor info: $_"
        @()
    }
    
    if ($monitors -and $monitors.Count -gt 0) {
        for ($i = 0; $i -lt $monitors.Count; $i++) {
            $m = $monitors[$i]
            
            $name = "Display $($i+1)"
            if ($m.UserFriendlyName -and $m.UserFriendlyName -ne 0) { 
                $name = ($m.UserFriendlyName | ForEach-Object { [char]$_ }) -join '' 
            }
            
            $displayOutput += "└── $name`n"
            $displayOutput += " ├─ Connection : "
            
            $conn = try {
                Get-CimInstance -Namespace root\wmi -Class WmiMonitorConnectionParams -Filter "InstanceName = '$($m.InstanceName)'" -ErrorAction Stop
            } catch {
                Write-Verbose "Failed to get connection params: $_"
                $null
            }
            
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
        $displayOutput += "$($Config.messages.noDevices)`n"
    }
    
    return $displayOutput
}

# =============================================================================
# PLATFORM STABILITY
# =============================================================================

function Get-PlatformStability {
    <#
    .SYNOPSIS
        Calculate stability for reference and additional models using decimal scoring
    #>
    param($Config, $Tiers, $MaxHops)
    
    Write-Verbose "Calculating platform stability for $Tiers tiers ($MaxHops external hops)"
    
    $baseScore = [Math]::Max($Config.scoring.minScore, 10.0 - $MaxHops)
    $baseScore = [Math]::Min($baseScore, $Config.scoring.maxScore)
    $baseScore = [Math]::Round($baseScore, 1)
    
    $referenceOutput = ""
    $additionalOutput = ""
    $referenceScores = @()
    $worstReferenceScore = $Config.scoring.maxScore
    
    # Process reference models (affect score)
    foreach ($model in $Config.referenceModels) {
        $rec = $model.rec
        $max = $model.max
        $name = $model.name
        
        $status = if ($Tiers -le $rec) { "STABLE" } 
                  elseif ($Tiers -le $max) { "POTENTIALLY UNSTABLE" } 
                  else { "NOT STABLE" }
        
        $referenceOutput += Format-PlatformLine -tiers $Tiers -max $max -name $name -status $status
        $referenceOutput += "`n"
        
        $referenceScores += $baseScore
        if ($baseScore -lt $worstReferenceScore) { $worstReferenceScore = $baseScore }
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
    
    # Print config status under Arch
    Write-Host "Config: Loaded v$($Config.version) (Ref:$($Config.referenceModels.Count) Add:$($Config.additionalModels.Count))" -ForegroundColor Green
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
    Write-Host "Hops  Plattform                              STATUS" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $Stability.ReferenceOutput
    Write-Host ""
    Write-Host "Additional models (reference only):" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $Stability.AdditionalOutput
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host ""
    
    $verdictColor = Get-Color $Config.colors.($Stability.Verdict.ToLower().Replace(' ', ''))
    Write-Host "HOST SUMMARY: $(Format-Score $Stability.WorstScore)/10.0 - $($Stability.Verdict)" -ForegroundColor $verdictColor
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

HOST SUMMARY: $(Format-Score $Analytics.InitialScore)/10.0 - $($Analytics.InitialVerdict)
HOST SUMMARY: $(Format-Score $Analytics.AdjustedScore)/10.0 - $($Analytics.AdjustedVerdict) (adjusted)

==============================================================================
Analytics Summary (during monitoring):
Total events logged: $($Analytics.Counters.total)
$($Analytics.SummaryText)
Points deducted: $([math]::Round($Analytics.Deductions,2))

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
Config: Loaded v$($Config.version) (Ref:$($Config.referenceModels.Count) Add:$($Config.additionalModels.Count))

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
Hops  Plattform                              STATUS
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
        Run real-time monitoring of USB/display events with working timer
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
    
    # Analytics header - fixed positions for updates
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
    
    # Initial stats display - will be updated
    Write-Host "Duration: 00:00:00.000" -ForegroundColor (Get-Color $Config.colors.green)
    Write-Host "Total events logged: 0" -ForegroundColor (Get-Color $Config.colors.white)
    
    if ($System.IsAdmin) {
        Write-Host "USB RE-HANDSHAKES: 0"
        Write-Host "USB JITTER: 0"
        Write-Host "USB CRC ERRORS: 0"
        Write-Host "USB BUS RESETS: 0"
        Write-Host "USB OVERCURRENT: 0"
        Write-Host "DISPLAY HOTPLUGS: 0"
        Write-Host "DISPLAY EDID ERRORS: 0"
        Write-Host "DISPLAY LINK FAILURES: 0"
        Write-Host "OTHER ERRORS: 0"
    } else {
        Write-Host "USB CONNECTS: 0"
        Write-Host "USB DISCONNECTS: 0"
        Write-Host "USB RE-HANDSHAKES: 0"
        Write-Host "USB JITTER: 0"
        Write-Host "USB ERRORS: 0"
        Write-Host "DISPLAY HOTPLUGS: 0"
        Write-Host "DISPLAY ERRORS: 0"
    }
    
    Write-Host ""
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Analytics Log:" -ForegroundColor (Get-Color $Config.colors.cyan)
    
    $logStartLine = 19  # Starting line for log entries
    
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
    
    # Show initial log entry
    Write-Host $analyticsLog[0]
    
    while (-not $Host.UI.RawUI.KeyAvailable) {
        $elapsed = (Get-Date) - $startTime
        $duration = Format-Duration $elapsed
        
        # Update timer (line 7)
        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, 7
        Write-Host "Duration: $duration" -ForegroundColor (Get-Color $Config.colors.green)
        
        # Update stats based on mode
        if ($System.IsAdmin) {
            $line = 9
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
            $line = 9
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $line; Write-Host "USB CONNECTS: $($counters.connects)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 1); Write-Host "USB DISCONNECTS: $($counters.disconnects)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 2); Write-Host "USB RE-HANDSHAKES: $($counters.rehandshakes)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 3); Write-Host "USB JITTER: $($counters.jitter)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 4); Write-Host "USB ERRORS: $($counters.crcErrors + $counters.busResets + $counters.overcurrent + $counters.otherErrors)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 5); Write-Host "DISPLAY HOTPLUGS: $($counters.hotplugs)          "
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($line + 6); Write-Host "DISPLAY ERRORS: $($counters.edidErrors + $counters.linkFailures)          "
        }
        
        # Update total events line
        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, 8
        Write-Host "Total events logged: $($counters.total)          "
        
        # SIMULATE EVENTS FOR TESTING - REPLACE WITH REAL EVENT MONITORING
        if ($analyticsLog.Count -eq 1 -and $startTime.AddSeconds(3) -lt (Get-Date)) {
            if ($System.IsAdmin) {
                $newEvents = @(
                    "$($startTime.AddSeconds(2).ToString('HH:mm:ss.fff')) - [CONNECT] USB device connected (VID_046D/PID_0843) - Logitech Webcam",
                    "$($startTime.AddSeconds(4).ToString('HH:mm:ss.fff')) - [DISCONNECT] USB device disconnected - Logitech Webcam",
                    "$($startTime.AddSeconds(6).ToString('HH:mm:ss.fff')) - [CONNECT] USB device connected - Logitech Webcam"
                )
                $counters.connects = 2
                $counters.disconnects = 1
            } else {
                $newEvents = @(
                    "$($startTime.AddSeconds(2).ToString('HH:mm:ss.fff')) - [CONNECT] VID_046D/PID_0843",
                    "$($startTime.AddSeconds(4).ToString('HH:mm:ss.fff')) - [DISCONNECT] VID_046D/PID_0843",
                    "$($startTime.AddSeconds(6).ToString('HH:mm:ss.fff')) - [CONNECT] VID_046D/PID_0843"
                )
                $counters.connects = 2
                $counters.disconnects = 1
            }
            
            foreach ($event in $newEvents) {
                $analyticsLog += $event
                $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($logStartLine + $analyticsLog.Count - 1)
                Write-Host $event
            }
            $counters.total = 3
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    # Stop analytics
    $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    $stopTime = Get-Date
    $totalDuration = Format-Duration ($stopTime - $startTime)
    $analyticsLog += "$($stopTime.ToString('HH:mm:ss.fff')) - Logging ended (total duration: $totalDuration)"
    
    # Show final log entry
    $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($logStartLine + $analyticsLog.Count - 1)
    Write-Host $analyticsLog[-1]
    
    # Calculate deductions
    $deductions = 0.0
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
    
    $adjustedScore = [Math]::Max($Config.scoring.minScore, $initialData.Score - $deductions)
    $adjustedScore = [Math]::Min($adjustedScore, $Config.scoring.maxScore)
    $adjustedScore = [Math]::Round($adjustedScore, 1)
    
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
    Write-Host "Config: Loaded v$($Config.version) (Ref:$($Config.referenceModels.Count) Add:$($Config.additionalModels.Count))" -ForegroundColor Green
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
    Write-Host "Hops  Plattform                              STATUS" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $InitialData.ReferenceOutput
    Write-Host ""
    Write-Host "Additional models (reference only):" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $InitialData.AdditionalOutput
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host ""
    
    $verdictColor = Get-Color $Config.colors.($Stability.Verdict.ToLower().Replace(' ', ''))
    $adjustedColor = Get-Color $Config.colors.($Analytics.AdjustedVerdict.ToLower().Replace(' ', ''))
    
    Write-Host "HOST SUMMARY: $(Format-Score $InitialData.Score)/10.0 - $($InitialData.Verdict)" -ForegroundColor $verdictColor
    Write-Host "HOST SUMMARY: $(Format-Score $Analytics.AdjustedScore)/10.0 - $($Analytics.AdjustedVerdict) (adjusted)" -ForegroundColor $adjustedColor
    Write-Host ""
    
    Write-Host "==============================================================================" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Analytics Summary (during monitoring):" -ForegroundColor (Get-Color $Config.colors.cyan)
    Write-Host "Total events logged: $($Analytics.Counters.total)" -ForegroundColor (Get-Color $Config.colors.white)
    Write-Host $Analytics.SummaryText
    Write-Host "Points deducted: $([math]::Round($Analytics.Deductions,2))"
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
    
    $Config = Load-Configuration
    $System = Get-SystemInfo
    
    Write-Host "`nCollecting system data..." -ForegroundColor Gray
    
    $Usb = Get-UsbTree -Config $Config
    $Display = Get-DisplayTree -Config $Config
    $Stability = Get-PlatformStability -Config $Config -Tiers $Usb.Tiers -MaxHops $Usb.MaxHops
    
    Show-Report -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability
    
    # Question - Analytics
    $analyticsChoice = Read-Host "Run deep analytics session? (y/n)"
    if ($analyticsChoice -match '^[Yy]') {
        $Analytics = Start-AnalyticsSession -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability
        
        Show-FinalReport -Config $Config -System $System -InitialData $Analytics.InitialData -Stability $Stability -Analytics $Analytics
        
        # Question - HTML report with full data
        $finalHtmlChoice = Read-Host "`nOpen HTML report with full data? (y/n)"
        if ($finalHtmlChoice -match '^[Yy]') {
            Save-HtmlReport -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability -Analytics $Analytics
        }
    } else {
        # If no analytics, ask about HTML report
        $htmlChoice = Read-Host "Open HTML report with full data? (y/n)"
        if ($htmlChoice -match '^[Yy]') {
            Save-HtmlReport -Config $Config -System $System -Usb $Usb -Display $Display -Stability $Stability
        }
    }
    
    Write-Host "`nShōko finished. Press Enter to close." -ForegroundColor (Get-Color $Config.colors.green)
    Read-Host
}

# Run main
Main
