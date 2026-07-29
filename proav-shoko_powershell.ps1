<#
.SYNOPSIS
    ProAV Shoko - USB + Display Diagnostic Tool (PowerShell CLI)
.DESCRIPTION
    100% PowerShell implementation matching python run.py --cli behavior.
    Analyzes USB topology, display connections, and platform stability
    with per-platform verdicts using USB spec limits from CSV data.
    Defaults to elevated (admin) mode for maximum data.
.EXAMPLE
    .\proav-shoko_powershell.ps1
    Run in interactive mode
#>

# ============================================================================
# BAKED-IN USB DATA - UPDATE ON EVERY PUSH TO MAIN
# ============================================================================
$script:BakedInCsv = @"
system,platform,arch,max_hops,max_tiers,max_hubs,description,notes,source
Windows (x86/AMD64),windows_x86,x86,7,7,5,"Windows (Intel/AMD)","Full USB 2.0/3.x spec supported. Safe ProAV limit = 5 powered external hubs. Unpowered hubs drastically reduce depth. Endpoint limits affect total device count, not depth.","xHCI Spec: https://www.intel.com/content/dam/www/public/us/en/documents/technical-specifications/extensible-host-controller-interface-usb-xhci.pdf, USB 2.0 Spec: https://patentimages.storage.googleapis.com/2c/17/da/f84081217c2868/US20060020736A1.pdf"
Windows (ARM),windows_arm,arm,7,7,5,"Windows (ARM / Snapdragon)","Snapdragon X / SQ series use standard xHCI controllers. Safe ProAV limit = 5 powered hubs. Same as x86 Windows.","xHCI Spec: https://www.intel.com/content/dam/www/public/us/en/documents/technical-specifications/extensible-host-controller-interface-usb-xhci.pdf"
macOS (Intel),mac_intel,x86,7,7,5,"Mac (Intel)","Full USB spec support. No internal tier penalty (unlike Apple Silicon). Safe ProAV limit = 5 powered hubs.","USB 2.0 Spec: https://patentimages.storage.googleapis.com/2c/17/da/f84081217c2868/US20060020736A1.pdf"
macOS (Apple Silicon),mac_apple_silicon,arm,6,6,4,"Mac (Apple Silicon)","Internal Thunderbolt/USB hub consumes 1 tier. Safe ProAV limit = 4 powered external hubs. 5 may cause bus timeouts on some M2/M3 models.","Biamp Tech Doc: https://support.biamp.com/@api/deki/pages/9655/pdf/EasyConnect%2bMPX%2b250%2b-%2bUSB%2bhub%2blimitations.pdf#1#1, Community: https://forums.macrumors.com/threads/best-way-to-extend-usb-or-thunderbolt.2484510/"
Linux (x86/AMD64),linux_x86,x86,7,7,5,"Linux (Intel/AMD)","Kernel supports full USB spec. Safe ProAV limit = 5 powered hubs. Identical to Windows for client device usage.","xHCI Spec: https://www.intel.com/content/dam/www/public/us/en/documents/technical-specifications/extensible-host-controller-interface-usb-xhci.pdf, Linux Kernel Commit: https://git.zx2c4.com/linux-dev/commit/drivers/usb/core/usb.c?id=4a0cd9670f22c308bc5936ee9734d8ee3f1baa52"
Linux (ARM),linux_arm,arm,6,6,4,"Linux (ARM / RPi, Rockchip)","Kernel supports 7 tiers, but many ARM SoCs (Raspberry Pi, Rockchip) have firmware caps. Safe ProAV limit = 4 powered hubs. Rare client device.","xHCI Spec: https://www.intel.com/content/dam/www/public/us/en/documents/technical-specifications/extensible-host-controller-interface-usb-xhci.pdf, Raspberry Pi Forum: https://forums.raspberrypi.com/viewtopic.php?t=390587"
iPhone (USB-C),iphone_usbc,arm,4,4,2,"iPhone (USB-C)","iOS power management aggressively caps depth. Safe ProAV limit = 2 powered external hubs. Do not attempt 3 in a room system.","Community: https://forums.macrumors.com/threads/best-way-to-extend-usb-or-thunderbolt.2484510/"
Android (USB-C),android_usbc,arm,5,5,3,"Android (USB-C - Samsung, Pixel, OnePlus)","Vendor firmware (Qualcomm/Exynos/Tensor) varies. Safe ProAV limit = 3 powered hubs covers all major brands.","xHCI Spec: https://www.intel.com/content/dam/www/public/us/en/documents/technical-specifications/extensible-host-controller-interface-usb-xhci.pdf, Community: https://forums.macrumors.com/threads/best-way-to-extend-usb-or-thunderbolt.2484510/"
iPad (USB-C),ipad_usbc,arm,5,5,3,"iPad (USB-C)","M-series iPads can handle 4 hubs, A-series (Air/mini) cap at 2. Safe ProAV limit = 3 powered hubs covers 100% of USB-C iPads.","Community: https://forums.macrumors.com/threads/best-way-to-extend-usb-or-thunderbolt.2484510/, Satechi Compatibility: https://support.satechi.com/hc/en-us/articles/39712604458907-Do-Satechi-products-work-with-the-newest-Apple-silicon-M1-M2-M3-chips"
"@

# ============================================================================
# PLATFORM DEFINITIONS (mirrors Python _PLATFORMS list)
# ============================================================================
$script:PlatformDefs = @(
    @{ id = 'windows_x86';        name = 'Windows';          arch = 'x86/x64' }
    @{ id = 'mac_intel';          name = 'Mac Intel';        arch = 'x86/x64' }
    @{ id = 'linux_x86';          name = 'Linux';            arch = 'x86/x64' }
    @{ id = 'windows_arm';        name = 'Windows';          arch = 'ARM'     }
    @{ id = 'mac_apple_silicon';  name = 'Mac Apple Silicon'; arch = 'ARM'    }
    @{ id = 'linux_arm';          name = 'Linux';            arch = 'ARM'     }
    @{ id = 'iphone_usbc';        name = 'iPhone USB-C';     arch = 'Mobile'  }
    @{ id = 'android_usbc';       name = 'Android USB-C';    arch = 'Mobile'  }
    @{ id = 'ipad_usbc';          name = 'iPad USB-C';       arch = 'Mobile'  }
)

$script:StatusRank = @{ 'STABLE' = 0; 'AT LIMIT' = 1; 'UNSTABLE' = 2 }
$script:InternalKeywords = @('integrated', 'bluetooth', 'camera', 'intel', 'razer', 'hid event', 'hid filter', 'i2c', 'surface', 'dell dock', 'thinkpad', 'realtek')
$script:Tee = '├── '  # tree branch
$script:TeeEnd = '└── ' # tree last branch
$script:TeeBar = '│'    # tree vertical bar

# ============================================================================
# CSV LOADING
# ============================================================================
function Load-UsbData {
    param([string]$CsvUrl)
    $hopLimits = @{}; $tierLimits = @{}; $hubLimits = @{}
    $platformNotes = @(); $csvContent = $null
    if ($CsvUrl) {
        try {
            Write-Verbose "Loading CSV from: $CsvUrl"
            $response = Invoke-WebRequest -Uri $CsvUrl -UseBasicParsing -TimeoutSec 10
            $csvContent = $response.Content
        } catch { Write-Verbose "Failed to load CSV from URL: $_" }
    }
    if (-not $csvContent) {
        Write-Verbose "Using baked-in CSV data"
        $csvContent = $script:BakedInCsv
    }
    try {
        $lines = $csvContent -split "`n" | Where-Object { $_.Trim() -ne '' }
        $header = $lines[0] -split ','
        $dataRows = $lines[1..($lines.Count - 1)]
        $platformCol = [array]::IndexOf($header, 'platform')
        $hopsCol = [array]::IndexOf($header, 'max_hops')
        $tiersCol = [array]::IndexOf($header, 'max_tiers')
        $hubsCol = [array]::IndexOf($header, 'max_hubs')
        $systemCol = [array]::IndexOf($header, 'system')
        $descCol = [array]::IndexOf($header, 'description')
        $notesCol = [array]::IndexOf($header, 'notes')
        foreach ($line in $dataRows) {
            $cols = Parse-CsvLine $line
            $platform = $cols[$platformCol].Trim()
            if (-not $platform) { continue }
            $hops = 0; $tiers = 0; $hubs = 0
            if ([int]::TryParse($cols[$hopsCol].Trim(), [ref]$hops)) { }
            if ([int]::TryParse($cols[$tiersCol].Trim(), [ref]$tiers)) { }
            $hubStr = if ($hubsCol -ge 0) { $cols[$hubsCol].Trim() } else { '0' }
            if (-not [int]::TryParse($hubStr, [ref]$hubs)) { $hubs = 0 }
            $system = if ($systemCol -ge 0) { $cols[$systemCol].Trim() } else { '' }
            $description = if ($descCol -ge 0) { $cols[$descCol].Trim() } else { '' }
            $notes = if ($notesCol -ge 0) { $cols[$notesCol].Trim() } else { '' }
            if ($hops -gt 0) { $hopLimits[$platform] = $hops }
            if ($tiers -gt 0) { $tierLimits[$platform] = $tiers }
            if ($hubs -gt 0) { $hubLimits[$platform] = $hubs }
            $platformNotes += @{ platform = $platform; system = $system; description = $description; note = $notes }
        }
    } catch { Write-Verbose "Error parsing CSV: $_" }
    if ($hopLimits.Count -eq 0) {
        $hopLimits = @{ 'windows_x86' = 7; 'windows_arm' = 7; 'mac_intel' = 7; 'mac_apple_silicon' = 6; 'linux_x86' = 7; 'linux_arm' = 6; 'iphone_usbc' = 4; 'android_usbc' = 5; 'ipad_usbc' = 5 }
        $tierLimits = @{ 'windows_x86' = 7; 'windows_arm' = 7; 'mac_intel' = 7; 'mac_apple_silicon' = 6; 'linux_x86' = 7; 'linux_arm' = 6; 'iphone_usbc' = 4; 'android_usbc' = 5; 'ipad_usbc' = 5 }
        $hubLimits  = @{ 'windows_x86' = 5; 'windows_arm' = 5; 'mac_intel' = 5; 'mac_apple_silicon' = 4; 'linux_x86' = 5; 'linux_arm' = 4; 'iphone_usbc' = 2; 'android_usbc' = 3; 'ipad_usbc' = 3 }
    }
    return @{ hopLimits = $hopLimits; tierLimits = $tierLimits; hubLimits = $hubLimits; platformNotes = $platformNotes }
}

function Parse-CsvLine {
    param([string]$line)
    $cols = @(); $current = ""; $inQuotes = $false
    for ($i = 0; $i -lt $line.Length; $i++) {
        $c = $line[$i]
        if ($c -eq '"') { $inQuotes = -not $inQuotes }
        elseif ($c -eq ',' -and -not $inQuotes) { $cols += $current; $current = "" }
        else { $current += $c }
    }
    $cols += $current
    return $cols
}

# ============================================================================
# PLATFORM INFO
# ============================================================================
function Get-PlatformInfo {
    $isAdmin = $false
    try {
        $identity = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        $isAdmin = $identity.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { }
    $osInfo = try { Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop } catch { $null }
    $osCaption = if ($osInfo) { $osInfo.Caption } else { "Windows" }
    $osVersion = if ($osInfo) { $osInfo.Version } else { "10.0" }
    $osBuild = if ($osInfo) { $osInfo.BuildNumber } else { "0000" }
    $arch = if ([Environment]::Is64BitOperatingSystem) { "AMD64" } else { "x86" }
    if ([Environment]::Is64BitOperatingSystem -and [Environment]::ProcessorArchitecture -eq 'Arm64') { $arch = "ARM64" }
    return @{ os = $osCaption; version = $osVersion; build = $osBuild; architecture = $arch; isAdmin = $isAdmin; isAppleSilicon = $false; versionString = "$osCaption $osVersion (Build $osBuild)" }
}

# ============================================================================
# USB TREE
# ============================================================================
function Get-UsbTree {
    $ErrorActionPreference = 'SilentlyContinue'
    $pnpDevices = Get-PnpDevice -PresentOnly | Where-Object { $_.Class -eq 'USB' -or $_.Class -eq 'USBHub' -or $_.Class -eq 'HIDClass' -or $_.InstanceId -like 'USB\*' -or $_.InstanceId -like 'USBSTOR\*' }
    $deviceMap = @{}; $parentMap = @{}
    foreach ($d in $pnpDevices) {
        $model = if ($d.FriendlyName) { $d.FriendlyName } else { if ($d.Name) { $d.Name } else { 'Unknown' } }
        $modelLower = $model.ToLower()
        if ($modelLower -like '*host controller*' -or $modelLower -like '*root hub*') { continue }
        $isHub = ($d.Class -eq 'USBHub') -or ($modelLower -like '*hub*') -or ($modelLower -like '*nav*')
        $vid = ''; $devPid = ''; $miNum = $null; $isComposite = $false
        if ($d.InstanceId -match 'VID_([0-9A-F]{4})') { $vid = $matches[1] }
        if ($d.InstanceId -match 'PID_([0-9A-F]{4})') { $devPid = $matches[1] }
        if ($d.InstanceId -match 'MI_(\d{2})') { $miNum = [int]$matches[1]; $isComposite = $true }
        $deviceInfo = if ($vid -and $devPid) { "VID_${vid}&PID_${devPid}" } else { '' }
        $portNum = 0
        if ($d.InstanceId -match '&(\d+)$') { $portNum = [int]$matches[1] }
        $ifaceDesc = ''
        if ($d.Class -eq 'HIDClass') {
            if ($modelLower -like '*keyboard*') { $ifaceDesc = 'HID Keyboard' }
            elseif ($modelLower -like '*mouse*') { $ifaceDesc = 'HID Mouse' }
            else { $ifaceDesc = 'HID' }
        }
        $deviceMap[$d.InstanceId] = @{ name = $d.InstanceId; model = $model; vendor = ''; vid = $vid; pid = $devPid; deviceInfo = $deviceInfo; isHub = $isHub; isComposite = $isComposite; interfaceNumber = $miNum; interfaceDesc = $ifaceDesc; isInternal = $false; depth = 0; hops = 0; port = $portNum; children = @(); instanceId = $d.InstanceId }
    }
    foreach ($id in $deviceMap.Keys) {
        try { $prop = Get-PnpDeviceProperty -InstanceId $id -KeyName "DEVPKEY_Device_Parent" -ErrorAction Stop; if ($prop.Data) { $parentMap[$id] = $prop.Data.Trim() } } catch { }
    }
    $rootNodes = @(); $assigned = @{}
    foreach ($id in $deviceMap.Keys) {
        $parentId = $parentMap[$id]
        if ($parentId -and $deviceMap.ContainsKey($parentId)) { $deviceMap[$parentId].children += $id; $assigned[$id] = $true }
        else { $rootNodes += $id; $assigned[$id] = $true }
    }
    foreach ($id in $deviceMap.Keys) {
        $node = $deviceMap[$id]
        $combined = "$($node.model) $($node.vendor)".ToLower()
        $node.isInternal = ($script:InternalKeywords | Where-Object { $combined -match $_ }).Count -gt 0
    }
    function Build-NestedNode { param($id, $nodes, $depth = 0)
        $n = $nodes[$id]
        $result = @{ name = $n.name; model = $n.model; vendor = $n.vendor; vid = $n.vid; pid = $n.pid; deviceInfo = $n.deviceInfo; isHub = $n.isHub; isComposite = $n.isComposite; interfaceNumber = $n.interfaceNumber; interfaceDesc = $n.interfaceDesc; isInternal = $n.isInternal; depth = $depth; hops = $depth; port = $n.port; children = @(); instanceId = $n.instanceId }
        foreach ($childId in $n.children) { $result.children += Build-NestedNode $childId $nodes ($depth + 1) }
        return $result
    }
    $nestedRoots = @()
    foreach ($rid in $rootNodes) { $nestedRoots += Build-NestedNode $rid $deviceMap 1 }
    function Group-CompanionHubs {
        param($roots)
        $groups = @{}
        for ($gi = 0; $gi -lt $roots.Count; $gi++) {
            $r = $roots[$gi]
            if (-not $r.isHub) { continue }
            $inst = $r.instanceId
            if (-not $inst) { continue }
            $lastSeg = $inst.Split('\')[-1]
            if ($lastSeg -match '^MSFT(\d+)') {
                $base = $lastSeg.Substring($matches[0].Length)
                if (-not $groups.ContainsKey($base)) { $groups[$base] = @() }
                $groups[$base] += @{ index = $gi; node = $r }
            }
        }
        if ($groups.Count -eq 0) { return $roots }
        $replaced = @{}; $newRoots = @()
        foreach ($base in $groups.Keys) {
            $entries = $groups[$base]
            if ($entries.Count -lt 2) { continue }
            $groupNodes = $entries | ForEach-Object { $_.node }
            $label = "USB Hub Group"
            foreach ($gn in $groupNodes) {
                $foundLabel = $false
                foreach ($child in $gn.children) {
                    $cm = if ($child.model) { $child.model.ToLower() } else { '' }
                    if ($cm -match 'multiport|adapter') { $label = $child.model; $foundLabel = $true; break }
                }
                if ($foundLabel) { break }
            }
            if ($label -eq "USB Hub Group" -and $groupNodes[0].vid) { $label = "USB-C Adapter (VID_$($groupNodes[0].vid))" }
            $virtual = @{ name = $label; model = $label; vendor = 'Generic'; vid = ''; pid = ''; deviceInfo = ''; isHub = $false; isComposite = $false; isVirtual = $true; interfaceNumber = $null; interfaceDesc = ''; isInternal = $false; depth = 1; hops = 1; port = 0; instanceId = "\virtual_$base"; children = @() }
            foreach ($gn in $groupNodes) { $virtual.children += $gn }
            $newRoots += $virtual
            foreach ($entry in $entries) { $replaced[$entry.index] = $true }
        }
        for ($gi = 0; $gi -lt $roots.Count; $gi++) { if (-not $replaced.ContainsKey($gi)) { $newRoots += $roots[$gi] } }
        return ,$newRoots
    }
    $nestedRoots = Group-CompanionHubs $nestedRoots
    function Propagate-InternalDown {
        param($node)
        if ($node.isInternal -and $node.children -and $node.children.Count -gt 0) {
            foreach ($child in $node.children) { $child.isInternal = $true; Propagate-InternalDown $child }
        } elseif ($node.children -and $node.children.Count -gt 0) {
            foreach ($child in $node.children) { Propagate-InternalDown $child }
        }
    }
    foreach ($root in $nestedRoots) { Propagate-InternalDown $root }
    function Propagate-Internal {
        param($nodeList)
        foreach ($n in $nodeList) {
            if ($n.children.Count -gt 0) { Propagate-Internal $n.children; if (($n.children | Where-Object { $_.isInternal }).Count -gt 0) { $n.isInternal = $true } }
        }
    }
    Propagate-Internal $nestedRoots
    $nestedRoots = @($nestedRoots | Where-Object { -not $_.isInternal }) + @($nestedRoots | Where-Object { $_.isInternal })
    $tree = @(@{ name = 'This Computer'; model = $env:COMPUTERNAME; vendor = 'Microsoft Windows'; isHub = $false; isComposite = $false; isInternal = $false; depth = 0; hops = 0; port = 0; children = $nestedRoots })
    function Traverse { param($node, $depth)
        $result = @($depth)
        foreach ($child in $node.children) { $result += Traverse $child ($depth + 1) }
        return $result
    }
    $depths = @()
    foreach ($root in $tree) { $depths += Traverse $root 0 }
    $maxHops = if ($depths.Count -gt 0) { ($depths | Measure-Object -Maximum).Maximum } else { 0 }
    $tierCount = ($depths | Select-Object -Unique).Count
    return @{ tree = $tree; maxHops = $maxHops; maxTiers = $tierCount; deviceCount = $deviceMap.Count }
}

# ============================================================================
# TREE FORMATTING
# ============================================================================
function Get-NodeLabel {
    param($node)
    $model = if ($node.model) { $node.model } else { if ($node.name) { $node.name } else { 'Unknown' } }
    $deviceInfo = if ($node.deviceInfo) { $node.deviceInfo } else { '' }
    $ifaceDesc = if ($node.interfaceDesc) { $node.interfaceDesc } else { '' }
    $ifaceNum = $node.interfaceNumber
    $isComposite = $node.isComposite -eq $true
    if ($isComposite) {
        $mi = if ($ifaceNum -ne $null) { "MI_$($ifaceNum.ToString('00'))" } else { '' }
        $suffix = if ($deviceInfo) { " ($deviceInfo)" } else { '' }
        if ($model -and $model -notmatch '(?i)USB-enhet|sammansatt|Composite') {
            $label = $model
            if ($mi) { $label += " $mi" }
            return "$label$suffix"
        }
        if ($ifaceDesc) {
            $ifaceTag = if ($ifaceDesc -match 'Keyboard') { 'HID Keyboard' } elseif ($ifaceDesc -match 'Mouse') { 'HID Mouse' } else { $ifaceDesc }
            return "$ifaceTag $mi$suffix".Trim()
        }
        return "$model$suffix"
    }
    if ($deviceInfo) { return "$model ($deviceInfo)" } else { return $model }
}

function Write-Tree {
    param($nodes, [string]$prefix = '', [switch]$showInternal, [bool]$parentIsInternal = $false)
    for ($i = 0; $i -lt $nodes.Count; $i++) {
        $node = $nodes[$i]
        $isLast = ($i -eq $nodes.Count - 1)
        $connector = if ($isLast) { $script:TeeEnd } else { $script:Tee }
        $badges = @()
        if ($node.isHub) { $badges += 'HUB' }
        if ($node.isDisplay) { $badges += 'DISPLAY' }
        if ($node.isInternal -and $showInternal -and -not $parentIsInternal) { $badges = @('INTERNAL') + $badges }
        $badgeStr = if ($badges.Count -gt 0) { '[' + ($badges -join '][') + '] ' } else { '' }
        $port = if ($node.port) { $node.port } else { 0 }
        $showPort = $port -gt 0 -and -not $node.isComposite
        $portStr = if ($showPort) { " [port $port]" } else { '' }
        $label = Get-NodeLabel $node
        Write-Host "$prefix$connector$badgeStr$label$portStr"
        if ($node.children -and $node.children.Count -gt 0) {
            if ($isLast) { $childPrefix = $prefix + '    ' } else { $childPrefix = $prefix + "$($script:TeeBar)   " }
            $collapseChildren = -not $node.isHub -and $node.isInternal
            if ($collapseChildren) {
                Write-Host "$childPrefix$($script:TeeEnd)[$($node.children.Count) interfaces]"
            } else {
                $newParentInt = $parentIsInternal -or ($node.isInternal -eq $true)
                Write-Tree $node.children $childPrefix -showInternal:$showInternal -parentIsInternal:$newParentInt
            }
        }
    }
}

# ============================================================================
# DISPLAY INFO
# ============================================================================
function Get-DisplayInfo {
    $displays = @()
    $monitors = try { Get-CimInstance -Namespace root\wmi -Class WmiMonitorID -ErrorAction Stop } catch { @() }
    if (-not $monitors -or $monitors.Count -eq 0) {
        $pnpMonitors = try { Get-PnpDevice -Class Monitor -ErrorAction SilentlyContinue | Where-Object Status -eq 'OK' } catch { @() }
        $i = 0
        foreach ($m in $pnpMonitors) {
            $name = if ($m.FriendlyName) { $m.FriendlyName } else { $m.Name }
            $displays += @{ index = $i; name = $name; resolution = 'Unknown'; isPrimary = ($i -eq 0); isInternal = $false }
            $i++
        }
        if ($displays.Count -gt 0) { return $displays }
        try {
            $cimMon = Get-CimInstance -ClassName Win32_DesktopMonitor -ErrorAction Stop | Where-Object { $_.Availability -eq 3 -or $_.ConfigManagerErrorCode -eq 0 }
            $i = 0
            foreach ($m in $cimMon) {
                $name = if ($m.MonitorManufacturerName) { $m.MonitorManufacturerName } else { if ($m.Name) { $m.Name } else { "Monitor $($i+1)" } }
                $res = if ($m.ScreenWidth -and $m.ScreenHeight -and $m.ScreenWidth -gt 0) { "$($m.ScreenWidth)x$($m.ScreenHeight)" } else { 'Unknown' }
                $displays += @{ index = $i; name = $name; resolution = $res; isPrimary = ($i -eq 0); isInternal = $false }
                $i++
            }
        } catch { }
        if ($displays.Count -gt 0) { return $displays }
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            $screens = [System.Windows.Forms.Screen]::AllScreens
            for ($i = 0; $i -lt $screens.Count; $i++) {
                $s = $screens[$i]; $res = "$($s.Bounds.Width)x$($s.Bounds.Height)"
                $displays += @{ index = $i; name = $s.DeviceName; resolution = $res; isPrimary = $s.Primary; isInternal = $false }
            }
        } catch { }
        return $displays
    }
    $basicParams = try { Get-CimInstance -Namespace root\wmi -Class WmiMonitorBasicDisplayParams -ErrorAction Stop } catch { @() }
    for ($i = 0; $i -lt $monitors.Count; $i++) {
        $m = $monitors[$i]
        $name = "Display $($i+1)"
        if ($m.UserFriendlyName -and $m.UserFriendlyName -ne 0) { $name = ($m.UserFriendlyName | ForEach-Object { [char]$_ }) -join '' }
        $width = 0; $height = 0; $widthMm = 0; $heightMm = 0
        $bp = $basicParams | Where-Object { $_.InstanceName -eq $m.InstanceName } | Select-Object -First 1
        if ($bp) { $width = $bp.HorizontalResolution; $height = $bp.VerticalResolution; $widthMm = $bp.MaxHorizontalImageSize; $heightMm = $bp.MaxVerticalImageSize }
        $isInternal = $false
        if ($widthMm -gt 0 -and $heightMm -gt 0) {
            $diagInches = [Math]::Sqrt($widthMm * $widthMm + $heightMm * $heightMm) / 25.4
            if ($diagInches -lt 18) { $isInternal = $true }
        }
        $displays += @{ index = $i; name = $name; width = $width; height = $height; resolution = if ($width -gt 0 -and $height -gt 0) { "${width}x${height}" } else { 'Unknown' }; isPrimary = ($i -eq 0); isInternal = $isInternal }
    }
    return $displays
}

# ============================================================================
# STABILITY ASSESSMENT
# ============================================================================
function Get-PlatformVerdicts {
    param($usbData, $currentHops, $currentTiers, $currentHubs)
    $notesByPlatform = @{}
    foreach ($n in $usbData.platformNotes) { $notesByPlatform[$n.platform] = $n }
    $verdicts = @()
    foreach ($pdef in $script:PlatformDefs) {
        $platformId = $pdef.id
        $maxH = if ($usbData.hopLimits.ContainsKey($platformId)) { $usbData.hopLimits[$platformId] } else { 4 }
        $maxT = if ($usbData.tierLimits.ContainsKey($platformId)) { $usbData.tierLimits[$platformId] } else { $maxH }
        $maxHub = if ($usbData.hubLimits.ContainsKey($platformId)) { $usbData.hubLimits[$platformId] } else { $maxH }
        $hStatus = Evaluate-Limit $currentHops $maxH
        $tStatus = Evaluate-Limit $currentTiers $maxT
        $hubStatus = Evaluate-Limit $currentHubs $maxHub
        $allStatus = @($hStatus, $tStatus, $hubStatus)
        $worst = ($allStatus | Sort-Object { $script:StatusRank[$_.status] } -Descending)[0]
        $status = $worst.status
        $warnings = @()
        if ($hStatus.warning) { $warnings += "Hops $($hStatus.warning)" }
        if ($tStatus.warning) { $warnings += "Tiers $($tStatus.warning)" }
        if ($hubStatus.warning) { $warnings += "Hubs $($hubStatus.warning)" }
        $colorMap = @{ 'STABLE' = 'green'; 'AT LIMIT' = 'orange'; 'UNSTABLE' = 'red' }
        $note = if ($notesByPlatform.ContainsKey($platformId)) { $notesByPlatform[$platformId] } else { $null }
        $desc = if ($note) { $note.system } else { $pdef.name }
        $verdicts += @{ id = $platformId; name = $pdef.name; arch = $pdef.arch; description = $desc; maxHops = $maxH; currentHops = $currentHops; maxTiers = $maxT; currentTiers = $currentTiers; maxHubs = $maxHub; currentHubs = $currentHubs; status = $status; color = $colorMap[$status]; isStable = ($status -ne 'UNSTABLE'); warning = if ($warnings.Count -gt 0) { $warnings -join '; ' } else { $null } }
    }
    return $verdicts
}

function Evaluate-Limit {
    param($current, $limit)
    if ($current -lt $limit) { return @{ status = 'STABLE'; warning = $null } }
    if ($current -eq $limit) { return @{ status = 'AT LIMIT'; warning = "at the limit ($current/$limit)" } }
    return @{ status = 'UNSTABLE'; warning = "exceeds limit ($current/$limit)" }
}

function Count-HubsInPath {
    param($tree)
    $state = @{ maxHubs = 0 }
    function Walk { param($node, $hubCount)
        $localMax = $hubCount
        if ($node.isHub) { $localMax = $hubCount + 1 }
        foreach ($child in $node.children) {
            $childMax = Walk $child $localMax
            if ($childMax -gt $localMax) { $localMax = $childMax }
        }
        if ($localMax -gt $state.maxHubs) { $state.maxHubs = $localMax }
        return $localMax
    }
    foreach ($root in $tree) { $null = Walk $root 0 }
    return $state.maxHubs
}

function Get-AllDevices {
    param($node)
    $result = @($node)
    foreach ($c in $node.children) { $result += Get-AllDevices $c }
    return $result
}

function Get-Stability {
    param($usbData, $tree, $hopsData)
    $currentHops = $hopsData.maxHops; $currentTiers = $hopsData.maxTiers
    $currentHubs = Count-HubsInPath $tree
    $overallVerdicts = Get-PlatformVerdicts $usbData $currentHops $currentTiers $currentHubs
    $ports = @()
    if ($tree -and $tree.Count -gt 0) {
        $root = $tree[0]; $rootChildIdx = 0
        foreach ($child in $root.children) {
            $rootChildIdx++
            $devicesUnder = Get-AllDevices $child
            if ($devicesUnder.Count -eq 0) { continue }
            $portState = @{ maxHopsFromPort = 0; depthLevels = @{}; hubsInChain = @{} }
            function Collect-Hops { param($node, $depthFromPort)
                if ($depthFromPort -gt $portState.maxHopsFromPort) { $portState.maxHopsFromPort = $depthFromPort }
                $portState.depthLevels[$depthFromPort] = $true
                if ($node.isHub) { $portState.hubsInChain["$($node.vid):$($node.pid)"] = $true }
                foreach ($c in $node.children) { Collect-Hops $c ($depthFromPort + 1) }
            }
            Collect-Hops $child 0
            $portHops = $portState.maxHopsFromPort
            $portTiers = $portState.depthLevels.Keys.Count
            $externalHubs = $portState.hubsInChain.Keys.Count
            $allUnder = Get-AllDevices $child
            $childrenOnly = $allUnder | Where-Object { $_ -ne $child }
            $endpointDevices = if ($childrenOnly.Count -gt 0) { $childrenOnly } else { @($child) }
            $endpointNames = @()
            foreach ($ed in $endpointDevices) { $endpointNames += if ($ed.model) { $ed.model } else { if ($ed.name) { $ed.name } else { 'Unknown' } } }
            $label = "$($child.model) ($($child.deviceInfo))"
            if (-not $child.deviceInfo) { $label = $child.model }
            $portVerdicts = Get-PlatformVerdicts $usbData $portHops $portTiers $externalHubs
            $ports += @{ id = $rootChildIdx; label = $label; maxHops = $portHops; maxTiers = $portTiers; externalHubs = $externalHubs; devices = $endpointNames; verdicts = $portVerdicts; nodeRef = $child }
        }
    }
    $totalEndpoints = ($ports | ForEach-Object { $_.devices.Count } | Measure-Object -Sum).Sum
    $overallWorst = 'STABLE'; $worstRank = 0
    foreach ($v in $overallVerdicts) { $r = $script:StatusRank[$v.status]; if ($r -gt $worstRank) { $worstRank = $r; $overallWorst = $v.status } }
    return @{ maxHops = $currentHops; maxTiers = $currentTiers; maxHubs = $currentHubs; totalEndpoints = $totalEndpoints; verdicts = $overallVerdicts; ports = $ports; overallWorst = $overallWorst }
}

# ============================================================================
# VERDICT FORMATTING
# ============================================================================
function Write-Verdict {
    param($v)
    $statusChar = if ($v.color -eq 'green') { '+' } elseif ($v.color -eq 'orange') { '~' } else { '!' }
    $hubsStr = "hubs $($v.currentHubs)/$($v.maxHubs)  "
    $desc = if ($v.description) { $v.description } else { $v.name }
    Write-Host ("    {0} {1,-22} {2,-9} hops {3}/{4}  tiers {5}/{6}  {7}" -f $statusChar, $desc, $v.status, $v.currentHops, $v.maxHops, $v.currentTiers, $v.maxTiers, $hubsStr)
}

# ============================================================================
# OUTPUT HELPERS
# ============================================================================
function Write-Header {
    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor Cyan
    Write-Host "KLANGCHE PROAV SHOKO - USB DETECTIVE" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
    Write-Host ""
}

function Write-Separator {
    Write-Host ("-" * 70)
}

# ============================================================================
# REPORT DISPLAY
# ============================================================================
function Show-Report {
    param($platformInfo, $usbData, $usbResult, $displays, $stability)
    Write-Header
    Write-Host "Platform: $($platformInfo.versionString)" -ForegroundColor Gray
    Write-Host "Architecture: $($platformInfo.architecture)" -ForegroundColor Gray
    if ($platformInfo.isAppleSilicon) { Write-Host "Apple Silicon: Yes" -ForegroundColor Yellow }
    Write-Separator
    $tree = $usbResult.tree; Write-Host ""
    $overall = $stability.overallWorst; $mh = $stability.maxHops; $mt = $stability.maxTiers; $mhub = $stability.maxHubs
    $total = $stability.totalEndpoints; $ep = if ($total -eq 1) { "endpoint" } else { "endpoints" }
    Write-Host "Overall: $overall ($total $ep, hops=$mh, tiers=$mt, hubs=$mhub)" -ForegroundColor White
    Write-Host ""
    $origChildren = @($tree[0].children)
    Write-Host "  Full USB and Display Tree"
    if ($displays.Count -gt 0) {
        foreach ($d in $displays) {
            $prim = if ($d.isPrimary) { " (Primary)" } else { "" }
            $intDisp = $d.isInternal
            $tree[0].children += @{ model = "$($d.resolution)  $($d.name)$prim"; name = $d.name; children = @(); hops = 1; isHub = $false; isInternal = $intDisp; isDisplay = $true; port = 0 }
        }
    }
    Write-Tree $tree "  " -showInternal
    Write-Host ""
    Write-Host "  Overall rating"
    foreach ($v in $stability.verdicts) { Write-Verdict $v }
    Write-Host ""
    Write-Host ("=" * 31) -NoNewline; Write-Host "PER PORT" -NoNewline; Write-Host ("=" * 31)
    Write-Host ""
    $sepLine = "  " + ("- " * 35)
    function Write-PortChild { param($child, $idx)
        $portInfo = $stability.ports | Where-Object { $_.id -eq ($idx + 1) } | Select-Object -First 1
        $label = if ($portInfo) { $portInfo.label } else { if ($child.model) { $child.model } else { 'Port' } }
        $dc = if ($portInfo) { $portInfo.devices.Count } else { 0 }
        $ph = if ($portInfo) { $portInfo.maxHops } else { 0 }
        $pt = if ($portInfo) { $portInfo.maxTiers } else { 0 }
        $pHub = if ($portInfo) { $portInfo.externalHubs } else { 0 }
        $isInt = $child.isInternal -eq $true
        $intPre = if ($isInt) { "[INTERNAL] " } else { "" }
        $ep = if ($dc -eq 1) { "endpoint" } else { "endpoints" }
        Write-Host "  $intPre$label ($dc $ep, hops=$ph, tiers=$pt, hubs=$pHub)"
        if ($child.children -and $child.children.Count -gt 0) {
            Write-Tree $child.children "    "
        }
        if ($isInt) { Write-Host "    (internal)" }
        elseif ($portInfo) { foreach ($v in $portInfo.verdicts) { Write-Verdict $v } }
    }
    function Write-SectionGroup { param([string]$header, [scriptblock]$filter)
        Write-Host $header -NoNewline; Write-Host ("-" * 31)
        $first = $true
        for ($idx = 0; $idx -lt $origChildren.Count; $idx++) {
            $child = $origChildren[$idx]
            if ($child.isDisplay) { continue }
            if (-not (&$filter $child)) { continue }
            if (-not $first) { Write-Host "" }
            Write-PortChild $child $idx; Write-Host ""; Write-Host $sepLine
            $first = $false
        }
    }
    Write-SectionGroup "EXTERNAL" { param($c) return ($c.isInternal -ne $true) }
    Write-SectionGroup "INTERNAL" { param($c) return ($c.isInternal -eq $true) }
    return $origChildren
}

# ============================================================================
# LIVE MONITORING
# ============================================================================
function Start-Monitoring {
    $logs = @(); $prevDevices = @{}; $deviceEvents = @{}
    Write-Host ""; Write-Host ("=" * 70)
    Write-Host "LIVE MONITORING - Press Enter to stop and generate report"
    Write-Host ("=" * 70); Write-Host "Logs appear in real-time below:"; Write-Host ""
    while (-not $Host.UI.RawUI.KeyAvailable) {
        $current = @{}
        try { $devices = Get-PnpDevice -Class USB -PresentOnly -ErrorAction SilentlyContinue | Where-Object Status -eq 'OK'; foreach ($d in $devices) { $current[$d.InstanceId] = @{ name = if ($d.FriendlyName) { $d.FriendlyName } else { $d.Name }; status = $d.Status } } } catch { }
        $now = Get-Date -Format 'HH:mm:ss'
        $newIds = $current.Keys | Where-Object { -not $prevDevices.ContainsKey($_) }
        $removedIds = $prevDevices.Keys | Where-Object { -not $current.ContainsKey($_) }
        foreach ($id in $newIds) {
            $info = $current[$id]; Write-Host "[$now] CONNECTED: $($info.name)"
            if (-not $deviceEvents.ContainsKey($id)) { $deviceEvents[$id] = @{ connects = 0; disconnects = 0; lastAction = $null } }
            if ($deviceEvents[$id].lastAction -eq 'DISCONNECTED') { Write-Host "  [!] RE-CONNECT (possible instability): $($info.name)" -ForegroundColor Yellow }
            $deviceEvents[$id].connects++; $deviceEvents[$id].lastAction = 'CONNECTED'
        }
        foreach ($id in $removedIds) {
            $info = $prevDevices[$id]; Write-Host "[$now] DISCONNECTED: $($info.name)"
            if (-not $deviceEvents.ContainsKey($id)) { $deviceEvents[$id] = @{ connects = 0; disconnects = 0; lastAction = $null } }
            $deviceEvents[$id].disconnects++; $deviceEvents[$id].lastAction = 'DISCONNECTED'
        }
        $prevDevices = $current; Start-Sleep -Milliseconds 500
    }
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Write-Host ""; Write-Host ("=" * 70); Write-Host "MONITORING STOPPED - Logs captured"; Write-Host ("=" * 70)
    $unstableList = @()
    foreach ($id in $deviceEvents.Keys) {
        $ev = $deviceEvents[$id]
        if ($ev.connects -gt 0 -and $ev.disconnects -gt 0) {
            $name = if ($prevDevices.ContainsKey($id)) { $prevDevices[$id].name } else { $id }
            $unstableList += $name
        }
    }
    if ($unstableList.Count -gt 0) { Write-Host "UNSTABLE DEVICES DETECTED (reconnected during monitoring):" -ForegroundColor Yellow; foreach ($dev in $unstableList | Sort-Object) { Write-Host "  ! $dev" -ForegroundColor Red } }
    Write-Host ("=" * 70); Write-Host ""
    return @{ logs = $logs; unstableDevices = $unstableList }
}

# ============================================================================
# HTML REPORT
# ============================================================================
function Save-HtmlReport {
    param($platformInfo, $usbData, $usbResult, $displays, $stability, $monitoringResult)
    $dateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $outHtml = "$env:TEMP\proav-shoko-report-$dateStamp.html"
    $overall = $stability.overallWorst; $mh = $stability.maxHops; $mt = $stability.maxTiers; $mhub = $stability.maxHubs
    $total = $stability.totalEndpoints; $ep = if ($total -eq 1) { "endpoint" } else { "endpoints" }
    $verdictLines = @("Overall: $overall ($total $ep, hops=$mh, tiers=$mt, hubs=$mhub)")
    foreach ($v in $stability.verdicts) {
        $sc = if ($v.color -eq 'green') { '+' } elseif ($v.color -eq 'orange') { '~' } else { '!' }
        $hs = "hubs $($v.currentHubs)/$($v.maxHubs)  "
        $vd = if ($v.description) { $v.description } else { $v.name }
        $verdictLines += "    $sc $($vd.PadRight(22)) $($v.status.PadRight(9)) hops $($v.currentHops)/$($v.maxHops)  tiers $($v.currentTiers)/$($v.maxTiers)  $hs"
    }
    $portLines = @("PER PORT===============================")
    foreach ($p in $stability.ports) {
        $ep2 = if ($p.devices.Count -eq 1) { "endpoint" } else { "endpoints" }
        $portLines += "  $($p.label) ($($p.devices.Count) $ep2, hops=$($p.maxHops), tiers=$($p.maxTiers), hubs=$($p.externalHubs))"
        foreach ($v in $p.verdicts) {
            $sc = if ($v.color -eq 'green') { '+' } elseif ($v.color -eq 'orange') { '~' } else { '!' }
            $hs = "hubs $($v.currentHubs)/$($v.maxHubs)  "
            $vd = if ($v.description) { $v.description } else { $v.name }
            $portLines += "    $sc $($vd.PadRight(22)) $($v.status.PadRight(9)) hops $($v.currentHops)/$($v.maxHops)  tiers $($v.currentTiers)/$($v.maxTiers)  $hs"
        }
    }
    $dispLines = @("Connected Displays:")
    foreach ($d in $displays) {
        $prim = if ($d.isPrimary) { " (Primary)" } else { "" }
        $int = if ($d.isInternal) { "[INTERNAL] " } else { "" }
        $dispLines += "  ${int}[DISPLAY] $($d.resolution)  $($d.name)$prim"
    }
    $mtText = ""
    if ($monitoringResult -and $monitoringResult.unstableDevices.Count -gt 0) {
        $mtText = "`nUNSTABLE DEVICES:`n" + ($monitoringResult.unstableDevices | ForEach-Object { "  ! $_" }) -join "`n"
    }
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>ProAV Shoko - USB Analysis</title>
<style>
  body { font-family: 'Consolas', 'Courier New', monospace; background: #0c0c10; color: #d4d4d4; padding: 24px; font-size: 13px; }
  .header { color: #569CD6; font-size: 1.3rem; font-weight: 600; border-bottom: 1px solid #333; padding-bottom: 10px; margin-bottom: 20px; }
  .section { background: #0a0a12; border: 1px solid #1a1a2e; border-radius: 3px; padding: 16px; margin: 12px 0; }
  .section-title { color: #569CD6; font-size: 1rem; font-weight: 600; margin: 16px 0 8px 0; padding-bottom: 6px; border-bottom: 1px solid #1a1a2e; }
  .pre { white-space: pre; font-family: inherit; font-size: 0.75rem; line-height: 1.4; color: #d4d4d4; }
  .tag { display: inline-block; background: #0a0a12; border: 1px solid #333; padding: 2px 10px; border-radius: 3px; font-size: 0.7rem; color: #808080; margin-right: 4px; }
  .stat-box { display: inline-block; background: #0a0a12; border: 1px solid #1a1a2e; padding: 8px 16px; margin: 4px; text-align: center; border-radius: 3px; }
  .stat-value { font-size: 1.1rem; font-weight: 700; color: #569CD6; }
  .stat-label { color: #808080; font-size: 0.6rem; text-transform: uppercase; }
  .footer { margin-top: 24px; padding-top: 12px; border-top: 1px solid #333; color: #555; font-size: 0.6rem; text-align: center; }
</style>
</head>
<body>
<div class="header">ProAV Shoko -- USB Analysis</div>
<div class="tag">$($platformInfo.versionString)</div>
<div class="tag">$($platformInfo.architecture)</div>
<br><br>
<div>
  <div class="stat-box"><div class="stat-value">$mh</div><div class="stat-label">Max Hops</div></div>
  <div class="stat-box"><div class="stat-value">$mt</div><div class="stat-label">Tiers</div></div>
  <div class="stat-box"><div class="stat-value">$mhub</div><div class="stat-label">Hubs</div></div>
  <div class="stat-box"><div class="stat-value">$($displays.Count)</div><div class="stat-label">Displays</div></div>
</div>
<div class="section">
  <div class="section-title">OVERALL RATING</div>
  <pre class="pre">$($verdictLines -join "`n")</pre>
</div>
<div class="section">
  <div class="section-title">PER PORT</div>
  <pre class="pre">$($portLines -join "`n")</pre>
</div>
<div class="section">
  <div class="section-title">CONNECTED DISPLAYS</div>
  <pre class="pre">$($dispLines -join "`n")</pre>
</div>
$mtText
<div class="footer">ProAV Shoko v1.0.0 -- $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>
</body>
</html>
"@
    try {
        $htmlContent | Out-File $outHtml -Encoding UTF8
        Write-Host "  HTML Report: $outHtml" -ForegroundColor Green
        Start-Process $outHtml
    } catch {
        Write-Host "  Could not open HTML report: $_" -ForegroundColor Red
        Write-Host "  Saved to: $outHtml" -ForegroundColor Gray
    }
}

# ============================================================================
# PROMPTS
# ============================================================================
function Prompt-ReportFormat {
    while ($true) {
        $choice = Read-Host "`nReport: [Enter]HTML / [P]DF / [N]o report"
        $choice = $choice.Trim().ToUpper()
        if ($choice -in @('', 'H', 'HTML')) { return 'html' }
        if ($choice -in @('P', 'PDF')) { return 'pdf' }
        if ($choice -in @('N', 'NO', 'NONE')) { return 'none' }
        Write-Host "  Please press Enter for HTML, P for PDF, or N for no report"
    }
}

function Prompt-Monitor {
    while ($true) {
        $choice = Read-Host "`nStart live USB monitoring? [Y/n]"
        $choice = $choice.Trim().ToLower()
        if ($choice -in @('', 'y', 'yes')) { return $true }
        if ($choice -in @('n', 'no')) { return $false }
        Write-Host "  Please enter Y or N"
    }
}

# ============================================================================
# MAIN
# ============================================================================
function Main {
    $isAdmin = $false
    try { $identity = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent(); $isAdmin = $identity.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch { }
    if (-not $isAdmin -and -not $env:SHOKO_SKIP_ADMIN_CHECK) {
        Write-Host "Administrator mode provides more USB data." -ForegroundColor Yellow
        $elevate = Read-Host "Run with administrator privileges? [Y/n]"
        if ($elevate -eq '' -or $elevate -match '^[Yy]') {
            if (Test-Path -LiteralPath $PSCommandPath -ErrorAction SilentlyContinue) {
                Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
            }
            exit
        }
        Write-Host "Continuing in basic mode (limited data)" -ForegroundColor Gray
    }

    $csvUrl = "https://raw.githubusercontent.com/klangche/klangche-proav-shoko/refs/heads/main/src/assets/usb_data.csv"
    $usbData = Load-UsbData $csvUrl
    $platformInfo = Get-PlatformInfo
    Write-Progress -Activity "ProAV Shoko" -Status "Scanning USB devices..." -PercentComplete -1
    $usbResult = Get-UsbTree
    Write-Progress -Activity "ProAV Shoko" -Status "Scanning displays..." -PercentComplete -1
    $displays = Get-DisplayInfo
    Write-Progress -Activity "ProAV Shoko" -Completed
    $stability = Get-Stability $usbData $usbResult.tree @{ maxHops = $usbResult.maxHops; maxTiers = $usbResult.maxTiers }
    $origChildren = Show-Report $platformInfo $usbData $usbResult $displays $stability
    $monitoringResult = $null
    if (Prompt-Monitor) { $monitoringResult = Start-Monitoring }
    $formatType = Prompt-ReportFormat
    if ($formatType -ne 'none') {
        Write-Progress -Activity "ProAV Shoko" -Status "Generating report..." -PercentComplete -1
        Save-HtmlReport $platformInfo $usbData $usbResult $displays $stability $monitoringResult
        Write-Progress -Activity "ProAV Shoko" -Completed
    }
    Write-Separator
    Write-Host "Done!" -ForegroundColor Green
    Write-Host ("=" * 70) -ForegroundColor Cyan
}

Main
