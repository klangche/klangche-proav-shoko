# proav-shoko_powershell.ps1 - Shōko Main Logic (USB + Display Diagnostics)

$RepoOwner = "klangche"
$RepoName   = "klangche-proav-shoko"
$Branch     = "main"
$BaseUrl    = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch"

# Load config (now from current repo)
try {
    $Config = Invoke-RestMethod -Uri "$BaseUrl/proav-shoko.json" -UseBasicParsing
    Write-Host "Config loaded (v$($Config.version))" -ForegroundColor Green
} catch {
    Write-Host "Config load failed - using defaults" -ForegroundColor Yellow
    $Config = [PSCustomObject]@{ version = "fallback" }
}

# Color helper using config if available
function Get-Color { param($Name)
    $map = @{ cyan = "Cyan"; magenta = "Magenta"; yellow = "Yellow"; green = "Green"; gray = "Gray" }
    return $map[$Name] ?? "White"
}

Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "Shōko - USB + Display Diagnostic Tool v$($Config.version)" -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Running in basic mode (limited features)" -ForegroundColor Yellow
}

# ────────────────────────────────────────────────────────────────────────────────
# USB TREE SECTION (your original logic – update any old URLs here if present)
# ────────────────────────────────────────────────────────────────────────────────
# ... Paste / keep your existing USB enumeration, tree building, stability calc, output code here ...
# Make sure any Invoke-RestMethod or download lines use $BaseUrl above instead of old repo paths.

# Example stability summary print (adapt to your code)
Write-Host "USB Stability Verdict" -ForegroundColor Magenta
# ... your verdict output ...

# ────────────────────────────────────────────────────────────────────────────────
# DISPLAY TREE & ANALYTICS (integrated)
# ────────────────────────────────────────────────────────────────────────────────

$showDisplay = Read-Host "`nShow display information and analytics? (y/n)"
if ($showDisplay -match '^[Yy]') {
    function Get-DisplayTree {
        # ... Paste the full Get-DisplayTree function from our previous response here ...
        # (the one with Decode-Connection, Detect-Transport, event logs, health hints, etc.)
        # It uses no external downloads, so no URL fixes needed inside.
    }
    
    Get-DisplayTree
}

# ────────────────────────────────────────────────────────────────────────────────
# DEEP ANALYTICS / MONITORING LOOP (your original – keep as is)
# ────────────────────────────────────────────────────────────────────────────────

# ... your existing deep USB analytics prompt + infinite loop until Ctrl+C ...

Write-Host "`nShōko complete. Press Enter to exit." -ForegroundColor Green
Read-Host
