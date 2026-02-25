# proav-shoko_powershell.ps1

# Load config (fixed URL)
try {
    $Config = Invoke-RestMethod "https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.json"
} catch {
    Write-Host "Config failed" -ForegroundColor Yellow
    $Config = [PSCustomObject]@{ version = "local" }
}

# ... your color helper function ...

Write-Host "Shoko v$($Config.version)" -ForegroundColor Cyan

$isAdmin = ... # your admin check

# ────────────────────────────────────────────────────────────────
# 1. USB TREE + STABILITY – always shown here, immediately
# ────────────────────────────────────────────────────────────────

# PASTE YOUR ORIGINAL USB CODE HERE (enumeration, map, Print-Tree recursive function, maxHops, numTiers, stability table, verdict, colors from config, etc.)

# Example skeleton (replace with your real code):
$treeOutput = "your tree string"
$maxHops = 3
# ... calculate score, status, etc. ...

Write-Host $treeOutput
Write-Host "Max hops: $maxHops ..."
Write-Host "Stability verdict: $verdict" -ForegroundColor $verdictColor

# Save txt & html reports (your original code)
$outHtml = ... 
$htmlContent = "<pre>$treeOutput ... verdict ...</pre>"
$htmlContent | Out-File $outHtml

# Ask browser
$open = Read-Host "Open HTML report in browser? (y/n)"
if ($open -match '^[Yy]') { Start-Process $outHtml }

# ────────────────────────────────────────────────────────────────
# 2. DISPLAY TREE & ANALYTICS – optional, but in same style
# ────────────────────────────────────────────────────────────────

$showDisplay = Read-Host "Show display tree & analytics? (y/n)"
if ($showDisplay -match '^[Yy]') {
    # The Get-DisplayTree function (PS5.1 safe version from earlier)
    function Get-DisplayTree {
        # ... full function with Decode-Connection, Detect-Transport, event query, tree printing ...
        # Make sure .Trim() is correct: (($chars) -join '').Trim()
        # Use if/else instead of ? :
        $adapterName = if ($adapter) { $adapter.Name } else { "" }
    }

    Get-DisplayTree

    # Optional second browser ask if you want separate display report
    $open2 = Read-Host "Open updated report? (y/n)"
    if ($open2 -match '^[Yy]') { Start-Process $outHtml }
}

# ────────────────────────────────────────────────────────────────
# 3. DEEP ANALYTICS / MONITORING – loop until Ctrl+C
# ────────────────────────────────────────────────────────────────

$wantDeep = Read-Host "Run deep analytics / monitoring? (y/n)"
if ($wantDeep -match '^[Yy]' -and $isAdmin) {
    Write-Host "Deep monitoring started. Press Ctrl+C to stop." -ForegroundColor Green

    try {
        while ($true) {
            # Your original deep analytics code (USB event counters, stability re-check, etc.)
            # Optionally add display event check here too
            Write-Host "$(Get-Date -Format HH:mm:ss) - Monitoring active..." -ForegroundColor Green
            Start-Sleep -Seconds 5
        }
    }
    catch {
        Write-Host "Monitoring stopped (Ctrl+C detected)." -ForegroundColor Yellow
    }
}

Write-Host "Shoko finished." -ForegroundColor Green
Read-Host "Press Enter to close"
