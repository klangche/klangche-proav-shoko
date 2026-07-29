# Save the gui.py file to a temporary location and search for it more carefully
$gui_path = 'C:\Users\linus\Documents\GitHub\klangche-proav-shoko\src\gui.py'

if (Test-Path $gui_path) {
    $content = Get-Content -Path $gui_path -Raw
    if ($content -match 'text_color') {
        Write-Host 'ERROR: Found text_color in gui.py - need manual fix'
        # Show lines with text_color
        $lines = $content -split '\\n'
        for ($i = 0; $i -lt $lines.Length; $i++) {
            if ($lines[$i] -match 'text_color') {
                Write-Host "Line $($i + 1): $($lines[$i].Trim())"
            }
        }
        exit 1
    } else {
        Write-Host 'SUCCESS: No text_color in gui.py'
    }
} else {
    Write-Host 'ERROR: gui.py not found at $gui_path'
}