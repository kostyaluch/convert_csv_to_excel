# PowerShell script to uninstall context menu for CSV to Excel converter
# Run this script as Administrator

Write-Host "Uninstalling context menu for CSV to Excel converter..." -ForegroundColor Yellow

# Remove context menu entry
$shellPath = "HKCR:\csvfile\shell\ConvertToExcel"
if (Test-Path $shellPath) {
    Remove-Item -Path $shellPath -Recurse -Force
    Write-Host "Context menu removed successfully!" -ForegroundColor Green
} else {
    Write-Host "Context menu entry not found." -ForegroundColor Yellow
}

Pause
