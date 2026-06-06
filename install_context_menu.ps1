# PowerShell script to install context menu for CSV to Excel converter
# Run this script as Administrator

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$exePath = Join-Path $scriptPath "convert_csv_to_excel_v3.exe"

# Check if exe exists
if (-Not (Test-Path $exePath)) {
    Write-Host "Error: convert_csv_to_excel_v3.exe not found in $scriptPath" -ForegroundColor Red
    Write-Host "Please make sure the script is in the same directory as the executable." -ForegroundColor Yellow
    Pause
    exit
}

Write-Host "Installing context menu for CSV to Excel converter..." -ForegroundColor Green
Write-Host "Executable path: $exePath" -ForegroundColor Cyan

# Ensure .csv file association exists
$csvFileType = "csvfile"
if (-Not (Test-Path "HKCR:\.$csv")) {
    New-Item -Path "HKCR:\.csv" -Force | Out-Null
}
Set-ItemProperty -Path "HKCR:\.csv" -Name "(Default)" -Value $csvFileType -Force

# Create csvfile type if it doesn't exist
if (-Not (Test-Path "HKCR:\$csvFileType")) {
    New-Item -Path "HKCR:\$csvFileType" -Force | Out-Null
}
Set-ItemProperty -Path "HKCR:\$csvFileType" -Name "(Default)" -Value "CSV File" -Force

# Create context menu entry
$shellPath = "HKCR:\$csvFileType\shell\ConvertToExcel"
New-Item -Path $shellPath -Force | Out-Null
Set-ItemProperty -Path $shellPath -Name "(Default)" -Value "Конвертувати у Excel" -Force
Set-ItemProperty -Path $shellPath -Name "Icon" -Value "$exePath,0" -Force

# Create command entry
$commandPath = "$shellPath\command"
New-Item -Path $commandPath -Force | Out-Null
Set-ItemProperty -Path $commandPath -Name "(Default)" -Value "`"$exePath`" `"%1`"" -Force

Write-Host "Context menu installed successfully!" -ForegroundColor Green
Write-Host "You can now right-click on any CSV file and select 'Конвертувати у Excel'" -ForegroundColor Cyan
Pause
