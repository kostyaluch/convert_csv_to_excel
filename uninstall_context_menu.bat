@echo off
:: Batch script to uninstall context menu for CSV to Excel converter
:: Run this script as Administrator

echo.
echo =====================================================
echo   CSV to Excel Converter - Context Menu Uninstaller
echo =====================================================
echo.

echo Removing context menu...

:: Remove registry entries
reg delete "HKCR\csvfile\shell\ConvertToExcel" /f >nul 2>&1

if %ERRORLEVEL% equ 0 (
    echo [SUCCESS] Context menu removed successfully!
) else (
    echo [INFO] Context menu entry not found or already removed.
)

echo.
pause
