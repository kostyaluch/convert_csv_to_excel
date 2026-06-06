@echo off
:: Batch script to install context menu for CSV to Excel converter
:: Run this script as Administrator

echo.
echo ===================================================
echo   CSV to Excel Converter - Context Menu Installer
echo ===================================================
echo.

:: Get the directory where this script is located
set SCRIPT_DIR=%~dp0
:: Remove trailing backslash
set SCRIPT_DIR=%SCRIPT_DIR:~0,-1%

set EXE_PATH=%SCRIPT_DIR%\convert_csv_to_excel_v3.exe

:: Check if exe exists
if not exist "%EXE_PATH%" (
    echo [ERROR] convert_csv_to_excel_v3.exe not found in %SCRIPT_DIR%
    echo Please make sure this script is in the same directory as the executable.
    echo.
    pause
    exit /b 1
)

echo Installing context menu...
echo Executable path: %EXE_PATH%
echo.

:: Create registry entries
reg add "HKCR\.csv" /ve /d "csvfile" /f >nul 2>&1
reg add "HKCR\csvfile" /ve /d "CSV File" /f >nul 2>&1
reg add "HKCR\csvfile\shell\ConvertToExcel" /ve /d "Конвертувати у Excel" /f >nul 2>&1
reg add "HKCR\csvfile\shell\ConvertToExcel" /v "Icon" /d "\"%EXE_PATH%\",0" /f >nul 2>&1
reg add "HKCR\csvfile\shell\ConvertToExcel\command" /ve /d "\"%EXE_PATH%\" \"%%1\"" /f >nul 2>&1

if %ERRORLEVEL% equ 0 (
    echo [SUCCESS] Context menu installed successfully!
    echo.
    echo You can now right-click on any CSV file and select
    echo "Конвертувати у Excel" to convert it.
) else (
    echo [ERROR] Failed to install context menu.
    echo Please make sure you run this script as Administrator.
)

echo.
pause
