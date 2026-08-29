@echo off
setlocal EnableExtensions
cd /d "%~dp0"

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "tools\build_portable.ps1"
if errorlevel 1 (
    echo.
    echo BUILD FAILED. Read the message above.
    pause
    exit /b 1
)

echo.
echo BUILD COMPLETE. Open the release folder for the portable ZIP.
start "" explorer "%~dp0release"
pause
