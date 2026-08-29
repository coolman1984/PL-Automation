@echo off
setlocal EnableExtensions
cd /d "%~dp0"

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "tools\prepare_offline_packages.ps1"
if errorlevel 1 (
    echo.
    echo PREPARATION FAILED. This step requires an internet-connected Windows build PC.
    pause
    exit /b 1
)

echo.
echo OFFLINE PACKAGES ARE READY.
pause
