@echo off
setlocal EnableExtensions
cd /d "%~dp0"

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "tools\repair_private_runtime.ps1"
if errorlevel 1 (
    echo.
    echo REPAIR FAILED. Use the complete portable release folder instead.
    pause
    exit /b 1
)

echo.
echo PRIVATE RUNTIME REPAIRED.
pause
