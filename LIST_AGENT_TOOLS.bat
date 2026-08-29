@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if exist "P_L_Excel_Agent.exe" (
    "P_L_Excel_Agent.exe" --list-tools
) else if exist ".runtime_env\Scripts\python.exe" (
    ".runtime_env\Scripts\python.exe" app.py --list-tools
) else if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py --list-tools
) else (
    echo ERROR: The portable application or private runtime is missing.
)
echo.
pause
