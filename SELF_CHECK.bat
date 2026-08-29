@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if exist "P_L_Excel_Agent.exe" (
    "P_L_Excel_Agent.exe" --self-check
    goto :done
)
if exist ".runtime_env\Scripts\python.exe" (
    ".runtime_env\Scripts\python.exe" app.py --self-check
    goto :done
)
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py --self-check
    goto :done
)

echo ERROR: The portable executable or a private runtime was not found.
echo Keep the complete release folder together.

:done
echo.
pause
