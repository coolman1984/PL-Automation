@echo off
setlocal EnableExtensions
title Excel Agent - Safe Preparation
cd /d "%~dp0"

set "WORKBOOK=%~1"
if "%WORKBOOK%"=="" set /p WORKBOOK=Drag an Excel workbook here or type its full path, then press ENTER: 
set "WORKBOOK=%WORKBOOK:"=%"
if not exist "%WORKBOOK%" (
    echo ERROR: The workbook was not found.
    pause
    exit /b 2
)

if exist "P_L_Excel_Agent.exe" (
    "P_L_Excel_Agent.exe" --file "%WORKBOOK%" --prepare --snapshot-mode auto
    goto :result
)
if exist ".runtime_env\Scripts\python.exe" (
    ".runtime_env\Scripts\python.exe" app.py --file "%WORKBOOK%" --prepare --snapshot-mode auto
    goto :result
)
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py --file "%WORKBOOK%" --prepare --snapshot-mode auto
    goto :result
)

echo ERROR: The portable application or private runtime is missing.
echo Keep the complete release folder together.
pause
exit /b 2

:result
if errorlevel 1 (
    echo.
    echo PREPARATION STOPPED SAFELY. Read the reason above.
    pause
    exit /b 1
)
echo.
echo BACKUP AND INVENTORY SNAPSHOT COMPLETED.
start "" explorer "%~dp0agent_artifacts"
pause
