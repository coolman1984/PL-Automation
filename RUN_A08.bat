@echo off
setlocal EnableExtensions
title P^&L August Actual Update

rem ---------------------------------------------------------------------------
rem Launcher for the P&L A08 automation.
rem Usage: drag a .xlsb workbook onto this file, or double-click and paste a path.
rem ---------------------------------------------------------------------------

cd /d "%~dp0"

set "WORKBOOK="
if not "%~1"=="" set "WORKBOOK=%~1"
if "%WORKBOOK%"=="" set /p WORKBOOK=Drag the protected workbook here or type its full path, then press ENTER: 

rem Strip surrounding quotes safely.
set "WORKBOOK=%WORKBOOK:"=%"

echo.
echo Source workbook: %WORKBOOK%
echo.

if not exist "%WORKBOOK%" (
    echo ERROR: The workbook was not found.
    goto :fail
)

rem The packaged executable is the production path. It contains its own Python
rem runtime and packages and never depends on the PC's Python installation.
set "APP_EXE="
if exist "P_L_Excel_Agent.exe" set "APP_EXE=%~dp0P_L_Excel_Agent.exe"

rem Source/developer fallback only. It always prefers a private local runtime.
set "PY="
if not defined APP_EXE if exist ".runtime_env\Scripts\python.exe" (
    set "PY=.runtime_env\Scripts\python.exe"
)
if not defined APP_EXE if not defined PY if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
)

if not defined APP_EXE if not defined PY (
    for /f "delims=" %%P in ('where python.exe 2^>nul') do call :try_python "%%P"
)
if not defined APP_EXE if not defined PY (
    for /f "delims=" %%P in ('where python3.exe 2^>nul') do call :try_python "%%P"
)
if not defined APP_EXE if not defined PY (
    for %%V in (314 313 312 311) do (
        call :try_python "%LocalAppData%\Programs\Python\Python%%V\python.exe"
        call :try_python "%ProgramFiles%\Python%%V\python.exe"
    )
)
if not defined APP_EXE if not defined PY (
    echo ERROR: The portable executable and private runtime are missing.
    echo Use the complete release folder, or run REPAIR_OFFLINE.bat.
    goto :fail
)

if defined APP_EXE echo Portable application selected: %APP_EXE%
if defined PY echo Private/developer Python selected: %PY%

call :run_app --self-check
if errorlevel 1 (
    echo.
    echo SELF-CHECK FAILED. Nothing was changed.
    goto :fail
)

echo STEP 1 of 3: Running quick protection and format check...
call :run_app --file "%WORKBOOK%" --year 2026 --month 8 --probe-only
if errorlevel 1 (
    echo.
    echo QUICK CHECK FAILED. The file was not recognized safely.
    goto :fail
)

echo.
echo STEP 2 of 3: Running read-only dry-run with the selected safe engine...
call :run_app --file "%WORKBOOK%" --year 2026 --month 8 --dry-run
if errorlevel 1 (
    echo.
    echo DRY-RUN DID NOT PASS. Nothing was changed. Read the reasons above.
    goto :fail
)

echo.
echo The dry-run passed. The original file is NEVER modified;
echo an updated copy will be written to the output folder.
choice /c YN /m "STEP 3 of 3: Proceed with creating the updated copy"
if errorlevel 2 goto :cancelled

call :run_app --file "%WORKBOOK%" --year 2026 --month 8 --execute
if errorlevel 1 (
    echo.
    echo EXECUTION FAILED SAFELY. The original file is unchanged.
    echo Details were saved under work\ and failed_runs\ next to this tool.
    goto :fail
)

echo.
echo Done. Opening the output folder...
start "" explorer "%~dp0output"
timeout /t 3 >nul
goto :eof

:cancelled
echo Cancelled. No changes were made.
pause
goto :eof

:fail
echo.
pause
goto :eof

:try_python
if defined PY exit /b 0
if not exist "%~1" exit /b 0
"%~1" -c "import sys; raise SystemExit(0 if sys.version_info ^>= (3, 11) else 1)" >nul 2>&1
if not errorlevel 1 set "PY=%~1"
exit /b 0

:run_app
if defined APP_EXE (
    "%APP_EXE%" %*
) else (
    "%PY%" app.py %*
)
exit /b %errorlevel%
