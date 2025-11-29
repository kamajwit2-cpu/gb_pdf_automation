@echo off
setlocal enabledelayedexpansion

REM Force working directory to the server folder
set "BASE_DIR=C:\Users\ksbot\Documents\UiPath\GB_PDF_Automation"
if not exist "%BASE_DIR%" (
    echo ERROR: Base directory not found: %BASE_DIR%
    exit /b 1
)
cd /d "%BASE_DIR%"

echo ========================================
echo GB PDF Automation - Update & Restart
echo ========================================

REM Kill existing server window by title (if any)
REM Try to stop any python/pythonw processes running serve.py (more reliable)
for /f "tokens=2 delims==" %%I in ('wmic process where "(name='python.exe' or name='pythonw.exe') and CommandLine like '%%serve.py%%'" get ProcessId /value ^| find "ProcessId="') do (
    echo Stopping existing python server (PID %%I)...
    taskkill /pid %%I /f >nul 2>&1
)

REM Git update
echo Pulling latest code...
git pull origin main
if errorlevel 1 (
    echo ERROR: git pull failed.
    exit /b 1
)

REM Ensure venv exists
if not exist "venv" (
    echo Creating virtual environment...
    py -3.10 -m venv venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment.
        exit /b 1
    )
)

REM Install requirements without activating venv
set "PYEXE=%CD%\venv\Scripts\python.exe"
set "PIPEXE=%CD%\venv\Scripts\pip.exe"
if not exist "%PIPEXE%" (
    echo ERROR: pip not found at %PIPEXE%
    exit /b 1
)

echo Installing/updating requirements...
"%PIPEXE%" install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install requirements.
    exit /b 1
)

REM Relaunch server headless using pythonw when available (no window)
echo Restarting server (headless)...
set "PYEXE_W=%CD%\venv\Scripts\pythonw.exe"
set "PYEXE_C=%CD%\venv\Scripts\python.exe"

if exist "%PYEXE_W%" (
    REM Use pythonw without window; redirect output to log via cmd /c
    if not exist "logs" mkdir logs
    start "" /b cmd /c ""%PYEXE_W%" serve.py >> "logs\server.log" 2>&1"
) else if exist "%PYEXE_C%" (
    REM Fallback: minimized console
    if not exist "logs" mkdir logs
    start "GB-PDF-Automation" /min "%PYEXE_C%" serve.py >> "logs\server.log" 2>&1
) else (
    echo ERROR: venv python not found; cannot restart server.
    exit /b 1
)

echo Update complete. Exiting updater.
endlocal
exit /b 0

 
