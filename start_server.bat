@echo off
setlocal enabledelayedexpansion

REM Force working directory to the server folder
set "BASE_DIR=C:\Users\ksbot\Documents\UiPath\GB_PDF_Automation"
if not exist "%BASE_DIR%" (
    echo ERROR: Base directory not found: %BASE_DIR%
    pause
    exit /b 1
)
cd /d "%BASE_DIR%"

echo ========================================
echo GB PDF Automation - Production Server
echo ========================================
echo.

REM Check if we're in the correct directory
if not exist "requirements.txt" (
    echo ERROR: requirements.txt not found!
    echo Please run this script from the project root directory.
    pause
    exit /b 1
)

REM Check if virtual environment exists, create if not
if not exist "venv" (
    echo Creating virtual environment...
    py -3.10 -m venv venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment!
        echo Please ensure Python 3.10 is installed.
        pause
        exit /b 1
    )
    echo Virtual environment created successfully.
) else (
    echo Virtual environment found.
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment!
    pause
    exit /b 1
)

REM Install/update requirements
echo Installing/updating requirements...
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install requirements!
    pause
    exit /b 1
)

REM Set production environment variable
set ENV=production

REM Start the production server
echo.
echo ========================================
echo Starting Production Server (detached)...
echo ========================================
echo Server will be available at: http://localhost:5002
echo Server will be accessible from local network
echo A minimized window will run in the background.
echo ========================================
echo.

REM Ensure logs directory exists
if not exist "logs" mkdir logs

REM Prefer pythonw.exe (no console) and fallback to python.exe minimized
set "PYEXE_W=%CD%\venv\Scripts\pythonw.exe"
set "PYEXE_C=%CD%\venv\Scripts\python.exe"

if exist "%PYEXE_W%" (
    REM Launch without visible window; redirect stdout/stderr to log via cmd /c
    start "GB-PDF-Automation" /b cmd /c ""%PYEXE_W%" serve.py >> "logs\server.log" 2>&1"
) else if exist "%PYEXE_C%" (
    REM Launch minimized console and log output
    start "GB-PDF-Automation" /min "%PYEXE_C%" serve.py >> "logs\server.log" 2>&1
) else (
    echo ERROR: venv python not found.
    exit /b 1
)

echo Server launched. You may close this window.
endlocal
exit /b 0
