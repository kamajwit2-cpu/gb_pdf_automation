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

REM Start server minimized in a new window using the venv python directly (no need to activate)
REM Title used for controlled restarts: GB-PDF-Automation
set "PYEXE=%CD%\venv\Scripts\python.exe"
if not exist "%PYEXE%" (
    echo ERROR: venv python not found at %PYEXE%
    echo Aborting.
    exit /b 1
)

REM Launch server detached and minimized; redirect output to log
start "GB-PDF-Automation" /min "%PYEXE%" serve.py >> "logs\server.log" 2>&1

echo Server launched. You may close this window.
endlocal
exit /b 0
