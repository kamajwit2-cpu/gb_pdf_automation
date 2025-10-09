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
for /f "tokens=2 delims==" %%I in ('wmic process where "name='cmd.exe' and CommandLine like '%%GB-PDF-Automation%%'" get ProcessId /value ^| find "ProcessId="') do (
    echo Stopping existing server (PID %%I)...
    taskkill /pid %%I /f >nul 2>&1
)
REM Also try to kill python by window title
taskkill /f /fi "WINDOWTITLE eq GB-PDF-Automation" >nul 2>&1

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

REM Relaunch server minimized via start_server.bat
echo Restarting server...
start "" /min cmd /c start_server.bat

echo Update complete. Exiting updater.
endlocal
exit /b 0

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
echo GB PDF Automation - Update and Restart
echo ========================================
echo.

REM Check if we're in the correct directory
if not exist "requirements.txt" (
    echo ERROR: requirements.txt not found!
    echo Please run this script from the project root directory.
    pause
    exit /b 1
)

REM Stop any running server processes
echo Stopping any running server processes...
taskkill /f /im python.exe /fi "WINDOWTITLE eq GB PDF Automation*" 2>nul
taskkill /f /im python.exe /fi "COMMANDLINE eq *serve.py*" 2>nul
timeout /t 2 /nobreak >nul

REM Pull latest code from GitHub
echo Pulling latest code from GitHub...
git pull origin main
if errorlevel 1 (
    echo WARNING: Git pull failed or no changes found.
    echo This might be normal if no updates are available.
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
echo Starting Updated Production Server...
echo ========================================
echo Server will be available at: http://localhost:5002
echo Server will be accessible from local network
echo Press Ctrl+C to stop the server
echo ========================================
echo.

python serve.py

REM If we get here, the server was stopped
echo.
echo ========================================
echo Server stopped.
echo ========================================
pause
endlocal
