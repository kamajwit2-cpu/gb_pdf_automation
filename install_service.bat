@echo off
setlocal enabledelayedexpansion

REM One-time installer for running the app as a Windows Service via NSSM
REM Requirements: nssm.exe available in PATH

set "SERVICE_NAME=GB-PDF-Automation"
set "BASE_DIR=C:\Users\ksbot\Documents\UiPath\GB_PDF_AutomATION"

if not exist "%BASE_DIR%" (
    echo ERROR: Base directory not found: %BASE_DIR%
    exit /b 1
)

cd /d "%BASE_DIR%"

if not exist "venv\Scripts\pythonw.exe" (
    echo Creating virtual environment (Python 3.10)...
    py -3.10 -m venv venv || (
        echo ERROR: Failed to create venv. Ensure Python 3.10 is installed.
        exit /b 1
    )
)

if not exist "logs" mkdir logs

echo Installing service %SERVICE_NAME% via NSSM...
nssm install %SERVICE_NAME% "%BASE_DIR%\venv\Scripts\pythonw.exe" serve.py || (
    echo ERROR: NSSM install failed. Is nssm.exe in PATH?
    exit /b 1
)

echo Configuring working directory and logs...
nssm set %SERVICE_NAME% AppDirectory "%BASE_DIR%"
nssm set %SERVICE_NAME% AppStdout "%BASE_DIR%\logs\server.log"
nssm set %SERVICE_NAME% AppStderr "%BASE_DIR%\logs\server.log"
nssm set %SERVICE_NAME% AppThrottle 1500
nssm set %SERVICE_NAME% AppRestartDelay 2000
nssm set %SERVICE_NAME% Start SERVICE_AUTO_START

echo Installing requirements...
"%BASE_DIR%\venv\Scripts\pip.exe" install -r "%BASE_DIR%\requirements.txt" || (
    echo ERROR: Failed to install requirements.
    exit /b 1
)

echo Starting service...
nssm start %SERVICE_NAME%

echo Service %SERVICE_NAME% installed and started.
endlocal
exit /b 0


