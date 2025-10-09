@echo off
setlocal enabledelayedexpansion

REM Minimal update + restart flow for NSSM service
set "SERVICE_NAME=GB-PDF-Automation"
set "BASE_DIR=C:\Users\ksbot\Documents\UiPath\GB_PDF_Automation"

nssm stop %SERVICE_NAME%
cd /d "%BASE_DIR%"
git pull origin main
"%BASE_DIR%\venv\Scripts\pip.exe" install -r "%BASE_DIR%\requirements.txt"
nssm start %SERVICE_NAME%

endlocal
exit /b 0


