@echo off
TITLE ADCS Health Check
echo ===============================================================
echo Starting ADCS Health Check...
echo ===============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ADCS_HealthCheck.ps1"
echo.
echo Execution finished.
pause
