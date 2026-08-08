@echo off
TITLE AD Health Check
echo ===============================================================
echo Starting AD Health Check...
echo ===============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0AD_HealthCheck.ps1"
echo.
echo Execution finished.
pause
