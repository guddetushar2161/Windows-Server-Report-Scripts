@echo off
TITLE WSUS Health Check
echo ===============================================================
echo Starting WSUS Health Check...
echo ===============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0WSUS_HealthCheck.ps1"
echo.
echo Execution finished.
pause
