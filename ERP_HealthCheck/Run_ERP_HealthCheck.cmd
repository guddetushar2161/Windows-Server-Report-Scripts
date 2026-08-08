@echo off
TITLE ERP Health Check
echo ===============================================================
echo Starting ERP Health Check...
echo ===============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ERP_HealthCheck.ps1"
echo.
echo Execution finished.
pause
