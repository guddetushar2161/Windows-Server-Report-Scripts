@echo off
TITLE Server Health Check
echo ===============================================================
echo Starting Server Health Check...
echo ===============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Server_HealthCheck.ps1"
echo.
echo Execution finished.
pause
