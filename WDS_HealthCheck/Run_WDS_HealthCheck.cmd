@echo off
TITLE WDS Health Check
echo ===============================================================
echo Starting WDS Health Check...
echo ===============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0WDS_HealthCheck.ps1"
echo.
echo Execution finished.
pause
