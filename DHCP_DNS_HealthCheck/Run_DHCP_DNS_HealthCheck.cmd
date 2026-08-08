@echo off
TITLE DHCP and DNS Health Check
echo ===============================================================
echo Starting DHCP and DNS Health Check...
echo ===============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0DHCP_DNS_HealthCheck.ps1"
echo.
echo Execution finished.
pause
