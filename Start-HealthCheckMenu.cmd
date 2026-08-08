@echo off
TITLE Windows Server Health Checks - Master Menu
echo Starting Master Menu...
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Maximized -File "%~dp0MasterMenu.ps1"
