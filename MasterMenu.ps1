#Requires -Version 5.1
<#
.SYNOPSIS
    Interactive Master Menu for Windows Server Health Checks
#>

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

function Show-Menu {
    Clear-Host
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "    Windows Server Health Check - Master Interactive Menu" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Please select a health check to run:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [1] Active Directory (AD)"
    Write-Host "  [2] AD Certificate Services (ADCS)"
    Write-Host "  [3] DHCP and DNS"
    Write-Host "  [4] ERP Server"
    Write-Host "  [5] General Server Baseline"
    Write-Host "  [6] Windows Deployment Services (WDS)"
    Write-Host "  [7] WSUS Server"
    Write-Host ""
    Write-Host "  [A] Run ALL Health Checks sequentially"
    Write-Host "  [Q] Quit"
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
}

function Run-Check {
    param([string]$Path, [string]$Name)
    $fullPath = Join-Path $ScriptDir $Path
    if (Test-Path $fullPath) {
        Write-Host "`n[+] Starting $Name..." -ForegroundColor Green
        & $fullPath
    } else {
        Write-Host "`n[-] Error: Cannot find script at $fullPath" -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

while ($true) {
    Show-Menu
    $choice = Read-Host "Enter your choice"
    
    switch ($choice) {
        '1' { Run-Check "AD_HealthCheck\AD_HealthCheck.ps1" "AD Health Check" }
        '2' { Run-Check "ADCS_HealthCheck\ADCS_HealthCheck.ps1" "ADCS Health Check" }
        '3' { Run-Check "DHCP_DNS_HealthCheck\DHCP_DNS_HealthCheck.ps1" "DHCP/DNS Health Check" }
        '4' { Run-Check "ERP_HealthCheck\ERP_HealthCheck.ps1" "ERP Health Check" }
        '5' { Run-Check "Server_HealthCheck\Server_HealthCheck.ps1" "Server Health Check" }
        '6' { Run-Check "WDS_HealthCheck\WDS_HealthCheck.ps1" "WDS Health Check" }
        '7' { Run-Check "WSUS_HealthCheck\WSUS_HealthCheck.ps1" "WSUS Health Check" }
        'A' {
            Run-Check "AD_HealthCheck\AD_HealthCheck.ps1" "AD Health Check"
            Run-Check "ADCS_HealthCheck\ADCS_HealthCheck.ps1" "ADCS Health Check"
            Run-Check "DHCP_DNS_HealthCheck\DHCP_DNS_HealthCheck.ps1" "DHCP/DNS Health Check"
            Run-Check "ERP_HealthCheck\ERP_HealthCheck.ps1" "ERP Health Check"
            Run-Check "Server_HealthCheck\Server_HealthCheck.ps1" "Server Health Check"
            Run-Check "WDS_HealthCheck\WDS_HealthCheck.ps1" "WDS Health Check"
            Run-Check "WSUS_HealthCheck\WSUS_HealthCheck.ps1" "WSUS Health Check"
        }
        'Q' { Write-Host "`nExiting..." -ForegroundColor Yellow; exit }
        default { Write-Host "`nInvalid selection. Please try again." -ForegroundColor Red; Start-Sleep -Seconds 1 }
    }
    
    if ($choice -ne 'Q') {
        Write-Host "`nPress any key to return to the menu..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}
