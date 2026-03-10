#Requires -Version 5.1
<#
.SYNOPSIS
    Comprehensive ERP Server Health Check Script (ERP-Agnostic)

.DESCRIPTION
    Performs a thorough health check of the ERP server environment and
    exports the results to a single self-contained HTML dashboard file.

    Compatible with SAP, Oracle EBS, Microsoft Dynamics, or any Windows-hosted
    ERP system.  Uses only Windows-native tools — no ERP-specific SDK required.

    Checks performed:
      1.  Server Baseline          (hostname, OS, uptime, CPU, RAM, .NET versions)
      2.  CPU & Memory Performance (live CPU%, top 5 process lists)
      3.  Disk Health              (all volumes with bars, warn >85%, crit >95%)
      4.  ERP Application Services (status, start type, PID, uptime per service)
      5.  Database Service & Connectivity (service state + SQL SELECT 1 test)
      6.  Database Details         (sizes, last backup date, recovery model)
      7.  IIS / Web Server         (app pools, site bindings, worker processes)
      8.  Network Port Connectivity(ERPPort, DB port, ExtraPorts — Open/Closed)
      9.  Windows Firewall         (profile state, rules allowing ERPPort)
      10. Scheduled Tasks          (ERP/backup/sync/job/batch tasks)
      11. Event Log Check          (last 24 h Critical/Error from App & System)
      12. Certificate Check        (SSL certs expiring within 30/60 days)
      13. Backup Verification      (*.bak / *.zip modified in last 24 h)

.PARAMETER ERPName
    Friendly display name for the ERP system shown in the report title.
    Default: 'ERP System'

.PARAMETER AppServiceNames
    Array of Windows service names to monitor as ERP application services.
    Example: @('SAPService00', 'SAPD00')

.PARAMETER DBServiceName
    Windows service name for the database engine.
    Example: 'MSSQLSERVER', 'MSSQL$SQLEXPRESS', 'OracleServiceORCL'

.PARAMETER DBInstance
    SQL Server instance used for the connectivity test (Windows Auth).
    Only tested when DBServiceName matches MSSQL.
    Example: 'localhost', '.\SQLEXPRESS', 'DBSERVER\PROD'

.PARAMETER ERPPort
    Primary ERP application TCP port to test connectivity on.

.PARAMETER ExtraPorts
    Additional TCP ports to test alongside ERPPort and the DB default port.

.NOTES
    Version    : 1.0.0
    Author     : Tushar Gudde
    Requires   : PowerShell 5.1+, Local Administrator rights
    Compatible : Windows Server 2016, 2019, 2022
    Companion  : ERP_HealthCheck_EmailAlert.ps1
#>

param(
    [string]   $ERPName         = 'ERP System',
    [string[]] $AppServiceNames = @('W3SVC'),
    [string]   $DBServiceName   = 'MSSQLSERVER',
    [string]   $DBInstance      = 'localhost',
    [int]      $ERPPort         = 8080,
    [int[]]    $ExtraPorts      = @()
)

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
$CompanyLogoURL = ''                         # URL/path to logo image. Leave blank to skip.
$CompanyWebsite = 'https://tushargudde.tech' # Company website URL for logo hyperlink.
$AuthorName     = 'Tushar Gudde'             # Author name shown in footer.

# Disk usage warning thresholds (%)
$DiskWarnPct    = 85
$DiskCritPct    = 95

# Certificate expiry warning thresholds (days ahead)
$CertWarnDays   = 60
$CertCritDays   = 30

# Backup file search: paths checked for recent backup files
$BackupSearchPaths = @(
    'C:\Backup', 'D:\Backup', 'E:\Backup',
    'C:\ERP\Backup', 'D:\ERP\Backup',
    'C:\MSSQL\Backup',
    'C:\Program Files\Microsoft SQL Server\MSSQL\Backup'
)
$BackupMaxAgeHours = 24

# Scheduled task name filter (regex).  Tasks whose names match this pattern
# are shown in Section 10.  Extend with your ERP system's task-naming convention.
$TaskNamePattern  = 'ERP|backup|sync|job|batch|report|sap|oracle|dynamics'

# Write a plain-text companion status file (_HEALTHY.txt / _CRITICAL.txt)
# alongside every HTML report.  Required by ERP_HealthCheck_EmailAlert.ps1.
$EnableStatusFile = $true
# ──────────────────────────────────────────────────────────────────────────────

$ScriptVersion = '1.0.0'
$StartTime     = Get-Date
$ScriptDir     = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ([string]::IsNullOrEmpty($ScriptDir)) { $ScriptDir = $PWD.Path }

# ── REPORTS FOLDER ────────────────────────────────────────────────────────────
$ReportsDir = Join-Path $ScriptDir 'Reports'
if (-not (Test-Path $ReportsDir)) {
    try { New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null }
    catch { Write-Warning "Could not create Reports folder: $_" }
}
$ReportStamp = "ERP_Health_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
$ReportFile  = Join-Path $ReportsDir ($ReportStamp + '.html')

# ── HELPER FUNCTIONS ──────────────────────────────────────────────────────────
function HtmlEncode {
    param([string]$text)
    if ([string]::IsNullOrEmpty($text)) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function StatusBadge {
    param([string]$text, [string]$color)
    $map = @{
        green  = '#2ea043'
        yellow = '#d29922'
        red    = '#da3633'
        blue   = '#1f6feb'
        grey   = '#484f58'
    }
    $bg = if ($map.ContainsKey($color)) { $map[$color] } else { '#484f58' }
    return "<span class='badge' style='background:$bg'>$(HtmlEncode $text)</span>"
}

function Write-Progress2 {
    param([string]$msg)
    Write-Host "  [$(Get-Date -Format 'HH:mm:ss')] $msg" -ForegroundColor Cyan
}

function BuildKVTable {
    param([array]$rows)   # each element: @('Label', 'ValueHtml')
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<div class='table-wrap'><table class='kv-table'><tbody>")
    foreach ($r in $rows) {
        [void]$sb.Append("<tr><td class='td-label'>$(HtmlEncode $r[0])</td><td>$($r[1])</td></tr>")
    }
    [void]$sb.Append("</tbody></table></div>")
    return $sb.ToString()
}

function BuildSection {
    param(
        [int]$num,
        [string]$title,
        [string]$body,
        [bool]$hasError = $false,
        [bool]$open     = $false
    )
    $openAttr  = if ($open) { ' open' } else { '' }
    $indicator = if ($hasError) {
        "<span style='color:#da3633;'>&#x2716;</span>"
    } else {
        "<span style='color:#2ea043;'>&#x2714;</span>"
    }
    return @"
<details class='section-card'$openAttr>
  <summary class='section-summary'>
    <span class='sec-arrow'>&#9654;</span>
    <span class='sec-num'>$num</span>
    <span class='sec-title'>$(HtmlEncode $title)</span>
    $indicator
  </summary>
  <div class='section-body'>
$body
  </div>
</details>
"@
}

function Test-TcpPort {
    param([string]$ComputerName = 'localhost', [int]$Port, [int]$TimeoutMs = 2000)
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $asyncRes  = $tcpClient.BeginConnect($ComputerName, $Port, $null, $null)
        $connected = $asyncRes.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if ($connected -and $tcpClient.Connected) {
            $tcpClient.EndConnect($asyncRes)
            $tcpClient.Close()
            return $true
        }
        $tcpClient.Close()
        return $false
    } catch {
        return $false
    }
}

# ── CRITICAL FINDINGS LIST ────────────────────────────────────────────────────
$CriticalFindings = [System.Collections.Generic.List[string]]::new()

# ── DEFAULT KPI VARIABLES (overridden in sections below) ─────────────────────
$ServerHostname    = $env:COMPUTERNAME
$UptimeStr         = 'Unknown'
$CpuPct            = 0
$RamUsedPct        = 0
$TotalRAM_GB       = 0
$UsedRAM_GB        = 0
$FreeRAM_GB        = 0
$DiskWorstPct      = 0
$ServiceOkCount    = 0
$ServiceTotalCount = $AppServiceNames.Count
$dbConnected       = $false

Write-Host ""
Write-Host "+==============================================================+" -ForegroundColor DarkGreen
Write-Host "|   ERP Server Health Check  v$ScriptVersion" -ForegroundColor DarkGreen
Write-Host "|   ERP System: $ERPName" -ForegroundColor DarkGreen
Write-Host "+==============================================================+" -ForegroundColor DarkGreen
Write-Host ""

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — SERVER BASELINE
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 1: Server Baseline..."
$Sec1Html = ''
try {
    $cs   = Get-CimInstance Win32_ComputerSystem  -ErrorAction Stop
    $os   = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $bios = Get-CimInstance Win32_BIOS            -ErrorAction Stop
    $cpus = @(Get-CimInstance Win32_Processor     -ErrorAction Stop)

    $TotalRAM_GB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
    # FreePhysicalMemory is in KB.  1 GB = 1024 * 1024 KB, so dividing by
    # (1024 * 1024) converts KB directly to GB without intermediate unit steps.
    $FreeRAM_GB  = [math]::Round($os.FreePhysicalMemory / (1024 * 1024), 2)
    $UsedRAM_GB  = [math]::Round($TotalRAM_GB - $FreeRAM_GB, 2)
    $RamUsedPct  = if ($TotalRAM_GB -gt 0) {
        [math]::Round(($UsedRAM_GB / $TotalRAM_GB) * 100, 1)
    } else { 0 }

    $ServerHostname = $env:COMPUTERNAME
    $bootTime = $os.LastBootUpTime
    if ($null -ne $bootTime) {
        $uptime     = New-TimeSpan -Start $bootTime -End (Get-Date)
        $UptimeStr  = '{0}d {1}h {2}m' -f [int]$uptime.TotalDays, $uptime.Hours, $uptime.Minutes
        $lastBootStr = $bootTime.ToString('yyyy-MM-dd HH:mm:ss')
    } else {
        $UptimeStr   = 'Unknown'
        $lastBootStr = 'N/A'
    }

    $isVirtual = ($cs.Model        -match 'Virtual|VMware|VirtualBox|QEMU|KVM|Xen|HVM') -or
                 ($cs.Manufacturer -match 'VMware|QEMU|Xen|Parallels|innotek') -or
                 ($cs.Manufacturer -eq 'Microsoft Corporation' -and $cs.Model -match 'Virtual')
    $serverTypeBadge = if ($isVirtual) { StatusBadge 'Virtual Machine' 'blue' } else { StatusBadge 'Physical Server' 'green' }

    $cpuNames = ($cpus | ForEach-Object {
        if ($null -ne $_.Name) { $_.Name.Trim() } else { 'Unknown' }
    } | Select-Object -Unique) -join '; '
    $cpuCores = if ($cpus.Count -gt 0 -and $null -ne $cpus[0].NumberOfLogicalProcessors) {
        $cpus[0].NumberOfLogicalProcessors
    } else { 'N/A' }

    # .NET Framework versions from registry
    $dotNetList = @()
    try {
        $dotNetList = @(
            Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.GetValue('Version') } |
            ForEach-Object { $_.GetValue('Version') } |
            Sort-Object -Unique |
            Sort-Object -Descending |
            Select-Object -First 6
        )
    } catch {}
    $dotNetStr = if ($dotNetList.Count -gt 0) { $dotNetList -join ', ' } else { 'Unknown' }

    $rows1 = @(
        @('Hostname',       (HtmlEncode $ServerHostname)),
        @('Server Type',    $serverTypeBadge),
        @('Manufacturer',   (HtmlEncode $cs.Manufacturer)),
        @('Model',          (HtmlEncode $cs.Model)),
        @('Serial Number',  (HtmlEncode $bios.SerialNumber)),
        @('OS Name',        (HtmlEncode $os.Caption)),
        @('OS Version',     (HtmlEncode $os.Version)),
        @('OS Build',       (HtmlEncode $os.BuildNumber)),
        @('CPU',            "$($cpus.Count) x $(HtmlEncode $cpuNames) &mdash; $cpuCores logical cores"),
        @('Total RAM',      "$TotalRAM_GB GB"),
        @('Used RAM',       "$UsedRAM_GB GB ($RamUsedPct%)"),
        @('Free RAM',       "$FreeRAM_GB GB"),
        @('Last Boot Time', (HtmlEncode $lastBootStr)),
        @('System Uptime',  (HtmlEncode $UptimeStr)),
        @('.NET Versions',  (HtmlEncode $dotNetStr))
    )
    $Sec1Html = BuildKVTable $rows1
} catch {
    $Sec1Html = "<p class='error'>Error retrieving server baseline: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 1 - Server baseline error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — CPU & MEMORY PERFORMANCE
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 2: CPU & Memory Performance..."
$Sec2Html = ''
try {
    # CPU load from WMI (LoadPercentage = current utilisation)
    try {
        $cpuInstances = @(Get-CimInstance Win32_Processor -ErrorAction Stop)
        $CpuPct = [math]::Round(
            ($cpuInstances | Measure-Object -Property LoadPercentage -Average).Average, 1)
    } catch { $CpuPct = 0 }

    # Top 5 processes by CPU time and by RAM
    $topCpuProcs = @()
    $topRamProcs = @()
    try {
        $allProcs    = @(Get-Process -ErrorAction Stop)
        $topCpuProcs = $allProcs |
            Where-Object { $null -ne $_.CPU } |
            Sort-Object CPU -Descending |
            Select-Object -First 5 |
            Select-Object Name, Id,
                @{N='CPU_s';  E={ [math]::Round($_.CPU, 1) }},
                @{N='RAM_MB'; E={ [math]::Round($_.WorkingSet64 / 1MB, 1) }}
        $topRamProcs = $allProcs |
            Sort-Object WorkingSet64 -Descending |
            Select-Object -First 5 |
            Select-Object Name, Id,
                @{N='CPU_s';  E={ if ($null -ne $_.CPU) { [math]::Round($_.CPU, 1) } else { 'N/A' } }},
                @{N='RAM_MB'; E={ [math]::Round($_.WorkingSet64 / 1MB, 1) }}
    } catch {}

    $cpuBadge = if ($CpuPct -ge 90) { StatusBadge "$CpuPct%" 'red' }
                elseif ($CpuPct -ge 75) { StatusBadge "$CpuPct%" 'yellow' }
                else { StatusBadge "$CpuPct%" 'green' }
    if ($CpuPct -ge 90) {
        $CriticalFindings.Add("Section 2 - CPU utilisation critically high: $CpuPct%")
    }

    $ramBadge = if ($RamUsedPct -ge 90) { StatusBadge "$RamUsedPct%" 'red' }
                elseif ($RamUsedPct -ge 75) { StatusBadge "$RamUsedPct%" 'yellow' }
                else { StatusBadge "$RamUsedPct%" 'green' }
    if ($RamUsedPct -ge 90) {
        $CriticalFindings.Add("Section 2 - RAM utilisation critically high: $RamUsedPct%")
    }

    $Sec2Html  = "<div style='display:flex;flex-wrap:wrap;gap:20px;margin-bottom:14px;'>"
    $Sec2Html += "<div><div style='color:var(--muted);font-size:.75rem;margin-bottom:4px;'>CPU Load</div>$cpuBadge</div>"
    $Sec2Html += "<div><div style='color:var(--muted);font-size:.75rem;margin-bottom:4px;'>RAM Used</div>$ramBadge</div>"
    $Sec2Html += "<div><div style='color:var(--muted);font-size:.75rem;margin-bottom:4px;'>Available RAM</div><strong>$FreeRAM_GB GB</strong></div>"
    $Sec2Html += "<div><div style='color:var(--muted);font-size:.75rem;margin-bottom:4px;'>Total RAM</div><strong>$TotalRAM_GB GB</strong></div>"
    $Sec2Html += "</div>"

    $Sec2Html += "<h4 style='margin:8px 0 6px;'>Top 5 Processes &mdash; CPU Time</h4>"
    $Sec2Html += "<div class='table-wrap'><table><thead><tr><th>Process Name</th><th>PID</th><th>CPU Time (s)</th><th>RAM (MB)</th></tr></thead><tbody>"
    if ($topCpuProcs.Count -gt 0) {
        foreach ($p in $topCpuProcs) {
            $Sec2Html += "<tr><td>$(HtmlEncode $p.Name)</td><td>$($p.Id)</td><td>$($p.CPU_s)</td><td>$($p.RAM_MB)</td></tr>"
        }
    } else {
        $Sec2Html += "<tr><td colspan='4' class='info'>No process data available.</td></tr>"
    }
    $Sec2Html += "</tbody></table></div>"

    $Sec2Html += "<h4 style='margin:12px 0 6px;'>Top 5 Processes &mdash; RAM Usage</h4>"
    $Sec2Html += "<div class='table-wrap'><table><thead><tr><th>Process Name</th><th>PID</th><th>CPU Time (s)</th><th>RAM (MB)</th></tr></thead><tbody>"
    if ($topRamProcs.Count -gt 0) {
        foreach ($p in $topRamProcs) {
            $Sec2Html += "<tr><td>$(HtmlEncode $p.Name)</td><td>$($p.Id)</td><td>$($p.CPU_s)</td><td>$($p.RAM_MB)</td></tr>"
        }
    } else {
        $Sec2Html += "<tr><td colspan='4' class='info'>No process data available.</td></tr>"
    }
    $Sec2Html += "</tbody></table></div>"
} catch {
    $Sec2Html = "<p class='error'>Error retrieving CPU/memory performance: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 2 - CPU/Memory check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — DISK HEALTH
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 3: Disk Health..."
$Sec3Html     = ''
$DiskInfo     = @()
$DiskWorstPct = 0
try {
    $logicalDisks = @(Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop)
    foreach ($disk in $logicalDisks) {
        if ($null -ne $disk.Size -and $disk.Size -gt 0) {
            $totalGB = [math]::Round($disk.Size      / 1GB, 2)
            $freeGB  = [math]::Round($disk.FreeSpace / 1GB, 2)
            $usedGB  = [math]::Round($totalGB - $freeGB, 2)
            $pctUsed = [math]::Round(($usedGB / $totalGB) * 100, 1)

            if ($pctUsed -gt $DiskWorstPct) { $DiskWorstPct = $pctUsed }

            if ($pctUsed -ge $DiskCritPct) {
                $CriticalFindings.Add("Section 3 - Disk $($disk.DeviceID) is ${pctUsed}% full (CRITICAL)")
            } elseif ($pctUsed -ge $DiskWarnPct) {
                $CriticalFindings.Add("Section 3 - Disk $($disk.DeviceID) is ${pctUsed}% full (Warning)")
            }

            $DiskInfo += [PSCustomObject]@{
                Drive   = $disk.DeviceID
                Label   = if ([string]::IsNullOrEmpty($disk.VolumeName)) { 'Local Disk' } else { $disk.VolumeName }
                TotalGB = $totalGB
                UsedGB  = $usedGB
                FreeGB  = $freeGB
                PctUsed = $pctUsed
            }
        }
    }

    if ($DiskInfo.Count -eq 0) {
        $Sec3Html = "<p class='warn'>No fixed disk volumes found.</p>"
    } else {
        # Visual progress bars
        $Sec3Html = "<div class='disk-container'>"
        foreach ($d in $DiskInfo) {
            $barColor    = if ($d.PctUsed -ge $DiskCritPct)  { '#da3633' }
                           elseif ($d.PctUsed -ge $DiskWarnPct) { '#d29922' }
                           else { '#2ea043' }
            $statusStr   = if ($d.PctUsed -ge $DiskCritPct)  { 'CRITICAL' }
                           elseif ($d.PctUsed -ge $DiskWarnPct) { 'Warning' }
                           else { 'OK' }
            $statusColor = if ($d.PctUsed -ge $DiskCritPct)  { 'red' }
                           elseif ($d.PctUsed -ge $DiskWarnPct) { 'yellow' }
                           else { 'green' }
            $badge = StatusBadge $statusStr $statusColor
            $Sec3Html += "<div class='disk-row'>"
            $Sec3Html += "<div class='disk-label'><strong>$(HtmlEncode $d.Drive)</strong>&nbsp;$(HtmlEncode $d.Label)</div>"
            $Sec3Html += "<div class='disk-bar-outer'><div class='disk-bar-inner' style='width:$($d.PctUsed)%;background:$barColor;'></div></div>"
            $Sec3Html += "<div class='disk-stat'>$($d.FreeGB) GB free / $($d.TotalGB) GB total &mdash; $($d.PctUsed)% used &nbsp;$badge</div>"
            $Sec3Html += "</div>"
        }
        $Sec3Html += "</div>"

        # Detail table
        $Sec3Html += "<div class='table-wrap' style='margin-top:14px;'><table>"
        $Sec3Html += "<thead><tr><th>Drive</th><th>Label</th><th>Total GB</th><th>Used GB</th><th>Free GB</th><th>% Used</th><th>Status</th></tr></thead><tbody>"
        foreach ($d in $DiskInfo) {
            $st = if ($d.PctUsed -ge $DiskCritPct) { 'CRITICAL' } elseif ($d.PctUsed -ge $DiskWarnPct) { 'Warning' } else { 'OK' }
            $sc = if ($d.PctUsed -ge $DiskCritPct) { 'red' }      elseif ($d.PctUsed -ge $DiskWarnPct) { 'yellow' }  else { 'green' }
            $Sec3Html += "<tr><td>$(HtmlEncode $d.Drive)</td><td>$(HtmlEncode $d.Label)</td>"
            $Sec3Html += "<td>$($d.TotalGB)</td><td>$($d.UsedGB)</td><td>$($d.FreeGB)</td>"
            $Sec3Html += "<td>$($d.PctUsed)%</td><td>$(StatusBadge $st $sc)</td></tr>"
        }
        $Sec3Html += "</tbody></table></div>"
    }
} catch {
    $Sec3Html = "<p class='error'>Error checking disk health: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 3 - Disk check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — ERP APPLICATION SERVICES
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 4: ERP Application Services..."
$Sec4Html       = ''
$ServiceResults = [System.Collections.Generic.List[hashtable]]::new()
$ServiceOkCount = 0

foreach ($svcName in $AppServiceNames) {
    $r = @{
        Name      = $svcName
        Status    = 'Not Found'
        StartType = 'N/A'
        PID       = 'N/A'
        Uptime    = 'N/A'
        Badge     = StatusBadge 'Not Found' 'grey'
    }
    try {
        $svc       = Get-Service -Name $svcName -ErrorAction Stop
        $r.Status  = $svc.Status.ToString()
        $r.StartType = $svc.StartType.ToString()

        # PID and uptime via WMI/CIM
        try {
            $wmiSvc = Get-CimInstance Win32_Service -Filter "Name='$svcName'" -ErrorAction SilentlyContinue
            if ($null -ne $wmiSvc -and $wmiSvc.ProcessId -gt 0) {
                $r.PID = $wmiSvc.ProcessId.ToString()
                try {
                    $proc = Get-Process -Id $wmiSvc.ProcessId -ErrorAction Stop
                    if ($null -ne $proc.StartTime) {
                        $svcUp = New-TimeSpan -Start $proc.StartTime -End (Get-Date)
                        $r.Uptime = '{0}d {1}h {2}m' -f [int]$svcUp.TotalDays, $svcUp.Hours, $svcUp.Minutes
                    }
                } catch {}
            }
        } catch {}

        if ($svc.Status -eq 'Running') {
            $r.Badge = StatusBadge 'Running' 'green'
            $ServiceOkCount++
        } elseif ($svc.Status -eq 'Stopped') {
            $r.Badge = StatusBadge 'Stopped' 'red'
            $CriticalFindings.Add("Section 4 - ERP service '$svcName' is Stopped")
        } else {
            $r.Badge = StatusBadge $svc.Status.ToString() 'yellow'
        }
    } catch {
        $r.Badge = StatusBadge 'Not Found' 'grey'
        $CriticalFindings.Add("Section 4 - ERP service '$svcName' not found")
    }
    $ServiceResults.Add($r)
}

if ($ServiceResults.Count -eq 0) {
    $Sec4Html = "<p class='warn'>No ERP services configured. Set <code>`$AppServiceNames</code> in the script parameters.</p>"
} else {
    $Sec4Html  = "<div class='table-wrap'><table><thead><tr>"
    $Sec4Html += "<th>Service Name</th><th>Status</th><th>Start Type</th><th>PID</th><th>Uptime</th>"
    $Sec4Html += "</tr></thead><tbody>"
    foreach ($r in $ServiceResults) {
        $Sec4Html += "<tr>"
        $Sec4Html += "<td>$(HtmlEncode $r.Name)</td><td>$($r.Badge)</td>"
        $Sec4Html += "<td>$(HtmlEncode $r.StartType)</td><td>$(HtmlEncode $r.PID)</td><td>$(HtmlEncode $r.Uptime)</td>"
        $Sec4Html += "</tr>"
    }
    $Sec4Html += "</tbody></table></div>"
    $Sec4Html += "<p class='info' style='margin-top:8px;'>$ServiceOkCount of $ServiceTotalCount service(s) running.</p>"
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — DATABASE SERVICE & CONNECTIVITY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 5: Database Service & Connectivity..."
$Sec5Html = ''
try {
    # Database service state
    $dbSvcBadge = StatusBadge 'Not Checked' 'grey'
    try {
        $dbSvc = Get-Service -Name $DBServiceName -ErrorAction Stop
        if ($dbSvc.Status -eq 'Running') {
            $dbSvcBadge = StatusBadge 'Running' 'green'
        } else {
            $dbSvcBadge = StatusBadge $dbSvc.Status.ToString() 'red'
            $CriticalFindings.Add("Section 5 - Database service '$DBServiceName' is $($dbSvc.Status)")
        }
    } catch {
        $dbSvcBadge = StatusBadge 'Not Found' 'grey'
        $CriticalFindings.Add("Section 5 - Database service '$DBServiceName' not found")
    }

    # SQL connectivity test (Windows Authentication) — only for MSSQL-type services
    $connBadge   = StatusBadge 'Not Tested' 'grey'
    $connNote    = 'SQL connectivity test skipped (service name does not match MSSQL).'
    $dbLatencyStr = 'N/A'

    if ($DBServiceName -match 'MSSQL|SQL') {
        try {
            $connStr = "Server=$DBInstance;Integrated Security=True;Connection Timeout=10;"
            $conn    = New-Object System.Data.SqlClient.SqlConnection($connStr)
            $sw      = [System.Diagnostics.Stopwatch]::StartNew()
            $conn.Open()
            $cmd     = $conn.CreateCommand()
            $cmd.CommandText    = 'SELECT 1'
            $cmd.CommandTimeout = 10
            $null = $cmd.ExecuteScalar()
            $sw.Stop()
            $latencyMs    = $sw.ElapsedMilliseconds
            $dbLatencyStr = "$latencyMs ms"
            $conn.Close()
            $conn.Dispose()
            $dbConnected = $true
            $connBadge   = StatusBadge 'Connected' 'green'
            $connNote    = "SELECT 1 completed in $latencyMs ms via Windows Authentication."
        } catch {
            $connBadge = StatusBadge 'Failed' 'red'
            $connNote  = "Connection failed: $(HtmlEncode $_.Exception.Message)"
            $CriticalFindings.Add("Section 5 - SQL connection to '$DBInstance' failed: $($_.Exception.Message)")
        }
    }

    $rows5 = @(
        @('DB Service Name',  (HtmlEncode $DBServiceName)),
        @('Service Status',   $dbSvcBadge),
        @('DB Instance',      (HtmlEncode $DBInstance)),
        @('Connectivity',     $connBadge),
        @('Query Latency',    (HtmlEncode $dbLatencyStr)),
        @('Connection Note',  $connNote)
    )
    $Sec5Html = BuildKVTable $rows5
} catch {
    $Sec5Html = "<p class='error'>Error checking database service: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 5 - Database service check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — DATABASE DETAILS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 6: Database Details..."
$Sec6Html = ''

if (-not $dbConnected) {
    $Sec6Html = "<p class='warn'>&#x26A0; Database connection not available &mdash; skipping database details.  " +
                "Ensure <code>`$DBInstance</code> is reachable and <code>`$DBServiceName</code> matches an MSSQL service.</p>"
} else {
    try {
        $connStr6 = "Server=$DBInstance;Integrated Security=True;Connection Timeout=10;"
        $conn6    = New-Object System.Data.SqlClient.SqlConnection($connStr6)
        $conn6.Open()

        # ── DB sizes ──────────────────────────────────────────────────────────
        $dbSizeHtml = ''
        try {
            $sqlSizes = "SELECT d.name AS DBName, " +
                "CAST(SUM(CASE WHEN mf.type=0 THEN mf.size END)*8.0/1024/1024 AS DECIMAL(10,2)) AS DataGB, " +
                "CAST(SUM(CASE WHEN mf.type=1 THEN mf.size END)*8.0/1024/1024 AS DECIMAL(10,2)) AS LogGB, " +
                "d.state_desc AS State, d.recovery_model_desc AS RecoveryModel " +
                "FROM sys.databases d " +
                "JOIN sys.master_files mf ON d.database_id=mf.database_id " +
                "WHERE d.database_id>0 " +
                "GROUP BY d.name,d.state_desc,d.recovery_model_desc ORDER BY DataGB DESC"
            $cmd6 = $conn6.CreateCommand()
            $cmd6.CommandText    = $sqlSizes
            $cmd6.CommandTimeout = 30
            $rdr6 = $cmd6.ExecuteReader()
            $dbSizeHtml  = "<div class='table-wrap'><table><thead><tr>"
            $dbSizeHtml += "<th>Database</th><th>Data (GB)</th><th>Log (GB)</th><th>State</th><th>Recovery Model</th>"
            $dbSizeHtml += "</tr></thead><tbody>"
            while ($rdr6.Read()) {
                $dbName   = [string]$rdr6['DBName']
                $dataGB   = if ($rdr6.IsDBNull($rdr6.GetOrdinal('DataGB')))  { 'N/A' } else { [string][decimal]$rdr6['DataGB'] }
                $logGB    = if ($rdr6.IsDBNull($rdr6.GetOrdinal('LogGB')))   { 'N/A' } else { [string][decimal]$rdr6['LogGB'] }
                $state    = [string]$rdr6['State']
                $recovery = [string]$rdr6['RecoveryModel']
                $stColor  = if ($state -eq 'ONLINE') { 'green' } else { 'red' }
                $dbSizeHtml += "<tr><td>$(HtmlEncode $dbName)</td><td>$dataGB</td><td>$logGB</td>"
                $dbSizeHtml += "<td>$(StatusBadge $state $stColor)</td><td>$(HtmlEncode $recovery)</td></tr>"
            }
            $rdr6.Close()
            $dbSizeHtml += "</tbody></table></div>"
        } catch {
            $dbSizeHtml = "<p class='error'>Error querying database sizes: $(HtmlEncode $_.Exception.Message)</p>"
        }

        # ── Last full backup per DB ───────────────────────────────────────────
        $bakHtml = ''
        try {
            $sqlBak = "SELECT d.name AS DBName, " +
                "MAX(b.backup_finish_date) AS LastBackup, " +
                "DATEDIFF(hour, MAX(b.backup_finish_date), GETDATE()) AS HoursAgo " +
                "FROM sys.databases d " +
                "LEFT JOIN msdb.dbo.backupset b ON b.database_name=d.name AND b.type='D' " +
                "WHERE d.database_id>0 " +
                "GROUP BY d.name ORDER BY d.name"
            $cmd6b = $conn6.CreateCommand()
            $cmd6b.CommandText    = $sqlBak
            $cmd6b.CommandTimeout = 30
            $rdr6b = $cmd6b.ExecuteReader()
            $bakHtml  = "<div class='table-wrap'><table><thead><tr>"
            $bakHtml += "<th>Database</th><th>Last Full Backup</th><th>Age</th><th>Status</th>"
            $bakHtml += "</tr></thead><tbody>"
            while ($rdr6b.Read()) {
                $dbName   = [string]$rdr6b['DBName']
                $lastBak  = if ($rdr6b.IsDBNull($rdr6b.GetOrdinal('LastBackup'))) {
                    'Never'
                } else {
                    ([datetime]$rdr6b['LastBackup']).ToString('yyyy-MM-dd HH:mm')
                }
                $hoursAgo = if ($rdr6b.IsDBNull($rdr6b.GetOrdinal('HoursAgo'))) { 9999 } else { [int]$rdr6b['HoursAgo'] }
                $bakBadge = if ($lastBak -eq 'Never') {
                    $CriticalFindings.Add("Section 6 - Database '$dbName' has never been backed up")
                    StatusBadge 'Never Backed Up' 'red'
                } elseif ($hoursAgo -gt 48) {
                    $CriticalFindings.Add("Section 6 - Database '$dbName' last backup was $hoursAgo hours ago")
                    StatusBadge "${hoursAgo}h ago" 'red'
                } elseif ($hoursAgo -gt 24) {
                    StatusBadge "${hoursAgo}h ago" 'yellow'
                } else {
                    StatusBadge 'Recent' 'green'
                }
                $ageStr = if ($hoursAgo -lt 9999) { "${hoursAgo}h" } else { 'N/A' }
                $bakHtml += "<tr><td>$(HtmlEncode $dbName)</td><td>$(HtmlEncode $lastBak)</td>"
                $bakHtml += "<td>$(HtmlEncode $ageStr)</td><td>$bakBadge</td></tr>"
            }
            $rdr6b.Close()
            $bakHtml += "</tbody></table></div>"
        } catch {
            $bakHtml = "<p class='error'>Error querying backup history: $(HtmlEncode $_.Exception.Message)</p>"
        }

        $conn6.Close()
        $conn6.Dispose()

        $Sec6Html  = "<h4 style='margin:0 0 8px;'>Database Sizes</h4>$dbSizeHtml"
        $Sec6Html += "<h4 style='margin:14px 0 8px;'>Last Full Backup per Database</h4>$bakHtml"
    } catch {
        $Sec6Html = "<p class='error'>Error retrieving database details: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("Section 6 - Database details error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — IIS / WEB SERVER
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 7: IIS / Web Server..."
$Sec7Html = ''
try {
    $webAdminOk = $false
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $webAdminOk = $true
    } catch {}

    if (-not $webAdminOk) {
        $Sec7Html = "<p class='warn'>WebAdministration module not available. IIS may not be installed on this server.</p>"
    } else {
        # Application pools
        $poolsHtml = ''
        try {
            $pools = @(Get-ChildItem IIS:\AppPools -ErrorAction Stop)
            $poolsHtml  = "<div class='table-wrap'><table><thead><tr>"
            $poolsHtml += "<th>App Pool</th><th>State</th><th>.NET CLR</th><th>Pipeline</th><th>Identity</th>"
            $poolsHtml += "</tr></thead><tbody>"
            foreach ($pool in $pools) {
                $state    = $pool.State
                $stColor  = if ($state -eq 'Started') { 'green' } elseif ($state -eq 'Stopped') { 'red' } else { 'yellow' }
                $clrVer   = if ($pool.ManagedRuntimeVersion) { $pool.ManagedRuntimeVersion } else { 'No Managed Code' }
                $pipeline = $pool.ManagedPipelineMode
                $identity = $pool.processModel.userName
                if ([string]::IsNullOrEmpty($identity)) { $identity = $pool.processModel.identityType }
                $poolsHtml += "<tr><td>$(HtmlEncode $pool.Name)</td><td>$(StatusBadge $state $stColor)</td>"
                $poolsHtml += "<td>$(HtmlEncode $clrVer)</td><td>$(HtmlEncode $pipeline)</td><td>$(HtmlEncode $identity)</td></tr>"
                if ($state -ne 'Started') {
                    $CriticalFindings.Add("Section 7 - IIS App Pool '$($pool.Name)' is $state")
                }
            }
            $poolsHtml += "</tbody></table></div>"
        } catch {
            $poolsHtml = "<p class='error'>Error reading app pools: $(HtmlEncode $_.Exception.Message)</p>"
        }

        # Sites and bindings
        $sitesHtml = ''
        try {
            $sites = @(Get-ChildItem IIS:\Sites -ErrorAction Stop)
            $sitesHtml  = "<div class='table-wrap'><table><thead><tr>"
            $sitesHtml += "<th>Site Name</th><th>State</th><th>App Pool</th><th>Bindings</th>"
            $sitesHtml += "</tr></thead><tbody>"
            foreach ($site in $sites) {
                $sState   = $site.State
                $sColor   = if ($sState -eq 'Started') { 'green' } elseif ($sState -eq 'Stopped') { 'red' } else { 'yellow' }
                $bindings = ($site.Bindings.Collection | ForEach-Object { $_.bindingInformation }) -join ', '
                $sitesHtml += "<tr><td>$(HtmlEncode $site.Name)</td><td>$(StatusBadge $sState $sColor)</td>"
                $sitesHtml += "<td>$(HtmlEncode $site.ApplicationPool)</td><td>$(HtmlEncode $bindings)</td></tr>"
            }
            $sitesHtml += "</tbody></table></div>"
        } catch {
            $sitesHtml = "<p class='error'>Error reading IIS sites: $(HtmlEncode $_.Exception.Message)</p>"
        }

        # Worker processes
        $w3wpCount = 0
        try { $w3wpCount = @(Get-Process w3wp -ErrorAction SilentlyContinue).Count } catch {}
        $workerHtml = "<p style='margin-top:8px;'>Active w3wp worker processes: <strong>$w3wpCount</strong></p>"

        $Sec7Html  = "<h4 style='margin:0 0 8px;'>Application Pools</h4>$poolsHtml"
        $Sec7Html += "<h4 style='margin:14px 0 8px;'>Sites &amp; Bindings</h4>$sitesHtml"
        $Sec7Html += $workerHtml
    }
} catch {
    $Sec7Html = "<p class='error'>Error checking IIS: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 7 - IIS check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — NETWORK PORT CONNECTIVITY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 8: Network Port Connectivity..."
$Sec8Html = ''
try {
    # Build port list: ERP port + DB port (1433 for MSSQL, 1521 for Oracle) + extra ports
    $portTests = [System.Collections.Generic.List[hashtable]]::new()
    $portTests.Add(@{ Port = $ERPPort; Label = "ERP App Port ($ERPPort)" })

    if ($DBServiceName -match 'MSSQL|SQL') {
        $portTests.Add(@{ Port = 1433; Label = 'SQL Server (1433)' })
    } elseif ($DBServiceName -match 'Oracle|ORA') {
        $portTests.Add(@{ Port = 1521; Label = 'Oracle Listener (1521)' })
    } else {
        $portTests.Add(@{ Port = 1433; Label = 'SQL Server (1433)' })
    }

    foreach ($ep in $ExtraPorts) {
        $portTests.Add(@{ Port = $ep; Label = "Extra Port ($ep)" })
    }

    $Sec8Html  = "<div class='table-wrap'><table><thead><tr><th>Port</th><th>Description</th><th>Status</th></tr></thead><tbody>"
    foreach ($pt in $portTests) {
        $isOpen  = Test-TcpPort -ComputerName 'localhost' -Port $pt.Port
        $badge   = if ($isOpen) { StatusBadge 'Open' 'green' } else { StatusBadge 'Closed' 'red' }
        if (-not $isOpen -and $pt.Port -eq $ERPPort) {
            $CriticalFindings.Add("Section 8 - ERP port $ERPPort is not reachable on localhost")
        }
        $Sec8Html += "<tr><td><strong>$($pt.Port)</strong></td><td>$(HtmlEncode $pt.Label)</td><td>$badge</td></tr>"
    }
    $Sec8Html += "</tbody></table></div>"
    $Sec8Html += "<p class='info' style='margin-top:8px;'>Ports tested against localhost with a 2-second TCP connect timeout.</p>"
} catch {
    $Sec8Html = "<p class='error'>Error checking port connectivity: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 8 - Port connectivity check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 9 — WINDOWS FIREWALL
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 9: Windows Firewall..."
$Sec9Html = ''
try {
    $profiles = @(Get-NetFirewallProfile -ErrorAction Stop)
    $profileHtml  = "<div class='table-wrap'><table><thead><tr><th>Profile</th><th>Enabled</th><th>Default Inbound</th><th>Default Outbound</th></tr></thead><tbody>"
    $anyDisabled  = $false
    foreach ($pf in $profiles) {
        $enabled    = $pf.Enabled
        $enColor    = if ($enabled) { 'green' } else { 'red' }
        if (-not $enabled) { $anyDisabled = $true }
        $profileHtml += "<tr><td>$(HtmlEncode $pf.Name)</td><td>$(StatusBadge $(if ($enabled) {'Enabled'} else {'Disabled'}) $enColor)</td>"
        $profileHtml += "<td>$(HtmlEncode $pf.DefaultInboundAction)</td><td>$(HtmlEncode $pf.DefaultOutboundAction)</td></tr>"
    }
    $profileHtml += "</tbody></table></div>"
    if ($anyDisabled) {
        $CriticalFindings.Add("Section 9 - One or more Windows Firewall profiles are Disabled")
    }

    # Rules that allow ERP port (inbound)
    $erpRulesHtml = ''
    try {
        $erpRules = @(
            Get-NetFirewallRule -Direction Inbound -Enabled True -Action Allow -ErrorAction Stop |
            Where-Object {
                $portFilter = $_ | Get-NetFirewallPortFilter -ErrorAction SilentlyContinue
                $portFilter -and ($portFilter.LocalPort -eq $ERPPort -or $portFilter.LocalPort -eq 'Any')
            } |
            Select-Object -First 20
        )
        if ($erpRules.Count -gt 0) {
            $erpRulesHtml  = "<div class='table-wrap'><table><thead><tr><th>Rule Name</th><th>Profile</th><th>Protocol</th></tr></thead><tbody>"
            foreach ($rule in $erpRules) {
                $pf2 = $rule | Get-NetFirewallPortFilter -ErrorAction SilentlyContinue
                $proto = if ($null -ne $pf2) { $pf2.Protocol } else { 'Any' }
                $erpRulesHtml += "<tr><td>$(HtmlEncode $rule.DisplayName)</td><td>$(HtmlEncode $rule.Profile.ToString())</td><td>$(HtmlEncode $proto)</td></tr>"
            }
            $erpRulesHtml += "</tbody></table></div>"
        } else {
            $erpRulesHtml = "<p class='warn'>No inbound Allow rules found for port $ERPPort.</p>"
        }
    } catch {
        $erpRulesHtml = "<p class='error'>Error querying firewall rules: $(HtmlEncode $_.Exception.Message)</p>"
    }

    $Sec9Html  = "<h4 style='margin:0 0 8px;'>Firewall Profiles</h4>$profileHtml"
    $Sec9Html += "<h4 style='margin:14px 0 8px;'>Inbound Allow Rules for ERP Port $ERPPort</h4>$erpRulesHtml"
} catch {
    $Sec9Html = "<p class='error'>Error checking Windows Firewall: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 9 - Firewall check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 10 — SCHEDULED TASKS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 10: Scheduled Tasks..."
$Sec10Html = ''
try {
    $allTasks = @(Get-ScheduledTask -ErrorAction Stop |
        Where-Object { $_.TaskName -match $TaskNamePattern })

    if ($allTasks.Count -eq 0) {
        $Sec10Html = "<p class='info'>No scheduled tasks matching ERP/backup/sync/job/batch patterns found.</p>"
    } else {
        $Sec10Html  = "<div class='table-wrap'><table><thead><tr>"
        $Sec10Html += "<th>Task Name</th><th>State</th><th>Last Run</th><th>Last Result</th><th>Next Run</th>"
        $Sec10Html += "</tr></thead><tbody>"
        foreach ($task in $allTasks | Select-Object -First 30) {
            $info = $null
            try { $info = $task | Get-ScheduledTaskInfo -ErrorAction Stop } catch {}
            $state       = $task.State.ToString()
            $stColor     = if ($state -eq 'Ready' -or $state -eq 'Running') { 'green' } else { 'yellow' }
            $lastRun     = if ($null -ne $info -and $null -ne $info.LastRunTime -and
                              $info.LastRunTime -gt [datetime]'2000-01-01') {
                $info.LastRunTime.ToString('yyyy-MM-dd HH:mm')
            } else { 'Never' }
            $lastResult  = if ($null -ne $info) {
                $rc = $info.LastTaskResult
                if ($rc -eq 0) { StatusBadge 'Success (0)' 'green' }
                elseif ($rc -eq 267011) { StatusBadge 'Never Run' 'grey' }
                else { StatusBadge "Error (0x$($rc.ToString('X')))" 'red' }
            } else { StatusBadge 'N/A' 'grey' }
            $nextRun     = if ($null -ne $info -and $null -ne $info.NextRunTime -and
                              $info.NextRunTime -gt [datetime]'2000-01-01') {
                $info.NextRunTime.ToString('yyyy-MM-dd HH:mm')
            } else { 'Not scheduled' }
            $Sec10Html += "<tr><td>$(HtmlEncode $task.TaskName)</td><td>$(StatusBadge $state $stColor)</td>"
            $Sec10Html += "<td>$(HtmlEncode $lastRun)</td><td>$lastResult</td><td>$(HtmlEncode $nextRun)</td></tr>"
        }
        $Sec10Html += "</tbody></table></div>"
        $Sec10Html += "<p class='info' style='margin-top:8px;'>Showing up to 30 matching tasks. Filter pattern: <code>$(HtmlEncode $TaskNamePattern)</code></p>"
    }
} catch {
    $Sec10Html = "<p class='error'>Error checking scheduled tasks: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 10 - Scheduled task check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 11 — EVENT LOG CHECK
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 11: Event Log Check (last 24 hours)..."
$Sec11Html = ''
try {
    $since  = (Get-Date).AddHours(-24)
    $events = [System.Collections.Generic.List[object]]::new()

    foreach ($logName in @('Application', 'System')) {
        try {
            $filterHash = @{
                LogName   = $logName
                Level     = @(1, 2)    # 1=Critical, 2=Error
                StartTime = $since
            }
            $logEvents = @(Get-WinEvent -FilterHashtable $filterHash -MaxEvents 50 -ErrorAction Stop)
            foreach ($ev in $logEvents) { $events.Add($ev) }
        } catch [System.Exception] {
            # No matching events is normal — continue silently
            if ($_.Exception.Message -notmatch 'No events were found') {
                $events.Add([PSCustomObject]@{
                    TimeCreated = (Get-Date)
                    LevelDisplayName = 'Error'
                    Id          = 0
                    ProviderName = "EventLog/$logName"
                    Message     = "Error reading $logName log: $($_.Exception.Message)"
                })
            }
        }
    }

    # Sort by time descending and take top 15
    $topEvents = $events | Sort-Object TimeCreated -Descending | Select-Object -First 15

    if ($topEvents.Count -eq 0) {
        $Sec11Html = "<p style='color:#2ea043;'>&#x2714; No Critical or Error events in the last 24 hours. System event logs are clean.</p>"
    } else {
        $critEventCount = ($topEvents | Where-Object { $_.LevelDisplayName -eq 'Critical' }).Count
        if ($critEventCount -gt 0) {
            $CriticalFindings.Add("Section 11 - $critEventCount Critical event(s) in event logs (last 24 h)")
        }

        $Sec11Html  = "<p style='color:var(--muted);margin-bottom:8px;'>Showing up to 15 Critical/Error events from the last 24 hours (Application + System logs).</p>"
        $Sec11Html += "<div class='table-wrap'><table><thead><tr>"
        $Sec11Html += "<th>Time</th><th>Level</th><th>Event ID</th><th>Source</th><th>Message</th>"
        $Sec11Html += "</tr></thead><tbody>"
        foreach ($ev in $topEvents) {
            $timeStr  = if ($null -ne $ev.TimeCreated) { $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
            $level    = if ($null -ne $ev.LevelDisplayName) { $ev.LevelDisplayName } else { 'Unknown' }
            $lvlColor = if ($level -eq 'Critical') { 'red' } else { 'yellow' }
            $rowClass = if ($level -eq 'Critical') { 'row-critical' } else { 'row-error' }
            $source   = if ($null -ne $ev.ProviderName) { $ev.ProviderName } else { 'Unknown' }
            $msgRaw   = if ($null -ne $ev.Message) { $ev.Message } else { '' }
            $msgShort = if ($msgRaw.Length -gt 200) { $msgRaw.Substring(0, 200) + '...' } else { $msgRaw }
            $Sec11Html += "<tr class='$rowClass'>"
            $Sec11Html += "<td style='white-space:nowrap;'>$(HtmlEncode $timeStr)</td>"
            $Sec11Html += "<td>$(StatusBadge $level $lvlColor)</td>"
            $Sec11Html += "<td>$($ev.Id)</td>"
            $Sec11Html += "<td>$(HtmlEncode $source)</td>"
            $Sec11Html += "<td>$(HtmlEncode $msgShort)</td>"
            $Sec11Html += "</tr>"
        }
        $Sec11Html += "</tbody></table></div>"
    }
} catch {
    $Sec11Html = "<p class='error'>Error reading event logs: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 11 - Event log check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 12 — CERTIFICATE CHECK
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 12: Certificate Check..."
$Sec12Html = ''
try {
    $certStores = @('Cert:\LocalMachine\My', 'Cert:\LocalMachine\WebHosting')
    $allCerts   = [System.Collections.Generic.List[object]]::new()

    foreach ($storePath in $certStores) {
        try {
            $certs = @(Get-ChildItem $storePath -ErrorAction Stop)
            foreach ($c in $certs) { $allCerts.Add($c) }
        } catch {}
    }

    $now         = Get-Date
    $warnDate    = $now.AddDays($CertWarnDays)
    $critDate    = $now.AddDays($CertCritDays)
    $expiredCerts = @($allCerts | Where-Object { $_.NotAfter -lt $now })
    $critCerts    = @($allCerts | Where-Object { $_.NotAfter -ge $now -and $_.NotAfter -lt $critDate })
    $warnCerts    = @($allCerts | Where-Object { $_.NotAfter -ge $critDate -and $_.NotAfter -lt $warnDate })
    $okCerts      = @($allCerts | Where-Object { $_.NotAfter -ge $warnDate })

    foreach ($c in $expiredCerts) {
        $CriticalFindings.Add("Section 12 - Certificate EXPIRED: '$($c.Subject)' (expired $($c.NotAfter.ToString('yyyy-MM-dd')))")
    }
    foreach ($c in $critCerts) {
        $CriticalFindings.Add("Section 12 - Certificate expiring within $CertCritDays days: '$($c.Subject)' (expires $($c.NotAfter.ToString('yyyy-MM-dd')))")
    }

    if ($allCerts.Count -eq 0) {
        $Sec12Html = "<p class='info'>No certificates found in LocalMachine\My or LocalMachine\WebHosting stores.</p>"
    } else {
        $Sec12Html  = "<p style='margin-bottom:8px;'><strong>$($allCerts.Count)</strong> cert(s) scanned &nbsp;|&nbsp; "
        $Sec12Html += "<span style='color:#da3633;'>$($expiredCerts.Count) expired</span> &nbsp;|&nbsp; "
        $Sec12Html += "<span style='color:#da3633;'>$($critCerts.Count) expiring &lt;$CertCritDays d</span> &nbsp;|&nbsp; "
        $Sec12Html += "<span style='color:#d29922;'>$($warnCerts.Count) expiring &lt;$CertWarnDays d</span> &nbsp;|&nbsp; "
        $Sec12Html += "<span style='color:#2ea043;'>$($okCerts.Count) healthy</span></p>"

        $Sec12Html += "<div class='table-wrap'><table><thead><tr>"
        $Sec12Html += "<th>Subject</th><th>Thumbprint</th><th>Expires</th><th>Days Left</th><th>Status</th>"
        $Sec12Html += "</tr></thead><tbody>"
        foreach ($c in ($allCerts | Sort-Object NotAfter | Select-Object -First 30)) {
            $daysLeft  = [int]($c.NotAfter - $now).TotalDays
            $certStatus = if ($daysLeft -lt 0) { 'Expired' }
                          elseif ($daysLeft -lt $CertCritDays) { "Expiring ($daysLeft d)" }
                          elseif ($daysLeft -lt $CertWarnDays) { "Expiring ($daysLeft d)" }
                          else { 'Valid' }
            $certColor  = if ($daysLeft -lt 0) { 'red' }
                          elseif ($daysLeft -lt $CertCritDays) { 'red' }
                          elseif ($daysLeft -lt $CertWarnDays) { 'yellow' }
                          else { 'green' }
            $subject    = $c.Subject
            if ($subject.Length -gt 60) { $subject = $subject.Substring(0, 60) + '...' }
            $thumb      = $c.Thumbprint
            if ($thumb.Length -gt 16) { $thumb = $thumb.Substring(0, 16) + '...' }
            $Sec12Html += "<tr><td>$(HtmlEncode $subject)</td><td><code>$(HtmlEncode $thumb)</code></td>"
            $Sec12Html += "<td>$(HtmlEncode $c.NotAfter.ToString('yyyy-MM-dd'))</td>"
            $Sec12Html += "<td>$(HtmlEncode $daysLeft.ToString())</td><td>$(StatusBadge $certStatus $certColor)</td></tr>"
        }
        $Sec12Html += "</tbody></table></div>"
    }
} catch {
    $Sec12Html = "<p class='error'>Error checking certificates: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 12 - Certificate check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 13 — BACKUP VERIFICATION
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 13: Backup Verification..."
$Sec13Html = ''
try {
    $cutoff      = (Get-Date).AddHours(-$BackupMaxAgeHours)
    $foundFiles  = [System.Collections.Generic.List[object]]::new()
    $searchedPaths = [System.Collections.Generic.List[string]]::new()

    foreach ($path in $BackupSearchPaths) {
        if (Test-Path $path) {
            $searchedPaths.Add($path)
            try {
                $files = @(
                    Get-ChildItem -Path $path -Recurse -Include '*.bak','*.zip','*.tar','*.gz' -ErrorAction SilentlyContinue |
                    Where-Object { $_.LastWriteTime -ge $cutoff } |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 20
                )
                foreach ($f in $files) { $foundFiles.Add($f) }
            } catch {}
        }
    }

    $Sec13Html  = "<p style='margin-bottom:8px;'>"
    $Sec13Html += "Searched paths: <code>$($searchedPaths -join '</code>, <code>')</code><br>"
    $Sec13Html += "File extensions: *.bak, *.zip, *.tar, *.gz &nbsp;|&nbsp; Age filter: last $BackupMaxAgeHours hours"
    $Sec13Html += "</p>"

    if ($searchedPaths.Count -eq 0) {
        $Sec13Html += "<p class='warn'>&#x26A0; None of the configured backup paths exist. Verify <code>`$BackupSearchPaths</code>.</p>"
        $CriticalFindings.Add("Section 13 - No backup paths found on this server")
    } elseif ($foundFiles.Count -eq 0) {
        $Sec13Html += "<p class='error'>&#x26A0; No backup files found in the last $BackupMaxAgeHours hours in any configured path.</p>"
        $CriticalFindings.Add("Section 13 - No recent backup files found in last $BackupMaxAgeHours hours")
    } else {
        $newestFile   = $foundFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $newestAgeMin = [int]((Get-Date) - $newestFile.LastWriteTime).TotalMinutes
        $ageStr       = if ($newestAgeMin -lt 60) { "${newestAgeMin}m ago" } else { "$([int]($newestAgeMin/60))h ago" }

        $Sec13Html += "<p style='color:#2ea043;'>&#x2714; $($foundFiles.Count) backup file(s) found (newest: $ageStr)</p>"
        $Sec13Html += "<div class='table-wrap'><table><thead><tr>"
        $Sec13Html += "<th>File Name</th><th>Directory</th><th>Size (MB)</th><th>Last Modified</th>"
        $Sec13Html += "</tr></thead><tbody>"
        foreach ($f in ($foundFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 20)) {
            $sizeMB  = [math]::Round($f.Length / 1MB, 2)
            $Sec13Html += "<tr><td>$(HtmlEncode $f.Name)</td><td>$(HtmlEncode $f.Directory)</td>"
            $Sec13Html += "<td>$sizeMB</td><td>$(HtmlEncode $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))</td></tr>"
        }
        $Sec13Html += "</tbody></table></div>"
    }
} catch {
    $Sec13Html = "<p class='error'>Error checking backup files: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 13 - Backup verification error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SUMMARY KPI VALUES
# ═══════════════════════════════════════════════════════════════════════════════
$CritCount  = $CriticalFindings.Count
$ReportDate = (Get-Date).ToString('dddd, dd MMMM yyyy HH:mm:ss')
$EndTime    = Get-Date
$Duration   = ($EndTime - $StartTime).ToString('hh\:mm\:ss')

# Uptime KPI
$UptimeTileHtml = if ($UptimeStr -ne 'Unknown') { "<strong>$UptimeStr</strong>" } else { "<span style='color:var(--muted);'>Unknown</span>" }

# CPU KPI
$cpuTileColor   = if ($CpuPct -ge 90) { '#f85149' } elseif ($CpuPct -ge 75) { '#d29922' } else { '#3fb950' }
$CpuTileHtml    = "<span style='color:$cpuTileColor;font-weight:700;'>$CpuPct%</span>"

# RAM KPI
$ramTileColor   = if ($RamUsedPct -ge 90) { '#f85149' } elseif ($RamUsedPct -ge 75) { '#d29922' } else { '#3fb950' }
$RamTileHtml    = "<span style='color:$ramTileColor;font-weight:700;'>$RamUsedPct%</span>"

# Disk KPI
$diskTileColor  = if ($DiskWorstPct -ge $DiskCritPct) { '#f85149' } elseif ($DiskWorstPct -ge $DiskWarnPct) { '#d29922' } else { '#3fb950' }
$DiskTileHtml   = "<span style='color:$diskTileColor;font-weight:700;'>$DiskWorstPct%</span>"

# Services KPI
$svcTileColor   = if ($ServiceOkCount -lt $ServiceTotalCount) { '#f85149' } else { '#3fb950' }
$SvcTileHtml    = "<span style='color:$svcTileColor;font-weight:700;'>$ServiceOkCount / $ServiceTotalCount</span>"

# Overall
$OverallStatusBadge = if ($CritCount -gt 0) { StatusBadge 'CRITICAL' 'red' } else { StatusBadge 'HEALTHY' 'green' }

# Critical findings panel HTML
$CritFindingsHtml = ''
if ($CritCount -gt 0) {
    $cfRows = ($CriticalFindings | ForEach-Object { "<li>$(HtmlEncode $_)</li>" }) -join ''
    $CritFindingsHtml = "<div style='background:rgba(218,54,51,.12);border:1px solid #da3633;border-radius:8px;" +
        "padding:14px 18px;margin-bottom:16px;'>" +
        "<strong style='color:#f85149;'>&#x26A0; $CritCount Critical Finding(s) Detected</strong>" +
        "<ul style='margin:.6rem 0 0 1.2rem;color:#f85149;'>$cfRows</ul></div>"
}

# Logo & author link
$LogoHtml  = ''
if (-not [string]::IsNullOrWhiteSpace($CompanyLogoURL)) {
    $logoImg  = "<img src='$(HtmlEncode $CompanyLogoURL)' alt='Logo' style='max-height:56px;vertical-align:middle;'>"
    if (-not [string]::IsNullOrWhiteSpace($CompanyWebsite)) {
        $LogoHtml = "<a href='$(HtmlEncode $CompanyWebsite)' target='_blank'>$logoImg</a>"
    } else { $LogoHtml = $logoImg }
    $LogoHtml = "<div class='logo-wrap'>$LogoHtml</div>"
}
$AuthorLink = if (-not [string]::IsNullOrWhiteSpace($CompanyWebsite)) {
    "<a href='$(HtmlEncode $CompanyWebsite)' target='_blank' style='color:var(--link);'>$(HtmlEncode $AuthorName)</a>"
} else { HtmlEncode $AuthorName }

Write-Progress2 "Building HTML report..."

# ═══════════════════════════════════════════════════════════════════════════════
# BUILD FULL HTML
# ═══════════════════════════════════════════════════════════════════════════════
$HtmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ERP Health Check - $(HtmlEncode $ERPName) - $(HtmlEncode $ServerHostname)</title>
<style>
/* ── CSS VARIABLES ── */
:root {
  --bg:       #0d1117;
  --card:     #161b22;
  --border:   #30363d;
  --text:     #c9d1d9;
  --muted:    #8b949e;
  --link:     #2ea043;
  --head-bg:  #010409;
  --th-bg:    #21262d;
  --tr-alt:   #1c2128;
  --pre-bg:   #0d1117;
  --green:    #2ea043;
  --yellow:   #d29922;
  --red:      #da3633;
  --blue:     #1f6feb;
}
[data-theme="light"] {
  --bg:       #ffffff;
  --card:     #f6f8fa;
  --border:   #d0d7de;
  --text:     #24292f;
  --muted:    #57606a;
  --link:     #1a7f37;
  --head-bg:  #f6f8fa;
  --th-bg:    #eaeef2;
  --tr-alt:   #f6f8fa;
  --pre-bg:   #f6f8fa;
  --green:    #1a7f37;
  --yellow:   #9a6700;
  --red:      #cf222e;
  --blue:     #0969da;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body  { font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
        background: var(--bg); color: var(--text); font-size: 14px; line-height: 1.5; }
a { color: var(--link); }
code { font-family: Consolas,'SFMono-Regular',monospace; font-size: .85em;
       background: var(--th-bg); padding: 1px 5px; border-radius: 4px; }

/* ── LAYOUT ── */
.page-wrap { max-width: 1280px; margin: 0 auto; padding: 0 20px 40px; }
.header    { background: var(--head-bg); border-bottom: 1px solid var(--border);
             padding: 16px 24px; display: flex; align-items: center;
             justify-content: space-between; flex-wrap: wrap; gap: 12px; }
.header-left  { display: flex; align-items: center; gap: 16px; }
.header-title h1 { font-size: 1.25rem; font-weight: 700; color: var(--text);
                   display: flex; align-items: center; gap: 8px; }
.header-title p  { font-size: .8rem; color: var(--muted); margin-top: 2px; }
.logo-wrap  { max-height: 60px; }
.theme-toggle { cursor: pointer; background: var(--card); border: 1px solid var(--border);
                color: var(--text); border-radius: 20px; padding: 6px 14px;
                font-size: 12px; display: flex; align-items: center; gap: 6px; }
.theme-toggle:hover { background: var(--th-bg); }

/* ── SUMMARY BAR ── */
.summary-bar { display: flex; flex-wrap: wrap; gap: 12px;
               background: var(--card); border: 1px solid var(--border);
               border-radius: 8px; padding: 16px 20px; margin: 20px 0; }
.stat-card   { flex: 1 1 130px; text-align: center; }
.stat-count  { font-size: 2rem; font-weight: 700; line-height: 1; }
.stat-label  { font-size: .72rem; color: var(--muted); margin-top: 4px; }
.c-green  { color: #3fb950; }
.c-yellow { color: #d29922; }
.c-red    { color: #f85149; }
.c-blue   { color: #58a6ff; }
.c-muted  { color: var(--muted); }

/* ── SECTION CARDS ── */
.section-card { background: var(--card); border: 1px solid var(--border);
                border-radius: 8px; margin-bottom: 16px; overflow: hidden; }
.section-summary { display: flex; align-items: center; gap: 10px; cursor: pointer;
                   padding: 14px 18px; list-style: none; user-select: none; }
.section-summary::-webkit-details-marker { display: none; }
.section-summary::marker { display: none; }
.sec-arrow { font-size: 10px; color: var(--muted); display: inline-block;
             transition: transform .2s; flex-shrink: 0; line-height: 1; }
details[open] > .section-summary .sec-arrow { transform: rotate(90deg); }
.section-summary:hover { background: var(--th-bg); }
.sec-num   { background: var(--green); color: #fff; font-size: .7rem; font-weight: 700;
             width: 22px; height: 22px; border-radius: 50%; display: flex;
             align-items: center; justify-content: center; flex-shrink: 0; }
.sec-title { font-weight: 600; font-size: .95rem; flex: 1; }
.section-body { padding: 16px 18px; border-top: 1px solid var(--border); }
h4 { font-size: .9rem; font-weight: 600; }

/* ── TABLES ── */
.table-wrap { overflow-x: auto; border-radius: 6px; border: 1px solid var(--border); }
table       { width: 100%; border-collapse: collapse; font-size: 13px; }
thead tr    { background: var(--th-bg); position: sticky; top: 0; z-index: 1; }
th          { padding: 8px 12px; text-align: left; font-weight: 600;
              border-bottom: 1px solid var(--border); white-space: nowrap; }
td          { padding: 7px 12px; border-bottom: 1px solid var(--border); vertical-align: top; }
tbody tr:nth-child(even) { background: var(--tr-alt); }
tbody tr:hover { background: var(--th-bg); }
.kv-table   { width: 100%; border-collapse: collapse; font-size: 13px; }
.kv-table td { padding: 7px 12px; border-bottom: 1px solid var(--border); }
.td-label   { font-weight: 600; white-space: nowrap; width: 200px; color: var(--muted); }

/* ── DISK BARS ── */
.disk-container  { display: flex; flex-direction: column; gap: 14px; }
.disk-row        { }
.disk-label      { font-size: .85rem; margin-bottom: 4px; font-weight: 600; }
.disk-bar-outer  { height: 18px; background: var(--th-bg); border-radius: 9px;
                   overflow: hidden; margin-bottom: 4px; border: 1px solid var(--border); }
.disk-bar-inner  { height: 100%; border-radius: 9px; transition: width .4s ease; }
.disk-stat       { font-size: .8rem; color: var(--muted); }

/* ── EVENT ROW COLORS ── */
.row-critical { background: rgba(218,54,51,.15) !important; }
.row-error    { background: rgba(218,54,51,.08) !important; }
.row-warning  { background: rgba(210,153,34,.12) !important; }

/* ── BADGES ── */
.badge { display: inline-block; font-size: .7rem; font-weight: 600; padding: 2px 8px;
         border-radius: 20px; color: #fff; white-space: nowrap; }

/* ── MESSAGE CLASSES ── */
.error { color: #f85149; padding: 8px 12px; background: rgba(248,81,73,.1);
         border-left: 3px solid #f85149; border-radius: 4px; }
.warn  { color: #d29922; padding: 8px 12px; background: rgba(210,153,34,.1);
         border-left: 3px solid #d29922; border-radius: 4px; }
.info  { color: var(--muted); }

/* ── FOOTER ── */
.footer { border-top: 1px solid var(--border); padding: 20px 0;
          margin-top: 24px; color: var(--muted); font-size: .8rem;
          display: flex; justify-content: space-between; flex-wrap: wrap; gap: 8px; }

/* ── EVENT TABLE MESSAGE CELL ── */
.event-table td:nth-child(5) {
  max-width: 360px; word-break: break-word; white-space: normal;
}

/* ── RESPONSIVE ── */
@media (max-width: 768px) {
  .summary-bar { gap: 8px; }
  .stat-card   { flex: 1 1 100px; }
  .stat-count  { font-size: 1.5rem; }
  .header      { flex-direction: column; align-items: flex-start; }
}
</style>
</head>
<body data-theme="dark">

<!-- HEADER -->
<div class="header">
  <div class="header-left">
    $LogoHtml
    <div class="header-title">
      <h1>&#x1F4CA; ERP Server Health Check</h1>
      <p>$(HtmlEncode $ERPName) &nbsp;|&nbsp; $(HtmlEncode $ServerHostname) &nbsp;|&nbsp; Generated: $(HtmlEncode $ReportDate)</p>
    </div>
  </div>
  <button class="theme-toggle" onclick="toggleTheme()" title="Toggle Dark/Light mode">
    <span id="theme-icon">&#x2600;&#xFE0F;</span> Toggle Theme
  </button>
</div>

<!-- SUMMARY BAR -->
<div class="page-wrap">
<div class="summary-bar">
  <div class="stat-card">
    <div class="stat-count" style="font-size:.9rem;padding-top:10px;">$UptimeTileHtml</div>
    <div class="stat-label">System Uptime</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.4rem;">$CpuTileHtml</div>
    <div class="stat-label">CPU Load</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.4rem;">$RamTileHtml</div>
    <div class="stat-label">RAM Used</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.4rem;">$DiskTileHtml</div>
    <div class="stat-label">Disk Used (worst)</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.1rem;padding-top:10px;">$SvcTileHtml</div>
    <div class="stat-label">Services Running</div>
  </div>
  <div class="stat-card">
    <div class="stat-count $(if ($CritCount -gt 0) { 'c-red' } else { 'c-green' })">$CritCount</div>
    <div class="stat-label">Critical Findings</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:.9rem;padding-top:10px;">$OverallStatusBadge</div>
    <div class="stat-label">Overall Status</div>
  </div>
</div>

<!-- CRITICAL FINDINGS PANEL -->
$CritFindingsHtml

<!-- 13 SECTIONS -->
$(BuildSection 1  'Server Baseline'                $Sec1Html  ($Sec1Html  -match 'class=.error')  $true)
$(BuildSection 2  'CPU & Memory Performance'       $Sec2Html  ($Sec2Html  -match 'class=.error')  $true)
$(BuildSection 3  'Disk Health'                    $Sec3Html  ($Sec3Html  -match 'class=.error')  $true)
$(BuildSection 4  'ERP Application Services'       $Sec4Html  ($Sec4Html  -match 'class=.error')  $true)
$(BuildSection 5  'Database Service & Connectivity' $Sec5Html ($Sec5Html  -match 'class=.error')  $true)
$(BuildSection 6  'Database Details'               $Sec6Html  ($Sec6Html  -match 'class=.error')  $false)
$(BuildSection 7  'IIS / Web Server'               $Sec7Html  ($Sec7Html  -match 'class=.error')  $false)
$(BuildSection 8  'Network Port Connectivity'      $Sec8Html  ($Sec8Html  -match 'class=.error')  $false)
$(BuildSection 9  'Windows Firewall'               $Sec9Html  ($Sec9Html  -match 'class=.error')  $false)
$(BuildSection 10 'Scheduled Tasks'                $Sec10Html ($Sec10Html -match 'class=.error')  $false)
$(BuildSection 11 'Event Log Check'                $Sec11Html ($Sec11Html -match 'class=.error')  $false)
$(BuildSection 12 'Certificate Check'              $Sec12Html ($Sec12Html -match 'class=.error')  $false)
$(BuildSection 13 'Backup Verification'            $Sec13Html ($Sec13Html -match 'class=.error')  $false)

<!-- FOOTER -->
<div class="footer">
  <div>
    <strong>ERP Server Health Check v$ScriptVersion</strong> &nbsp;|&nbsp;
    ERP: $(HtmlEncode $ERPName) &nbsp;|&nbsp;
    Generated: $(HtmlEncode $ReportDate) &nbsp;|&nbsp;
    Duration: $(HtmlEncode $Duration)
  </div>
  <div>Created by $AuthorLink</div>
</div>

</div><!-- end page-wrap -->

<script>
function toggleTheme() {
  var body = document.body;
  var icon = document.getElementById('theme-icon');
  if (body.getAttribute('data-theme') === 'dark') {
    body.setAttribute('data-theme', 'light');
    icon.textContent = '\uD83C\uDF19';
  } else {
    body.setAttribute('data-theme', 'dark');
    icon.textContent = '\u2600\uFE0F';
  }
}
</script>
</body>
</html>
"@

# ═══════════════════════════════════════════════════════════════════════════════
# WRITE HTML REPORT FILE
# ═══════════════════════════════════════════════════════════════════════════════
try {
    [System.IO.File]::WriteAllText($ReportFile, $HtmlReport, [System.Text.Encoding]::UTF8)
    Write-Host ""
    Write-Host "  [OK] Report saved to: $ReportFile" -ForegroundColor Green
} catch {
    Write-Warning "Failed to write report: $_"
}

# ═══════════════════════════════════════════════════════════════════════════════
# STATUS SUMMARY FILE  (_HEALTHY.txt or _CRITICAL.txt)
# Read by ERP_HealthCheck_EmailAlert.ps1 to trigger email notifications.
# ═══════════════════════════════════════════════════════════════════════════════
$isCritical   = $CriticalFindings.Count -gt 0
$statusSuffix = if ($isCritical) { '_CRITICAL' } else { '_HEALTHY' }
$StatusFile   = Join-Path $ReportsDir ($ReportStamp + $statusSuffix + '.txt')
$separator    = '=' * 70

if ($isCritical) {
    $statusContent  = "$separator`r`n"
    $statusContent += " ERP SERVER HEALTH CHECK  -  *** CRITICAL ALERT ***`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " Status       : CRITICAL`r`n"
    $statusContent += " ERP System   : $ERPName`r`n"
    $statusContent += " Server       : $ServerHostname`r`n"
    $statusContent += " Generated    : $ReportDate`r`n"
    $statusContent += " Duration     : $Duration`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += "`r`n CRITICAL FINDINGS ($($CriticalFindings.Count)):`r`n`r`n"
    $statusContent += (@($CriticalFindings) | ForEach-Object { "  [!] $_" }) -join "`r`n"
    $statusContent += "`r`n`r`n$separator`r`n"
    $statusContent += " Full HTML report : $ReportFile`r`n"
    $statusContent += " Status file      : $StatusFile`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " ERP Health Check v$ScriptVersion  by $AuthorName`r`n"
    $statusContent += "$separator`r`n"
} else {
    $statusContent  = "$separator`r`n"
    $statusContent += " ERP SERVER HEALTH CHECK  -  HEALTHY STATE`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " Status       : HEALTHY`r`n"
    $statusContent += " ERP System   : $ERPName`r`n"
    $statusContent += " Server       : $ServerHostname`r`n"
    $statusContent += " Generated    : $ReportDate`r`n"
    $statusContent += " Duration     : $Duration`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += "`r`n No critical findings, errors, or warnings were detected.`r`n"
    $statusContent += " Your ERP server is in a healthy state.`r`n"
    $statusContent += "`r`n$separator`r`n"
    $statusContent += " Full HTML report : $ReportFile`r`n"
    $statusContent += " Status file      : $StatusFile`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " ERP Health Check v$ScriptVersion  by $AuthorName`r`n"
    $statusContent += "$separator`r`n"
}

if ($EnableStatusFile) {
    try {
        [System.IO.File]::WriteAllText($StatusFile, $statusContent, [System.Text.Encoding]::UTF8)
        $statusColor = if ($isCritical) { 'Red' } else { 'Green' }
        Write-Host "  [OK] Status file saved to: $StatusFile" -ForegroundColor $statusColor
    } catch {
        Write-Warning "Failed to write status file: $_"
    }
}

# ── CONSOLE SUMMARY ───────────────────────────────────────────────────────────
Write-Host ""
Write-Host "===============================================================" -ForegroundColor DarkGreen
Write-Host "  ERP Health Check complete.  Duration: $Duration" -ForegroundColor DarkGreen
Write-Host "  ERP System : $ERPName" -ForegroundColor DarkGreen
Write-Host "  Report     : $ReportFile" -ForegroundColor Yellow
if ($EnableStatusFile) {
    $statusLabel = if ($isCritical) { 'Status (CRITICAL)' } else { 'Status (HEALTHY)' }
    Write-Host "  $statusLabel : $StatusFile" -ForegroundColor $(if ($isCritical) { 'Red' } else { 'Green' })
    Write-Host ""
    Write-Host "  To send email alerts, run: .\ERP_HealthCheck_EmailAlert.ps1  (configure SMTP settings inside first)" -ForegroundColor Cyan
}
Write-Host "===============================================================" -ForegroundColor DarkGreen
Write-Host ""
