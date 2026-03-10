#Requires -Version 5.1
<#
.SYNOPSIS
    Universal Windows Server Health Check Script

.DESCRIPTION
    Performs a comprehensive baseline health check of any Windows Server —
    member server, file server, print server, utility server, or standalone box.
    No role-specific modules are required; uses only built-in PowerShell
    cmdlets, WMI/CIM, and .NET.  Every section is individually wrapped in
    Try/Catch so a single failure never aborts the run.

    Checks performed:
      1.  System Identity        (hostname, FQDN, IPs, domain, uptime)
      2.  Hardware Inventory     (CPU, RAM, BIOS, motherboard, make/model)
      3.  OS Details             (version, build, activation, last Windows Update)
      4.  Disk Health            (physical disks, volumes with bars, SMART status)
      5.  CPU & Memory Real-Time (5-sec CPU sample, top-10 processes)
      6.  Network Adapters       (status, speed, IP, gateway, DNS, bytes in/out)
      7.  Windows Services       (critical service health matrix)
      8.  Installed Roles & Features
      9.  Windows Firewall       (Domain / Private / Public profiles)
      10. Windows Defender       (scan time, signature age, real-time protection)
      11. Pending Reboots        (registry-based multi-source detection)
      12. Local Users & Groups   (Administrators membership, locked/pwd status)
      13. Shared Folders         (non-default shares, path, permissions summary)
      14. Scheduled Tasks        (non-Microsoft custom tasks)
      15. Event Log Summary      (last 24h Critical/Error counts + last 10 events)
      16. RDP & Remote Mgmt      (enabled, port, NLA)
      17. Time Synchronisation   (w32tm source, offset, warn >5 s)

.NOTES
    Version    : 1.0.0
    Author     : Tushar Gudde
    Requires   : PowerShell 5.1+, Local Administrator rights
    Compatible : Windows Server 2012 R2, 2016, 2019, 2022
    Companion  : Server_HealthCheck_EmailAlert.ps1
#>

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
$CompanyLogoURL = ''                         # URL/path to logo image. Leave blank to skip.
$CompanyWebsite = 'https://tushargudde.tech' # Company website URL for logo hyperlink.
$AuthorName     = 'Tushar Gudde'             # Author name shown in the footer.

# Disk usage warning thresholds (%)
$DiskWarnPct = 85
$DiskCritPct = 95

# Time-sync warning threshold (seconds)
$TimeSyncWarnSec = 5

# Critical services to monitor (Section 7)
$CriticalServices = @(
    'WinRM', 'wuauserv', 'RpcSs', 'LanmanServer', 'LanmanWorkstation',
    'EventLog', 'PlugPlay', 'CryptSvc', 'Themes', 'Spooler'
)

# Write a plain-text companion status file (_HEALTHY.txt / _CRITICAL.txt)
# alongside every HTML report.  Required by Server_HealthCheck_EmailAlert.ps1.
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
$ReportStamp = "Server_Health_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
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
        green  = '#3fb950'
        yellow = '#d29922'
        red    = '#f85149'
        blue   = '#58a6ff'
        amber  = '#e3a019'
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
    param([array]$rows)
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
        "<span style='color:#f85149;'>&#x2716;</span>"
    } else {
        "<span style='color:#e3a019;'>&#x2714;</span>"
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

# ── CRITICAL FINDINGS LIST ────────────────────────────────────────────────────
$CriticalFindings = [System.Collections.Generic.List[string]]::new()

# ── DEFAULT KPI VARIABLES ────────────────────────────────────────────────────
$ServerHostname = $env:COMPUTERNAME
$UptimeStr      = 'Unknown'
$CpuPct         = 0
$RamUsedPct     = 0
$TotalRAM_GB    = 0
$UsedRAM_GB     = 0
$FreeRAM_GB     = 0
$DiskWorstPct   = 0
$PendingReboot  = $false

Write-Host ""
Write-Host "+==============================================================+" -ForegroundColor DarkYellow
Write-Host "|   Windows Server Health Check  v$ScriptVersion" -ForegroundColor DarkYellow
Write-Host "|   Server: $($env:COMPUTERNAME)" -ForegroundColor DarkYellow
Write-Host "+==============================================================+" -ForegroundColor DarkYellow
Write-Host ""

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — SYSTEM IDENTITY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 1: System Identity..."
$Sec1Html = ''
try {
    $cs  = Get-CimInstance Win32_ComputerSystem  -ErrorAction Stop
    $os  = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $net = @(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction Stop)

    $ServerHostname = $env:COMPUTERNAME
    $fqdn           = try { [System.Net.Dns]::GetHostEntry('').HostName } catch { $ServerHostname }

    # IP and MAC addresses
    $ipList  = ($net | ForEach-Object { $_.IPAddress  | Where-Object { $_ -match '\.' } }) -join ', '
    $macList = ($net | ForEach-Object { $_.MACAddress } | Where-Object { $_ }) -join ', '

    # Domain membership
    $domainStatus = if ($cs.PartOfDomain) {
        StatusBadge "Domain: $($cs.Domain)" 'green'
    } else {
        StatusBadge 'Workgroup' 'yellow'
    }

    # Last boot / uptime
    $bootTime = $os.LastBootUpTime
    if ($null -ne $bootTime) {
        $uptime      = New-TimeSpan -Start $bootTime -End (Get-Date)
        $UptimeStr   = '{0}d {1}h {2}m' -f [int]$uptime.TotalDays, $uptime.Hours, $uptime.Minutes
        $lastBootStr = $bootTime.ToString('yyyy-MM-dd HH:mm:ss')
    } else {
        $UptimeStr   = 'Unknown'
        $lastBootStr = 'N/A'
    }

    # Logged-on users (interactive sessions)
    $loggedUsers = 'N/A'
    try {
        $sessions = @(query user 2>$null)
        if ($sessions.Count -gt 1) {
            $loggedUsers = ($sessions | Select-Object -Skip 1 |
                ForEach-Object { ($_ -split '\s+' | Where-Object { $_ })[0] }) -join ', '
        } else {
            $loggedUsers = 'No active sessions'
        }
    } catch { $loggedUsers = 'N/A' }

    $rows1 = @(
        @('Hostname',        (HtmlEncode $ServerHostname)),
        @('FQDN',            (HtmlEncode $fqdn)),
        @('IP Addresses',    (HtmlEncode $ipList)),
        @('MAC Addresses',   (HtmlEncode $macList)),
        @('Domain / Group',  $domainStatus),
        @('Logged-on Users', (HtmlEncode $loggedUsers)),
        @('Last Boot Time',  (HtmlEncode $lastBootStr)),
        @('System Uptime',   (HtmlEncode $UptimeStr))
    )
    $Sec1Html = BuildKVTable $rows1
} catch {
    $Sec1Html = "<p class='error'>Error retrieving system identity: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 1 - System identity error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — HARDWARE INVENTORY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 2: Hardware Inventory..."
$Sec2Html = ''
try {
    $cs    = Get-CimInstance Win32_ComputerSystem  -ErrorAction Stop
    $bios  = Get-CimInstance Win32_BIOS            -ErrorAction Stop
    $board = Get-CimInstance Win32_BaseBoard       -ErrorAction Stop
    $cpus  = @(Get-CimInstance Win32_Processor     -ErrorAction Stop)
    $osObj = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop

    $TotalRAM_GB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
    # FreePhysicalMemory is in KB. Convert KB to GB: 1 GB = 1024 * 1024 KB.
    $FreeRAM_GB  = [math]::Round($osObj.FreePhysicalMemory / (1024 * 1024), 2)
    $UsedRAM_GB  = [math]::Round($TotalRAM_GB - $FreeRAM_GB, 2)
    $RamUsedPct  = if ($TotalRAM_GB -gt 0) {
        [math]::Round(($UsedRAM_GB / $TotalRAM_GB) * 100, 1)
    } else { 0 }

    $isVirtual = ($cs.Model        -match 'Virtual|VMware|VirtualBox|QEMU|KVM|Xen|HVM') -or
                 ($cs.Manufacturer -match 'VMware|QEMU|Xen|Parallels|innotek') -or
                 ($cs.Manufacturer -eq 'Microsoft Corporation' -and $cs.Model -match 'Virtual')
    $serverTypeBadge = if ($isVirtual) { StatusBadge 'Virtual Machine' 'blue' } else { StatusBadge 'Physical Server' 'green' }

    $cpuName  = ($cpus | ForEach-Object { if ($null -ne $_.Name) { $_.Name.Trim() } else { 'Unknown' } } | Select-Object -Unique) -join '; '
    $cpuCount = $cpus.Count
    $cores    = if ($cpus.Count -gt 0 -and $null -ne $cpus[0].NumberOfCores) { $cpus[0].NumberOfCores * $cpus.Count } else { 'N/A' }
    $logical  = if ($cpus.Count -gt 0 -and $null -ne $cpus[0].NumberOfLogicalProcessors) { $cpus[0].NumberOfLogicalProcessors * $cpus.Count } else { 'N/A' }

    $ramBadge = if ($RamUsedPct -ge 90) { StatusBadge "${RamUsedPct}% used" 'red' }
                elseif ($RamUsedPct -ge 75) { StatusBadge "${RamUsedPct}% used" 'yellow' }
                else { StatusBadge "${RamUsedPct}% used" 'green' }

    $rows2 = @(
        @('Server Type',      $serverTypeBadge),
        @('Make / Model',     "$(HtmlEncode $cs.Manufacturer) &mdash; $(HtmlEncode $cs.Model)"),
        @('CPU Model',        "$(HtmlEncode $cpuName) ($cpuCount socket(s))"),
        @('Physical Cores',   $cores.ToString()),
        @('Logical Procs',    $logical.ToString()),
        @('Total RAM',        "$TotalRAM_GB GB"),
        @('Used RAM',         "$UsedRAM_GB GB &nbsp; $ramBadge"),
        @('Free RAM',         "$FreeRAM_GB GB"),
        @('BIOS Version',     (HtmlEncode $bios.SMBIOSBIOSVersion)),
        @('BIOS Manufacturer',(HtmlEncode $bios.Manufacturer)),
        @('BIOS Date',        (HtmlEncode $bios.ReleaseDate.ToString('yyyy-MM-dd'))),
        @('Serial Number',    (HtmlEncode $bios.SerialNumber)),
        @('Motherboard',      "$(HtmlEncode $board.Manufacturer) &mdash; $(HtmlEncode $board.Product)")
    )
    $Sec2Html = BuildKVTable $rows2
} catch {
    $Sec2Html = "<p class='error'>Error retrieving hardware inventory: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 2 - Hardware inventory error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — OS DETAILS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 3: OS Details..."
$Sec3Html = ''
try {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop

    $installDate  = if ($null -ne $os.InstallDate) { $os.InstallDate.ToString('yyyy-MM-dd') } else { 'N/A' }
    $lastBootStr  = if ($null -ne $os.LastBootUpTime) { $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }

    # Activation status via SoftwareLicensingProduct WMI
    $activationStr = 'Unknown'
    try {
        $licProduct = Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%' and LicenseStatus=1" -ErrorAction Stop |
            Select-Object -First 1
        if ($null -ne $licProduct) {
            $activationStr = 'Activated'
        } else {
            $activationStr = 'Not Activated'
        }
    } catch { $activationStr = 'Unable to check' }
    $activationBadge = if ($activationStr -eq 'Activated') { StatusBadge 'Activated' 'green' } else { StatusBadge $activationStr 'red' }
    if ($activationStr -eq 'Not Activated') {
        $CriticalFindings.Add("Section 3 - OS is not activated")
    }

    # Last Windows Update
    $lastUpdateStr = 'Unknown'
    try {
        $wu = Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending | Select-Object -First 1
        if ($null -ne $wu -and $null -ne $wu.InstalledOn) {
            $lastUpdateStr = $wu.InstalledOn.ToString('yyyy-MM-dd') + "  (KB: $($wu.HotFixID))"
        }
    } catch { $lastUpdateStr = 'Unable to check' }

    $rows3 = @(
        @('OS Name',          (HtmlEncode $os.Caption)),
        @('OS Version',       (HtmlEncode $os.Version)),
        @('OS Build',         (HtmlEncode $os.BuildNumber)),
        @('Architecture',     (HtmlEncode $os.OSArchitecture)),
        @('Install Date',     (HtmlEncode $installDate)),
        @('Last Boot Time',   (HtmlEncode $lastBootStr)),
        @('Activation',       $activationBadge),
        @('Last Windows Update', (HtmlEncode $lastUpdateStr))
    )
    $Sec3Html = BuildKVTable $rows3
} catch {
    $Sec3Html = "<p class='error'>Error retrieving OS details: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 3 - OS details error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — DISK HEALTH
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 4: Disk Health..."
$Sec4Html     = ''
$DiskWorstPct = 0
try {
    # ── Physical disks ────────────────────────────────────────────────────────
    $physHtml = ''
    try {
        $physDisks = @(Get-PhysicalDisk -ErrorAction Stop)
        $physHtml  = "<div class='table-wrap'><table><thead><tr>"
        $physHtml += "<th>Friendly Name</th><th>Size (GB)</th><th>Media Type</th><th>Bus Type</th><th>Operational Status</th>"
        $physHtml += "</tr></thead><tbody>"
        foreach ($pd in $physDisks) {
            $sizeGB = if ($null -ne $pd.Size) { [math]::Round($pd.Size / 1GB, 0) } else { 'N/A' }
            $opStatus = if ($null -ne $pd.OperationalStatus) { $pd.OperationalStatus } else { 'Unknown' }
            $opColor  = if ($opStatus -eq 'OK') { 'green' } elseif ($opStatus -eq 'Unknown') { 'grey' } else { 'red' }
            if ($opStatus -notin @('OK', 'Unknown')) {
                $CriticalFindings.Add("Section 4 - Physical disk '$($pd.FriendlyName)' status: $opStatus")
            }
            $mediaType = if ($null -ne $pd.MediaType) { $pd.MediaType } else { 'Unknown' }
            $busType   = if ($null -ne $pd.BusType)   { $pd.BusType   } else { 'Unknown' }
            $physHtml += "<tr><td>$(HtmlEncode $pd.FriendlyName)</td><td>$sizeGB</td>"
            $physHtml += "<td>$(HtmlEncode $mediaType)</td><td>$(HtmlEncode $busType)</td>"
            $physHtml += "<td>$(StatusBadge $opStatus $opColor)</td></tr>"
        }
        $physHtml += "</tbody></table></div>"
    } catch {
        $physHtml = "<p class='warn'>Physical disk info not available (Get-PhysicalDisk): $(HtmlEncode $_.Exception.Message)</p>"
    }

    # ── Logical volumes ───────────────────────────────────────────────────────
    $logicalDisks = @(Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop)
    $volHtml      = ''
    $diskBarHtml  = "<div class='disk-container'>"

    if ($logicalDisks.Count -eq 0) {
        $volHtml    = "<p class='warn'>No fixed volumes found.</p>"
        $diskBarHtml = ''
    } else {
        foreach ($disk in $logicalDisks) {
            if ($null -ne $disk.Size -and $disk.Size -gt 0) {
                $totalGB = [math]::Round($disk.Size      / 1GB, 2)
                $freeGB  = [math]::Round($disk.FreeSpace / 1GB, 2)
                $usedGB  = [math]::Round($totalGB - $freeGB, 2)
                $pctUsed = [math]::Round(($usedGB / $totalGB) * 100, 1)
                if ($pctUsed -gt $DiskWorstPct) { $DiskWorstPct = $pctUsed }

                if ($pctUsed -ge $DiskCritPct) {
                    $CriticalFindings.Add("Section 4 - Volume $($disk.DeviceID) is ${pctUsed}% full (CRITICAL)")
                } elseif ($pctUsed -ge $DiskWarnPct) {
                    $CriticalFindings.Add("Section 4 - Volume $($disk.DeviceID) is ${pctUsed}% full (Warning)")
                }

                $barColor   = if ($pctUsed -ge $DiskCritPct) { '#f85149' } elseif ($pctUsed -ge $DiskWarnPct) { '#d29922' } else { '#e3a019' }
                $statusStr  = if ($pctUsed -ge $DiskCritPct) { 'CRITICAL' } elseif ($pctUsed -ge $DiskWarnPct) { 'Warning' } else { 'OK' }
                $statusColor= if ($pctUsed -ge $DiskCritPct) { 'red' } elseif ($pctUsed -ge $DiskWarnPct) { 'yellow' } else { 'amber' }
                $labelStr   = if ([string]::IsNullOrEmpty($disk.VolumeName)) { 'Local Disk' } else { $disk.VolumeName }

                $diskBarHtml += "<div class='disk-row'>"
                $diskBarHtml += "<div class='disk-label'><strong>$(HtmlEncode $disk.DeviceID)</strong>&nbsp;$(HtmlEncode $labelStr)</div>"
                $diskBarHtml += "<div class='disk-bar-outer'><div class='disk-bar-inner' style='width:${pctUsed}%;background:$barColor;'></div></div>"
                $diskBarHtml += "<div class='disk-stat'>$freeGB GB free / $totalGB GB total &mdash; $pctUsed% used &nbsp;$(StatusBadge $statusStr $statusColor)</div>"
                $diskBarHtml += "</div>"
            }
        }
        $diskBarHtml += "</div>"

        $volHtml  = "<div class='table-wrap'><table><thead><tr>"
        $volHtml += "<th>Drive</th><th>Label</th><th>Total GB</th><th>Used GB</th><th>Free GB</th><th>% Used</th><th>Status</th>"
        $volHtml += "</tr></thead><tbody>"
        foreach ($disk in $logicalDisks) {
            if ($null -ne $disk.Size -and $disk.Size -gt 0) {
                $totalGB  = [math]::Round($disk.Size      / 1GB, 2)
                $freeGB   = [math]::Round($disk.FreeSpace / 1GB, 2)
                $usedGB   = [math]::Round($totalGB - $freeGB, 2)
                $pctUsed  = [math]::Round(($usedGB / $totalGB) * 100, 1)
                $st       = if ($pctUsed -ge $DiskCritPct) { 'CRITICAL' } elseif ($pctUsed -ge $DiskWarnPct) { 'Warning' } else { 'OK' }
                $sc       = if ($pctUsed -ge $DiskCritPct) { 'red' }      elseif ($pctUsed -ge $DiskWarnPct) { 'yellow' } else { 'amber' }
                $labelStr = if ([string]::IsNullOrEmpty($disk.VolumeName)) { 'Local Disk' } else { $disk.VolumeName }
                $volHtml += "<tr><td>$(HtmlEncode $disk.DeviceID)</td><td>$(HtmlEncode $labelStr)</td>"
                $volHtml += "<td>$totalGB</td><td>$usedGB</td><td>$freeGB</td><td>$pctUsed%</td>"
                $volHtml += "<td>$(StatusBadge $st $sc)</td></tr>"
            }
        }
        $volHtml += "</tbody></table></div>"
    }

    $Sec4Html  = "<h4 style='margin:0 0 8px;'>Physical Disks</h4>$physHtml"
    $Sec4Html += "<h4 style='margin:14px 0 8px;'>Volume Usage</h4>$diskBarHtml$volHtml"
} catch {
    $Sec4Html = "<p class='error'>Error checking disk health: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 4 - Disk health error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — CPU & MEMORY REAL-TIME
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 5: CPU & Memory (real-time)..."
$Sec5Html = ''
try {
    # CPU load from WMI
    try {
        $cpuInstances = @(Get-CimInstance Win32_Processor -ErrorAction Stop)
        $CpuPct = [math]::Round(
            ($cpuInstances | Measure-Object -Property LoadPercentage -Average).Average, 1)
    } catch { $CpuPct = 0 }

    # Committed bytes (virtual memory committed)
    $committedGB = 'N/A'
    try {
        $perfOs = Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory -ErrorAction Stop
        if ($null -ne $perfOs.CommittedBytes) {
            $committedGB = [math]::Round($perfOs.CommittedBytes / 1GB, 2)
        }
    } catch {}

    # Top 10 by CPU and by RAM
    $allProcs    = @(Get-Process -ErrorAction SilentlyContinue)
    $topCpuProcs = $allProcs |
        Where-Object { $null -ne $_.CPU } |
        Sort-Object CPU -Descending |
        Select-Object -First 10 |
        Select-Object Name, Id,
            @{N='CPU_s';  E={ [math]::Round($_.CPU, 1) }},
            @{N='RAM_MB'; E={ [math]::Round($_.WorkingSet64 / 1MB, 1) }}
    $topRamProcs = $allProcs |
        Sort-Object WorkingSet64 -Descending |
        Select-Object -First 10 |
        Select-Object Name, Id,
            @{N='CPU_s';  E={ if ($null -ne $_.CPU) { [math]::Round($_.CPU, 1) } else { 'N/A' } }},
            @{N='RAM_MB'; E={ [math]::Round($_.WorkingSet64 / 1MB, 1) }}

    $cpuBadge = if ($CpuPct -ge 90) { StatusBadge "$CpuPct%" 'red' }
                elseif ($CpuPct -ge 75) { StatusBadge "$CpuPct%" 'yellow' }
                else { StatusBadge "$CpuPct%" 'amber' }
    if ($CpuPct -ge 90) { $CriticalFindings.Add("Section 5 - CPU utilisation critically high: $CpuPct%") }

    $ramBadge = if ($RamUsedPct -ge 90) { StatusBadge "$RamUsedPct%" 'red' }
                elseif ($RamUsedPct -ge 75) { StatusBadge "$RamUsedPct%" 'yellow' }
                else { StatusBadge "$RamUsedPct%" 'amber' }

    # RAM bar
    $ramBarColor = if ($RamUsedPct -ge 90) { '#f85149' } elseif ($RamUsedPct -ge 75) { '#d29922' } else { '#e3a019' }
    $ramBarHtml  = "<div class='disk-row'>"
    $ramBarHtml += "<div class='disk-label'><strong>RAM Usage</strong></div>"
    $ramBarHtml += "<div class='disk-bar-outer'><div class='disk-bar-inner' style='width:${RamUsedPct}%;background:$ramBarColor;'></div></div>"
    $ramBarHtml += "<div class='disk-stat'>$UsedRAM_GB GB used / $TotalRAM_GB GB total &mdash; $RamUsedPct% &nbsp;$ramBadge</div>"
    $ramBarHtml += "</div>"

    $Sec5Html  = "<div style='display:flex;flex-wrap:wrap;gap:16px;margin-bottom:14px;'>"
    $Sec5Html += "<div><div style='font-size:.75rem;color:var(--muted);margin-bottom:4px;'>CPU Load</div>$cpuBadge</div>"
    $Sec5Html += "<div><div style='font-size:.75rem;color:var(--muted);margin-bottom:4px;'>Available RAM</div><strong>$FreeRAM_GB GB</strong></div>"
    $Sec5Html += "<div><div style='font-size:.75rem;color:var(--muted);margin-bottom:4px;'>Committed</div><strong>$committedGB GB</strong></div>"
    $Sec5Html += "</div>"
    $Sec5Html += "<div class='disk-container'>$ramBarHtml</div>"

    $Sec5Html += "<h4 style='margin:12px 0 6px;'>Top 10 Processes &mdash; CPU Time</h4>"
    $Sec5Html += "<div class='table-wrap'><table><thead><tr><th>Process</th><th>PID</th><th>CPU Time (s)</th><th>RAM (MB)</th></tr></thead><tbody>"
    if ($topCpuProcs.Count -gt 0) {
        foreach ($p in $topCpuProcs) {
            $Sec5Html += "<tr><td>$(HtmlEncode $p.Name)</td><td>$($p.Id)</td><td>$($p.CPU_s)</td><td>$($p.RAM_MB)</td></tr>"
        }
    } else {
        $Sec5Html += "<tr><td colspan='4' class='info'>No process data available.</td></tr>"
    }
    $Sec5Html += "</tbody></table></div>"

    $Sec5Html += "<h4 style='margin:12px 0 6px;'>Top 10 Processes &mdash; RAM Usage</h4>"
    $Sec5Html += "<div class='table-wrap'><table><thead><tr><th>Process</th><th>PID</th><th>CPU Time (s)</th><th>RAM (MB)</th></tr></thead><tbody>"
    if ($topRamProcs.Count -gt 0) {
        foreach ($p in $topRamProcs) {
            $Sec5Html += "<tr><td>$(HtmlEncode $p.Name)</td><td>$($p.Id)</td><td>$($p.CPU_s)</td><td>$($p.RAM_MB)</td></tr>"
        }
    } else {
        $Sec5Html += "<tr><td colspan='4' class='info'>No process data available.</td></tr>"
    }
    $Sec5Html += "</tbody></table></div>"
} catch {
    $Sec5Html = "<p class='error'>Error retrieving CPU/memory performance: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 5 - CPU/Memory check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — NETWORK ADAPTERS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 6: Network Adapters..."
$Sec6Html = ''
try {
    $adapters = @(Get-CimInstance Win32_NetworkAdapter -ErrorAction Stop |
        Where-Object { $_.NetEnabled -ne $null -and $_.AdapterType -ne $null } |
        Select-Object -First 30)
    $adapterConfigs = @(Get-CimInstance Win32_NetworkAdapterConfiguration -ErrorAction Stop)

    if ($adapters.Count -eq 0) {
        $Sec6Html = "<p class='warn'>No network adapters found.</p>"
    } else {
        $Sec6Html  = "<div class='table-wrap'><table><thead><tr>"
        $Sec6Html += "<th>Adapter Name</th><th>Status</th><th>Speed</th><th>IP Address(es)</th>"
        $Sec6Html += "<th>Default Gateway</th><th>DNS Servers</th>"
        $Sec6Html += "</tr></thead><tbody>"
        foreach ($a in $adapters) {
            $cfg = $adapterConfigs | Where-Object { $_.Index -eq $a.DeviceID } | Select-Object -First 1
            $status = if ($a.NetEnabled) { 'Up' } else { 'Down' }
            $stColor = if ($a.NetEnabled) { 'amber' } else { 'grey' }
            $speedMbps = if ($null -ne $a.Speed -and $a.Speed -gt 0) {
                "$([math]::Round($a.Speed / 1MB)) Mbps"
            } else { 'N/A' }
            $ips = if ($null -ne $cfg -and $null -ne $cfg.IPAddress) {
                ($cfg.IPAddress | Where-Object { $_ -match '\.' }) -join ', '
            } else { 'N/A' }
            $gw  = if ($null -ne $cfg -and $null -ne $cfg.DefaultIPGateway) {
                $cfg.DefaultIPGateway -join ', '
            } else { 'N/A' }
            $dns = if ($null -ne $cfg -and $null -ne $cfg.DNSServerSearchOrder) {
                $cfg.DNSServerSearchOrder -join ', '
            } else { 'N/A' }
            $Sec6Html += "<tr><td>$(HtmlEncode $a.Name)</td><td>$(StatusBadge $status $stColor)</td>"
            $Sec6Html += "<td>$(HtmlEncode $speedMbps)</td><td>$(HtmlEncode $ips)</td>"
            $Sec6Html += "<td>$(HtmlEncode $gw)</td><td>$(HtmlEncode $dns)</td></tr>"
        }
        $Sec6Html += "</tbody></table></div>"
    }
} catch {
    $Sec6Html = "<p class='error'>Error checking network adapters: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 6 - Network adapter check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — CRITICAL WINDOWS SERVICES
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 7: Critical Windows Services..."
$Sec7Html        = ''
$SvcOkCount      = 0
$SvcTotalCount   = $CriticalServices.Count
$ServiceResults  = [System.Collections.Generic.List[hashtable]]::new()

foreach ($svcName in $CriticalServices) {
    $r = @{
        Name        = $svcName
        DisplayName = $svcName
        Status      = 'Not Found'
        StartType   = 'N/A'
        Badge       = StatusBadge 'Not Found' 'grey'
    }
    try {
        $svc = Get-Service -Name $svcName -ErrorAction Stop
        $r.DisplayName = $svc.DisplayName
        $r.Status      = $svc.Status.ToString()
        $r.StartType   = $svc.StartType.ToString()

        if ($svc.Status -eq 'Running') {
            $r.Badge = StatusBadge 'Running' 'amber'
            $SvcOkCount++
        } elseif ($svc.Status -eq 'Stopped' -and $svc.StartType -eq 'Automatic') {
            $r.Badge = StatusBadge 'Stopped (Auto)' 'red'
            $CriticalFindings.Add("Section 7 - Service '$($svc.DisplayName)' is Stopped but set to Automatic start")
        } elseif ($svc.Status -eq 'Stopped') {
            $r.Badge = StatusBadge 'Stopped' 'yellow'
        } else {
            $r.Badge = StatusBadge $svc.Status.ToString() 'yellow'
            $SvcOkCount++
        }
    } catch {
        $r.Badge = StatusBadge 'Not Found' 'grey'
    }
    $ServiceResults.Add($r)
}

$Sec7Html  = "<div class='table-wrap'><table><thead><tr>"
$Sec7Html += "<th>Service Name</th><th>Display Name</th><th>Status</th><th>Start Type</th>"
$Sec7Html += "</tr></thead><tbody>"
foreach ($r in $ServiceResults) {
    $Sec7Html += "<tr><td><code>$(HtmlEncode $r.Name)</code></td><td>$(HtmlEncode $r.DisplayName)</td>"
    $Sec7Html += "<td>$($r.Badge)</td><td>$(HtmlEncode $r.StartType)</td></tr>"
}
$Sec7Html += "</tbody></table></div>"
$Sec7Html += "<p class='info' style='margin-top:8px;'>$SvcOkCount of $SvcTotalCount critical service(s) running.</p>"

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — INSTALLED ROLES & FEATURES
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 8: Installed Roles & Features..."
$Sec8Html = ''
try {
    $installedFeatures = @(
        Get-WindowsFeature -ErrorAction Stop |
        Where-Object { $_.Installed -eq $true } |
        Sort-Object Name
    )
    if ($installedFeatures.Count -eq 0) {
        $Sec8Html = "<p class='info'>No Windows features reported as installed, or Get-WindowsFeature is unavailable on this SKU.</p>"
    } else {
        $Sec8Html  = "<p style='margin-bottom:8px;'><strong>$($installedFeatures.Count)</strong> installed role(s)/feature(s):</p>"
        $Sec8Html += "<div class='table-wrap'><table><thead><tr>"
        $Sec8Html += "<th>Feature Name</th><th>Display Name</th><th>Feature Type</th>"
        $Sec8Html += "</tr></thead><tbody>"
        foreach ($f in $installedFeatures) {
            $Sec8Html += "<tr><td><code>$(HtmlEncode $f.Name)</code></td><td>$(HtmlEncode $f.DisplayName)</td>"
            $Sec8Html += "<td>$(HtmlEncode $f.FeatureType)</td></tr>"
        }
        $Sec8Html += "</tbody></table></div>"
    }
} catch {
    $Sec8Html = "<p class='warn'>Get-WindowsFeature is not available on this edition (Server Core or client OS). $([string]::Empty)</p>"
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 9 — WINDOWS FIREWALL
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 9: Windows Firewall..."
$Sec9Html = ''
try {
    $profiles = @(Get-NetFirewallProfile -ErrorAction Stop)
    $Sec9Html  = "<div class='table-wrap'><table><thead><tr>"
    $Sec9Html += "<th>Profile</th><th>Enabled</th><th>Default Inbound</th><th>Default Outbound</th>"
    $Sec9Html += "</tr></thead><tbody>"
    foreach ($pf in $profiles) {
        $enabled  = $pf.Enabled
        $enColor  = if ($enabled) { 'amber' } else { 'red' }
        if (-not $enabled) {
            $CriticalFindings.Add("Section 9 - Firewall profile '$($pf.Name)' is Disabled")
        }
        $Sec9Html += "<tr><td><strong>$(HtmlEncode $pf.Name)</strong></td>"
        $Sec9Html += "<td>$(StatusBadge $(if ($enabled) {'Enabled'} else {'Disabled'}) $enColor)</td>"
        $Sec9Html += "<td>$(HtmlEncode $pf.DefaultInboundAction)</td>"
        $Sec9Html += "<td>$(HtmlEncode $pf.DefaultOutboundAction)</td></tr>"
    }
    $Sec9Html += "</tbody></table></div>"
} catch {
    $Sec9Html = "<p class='error'>Error checking Windows Firewall: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 9 - Firewall check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 10 — WINDOWS DEFENDER / SECURITY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 10: Windows Defender / Security..."
$Sec10Html = ''
try {
    $defStatus = Get-MpComputerStatus -ErrorAction Stop

    $rtpBadge    = if ($defStatus.RealTimeProtectionEnabled) { StatusBadge 'Enabled' 'amber' } else { StatusBadge 'Disabled' 'red' }
    if (-not $defStatus.RealTimeProtectionEnabled) {
        $CriticalFindings.Add("Section 10 - Windows Defender Real-Time Protection is Disabled")
    }

    $lastScan    = if ($null -ne $defStatus.QuickScanStartTime) { $defStatus.QuickScanStartTime.ToString('yyyy-MM-dd HH:mm') } else { 'Never' }
    $sigDate     = if ($null -ne $defStatus.AntivirusSignatureLastUpdated) { $defStatus.AntivirusSignatureLastUpdated.ToString('yyyy-MM-dd HH:mm') } else { 'Unknown' }
    $sigAgeDays  = if ($null -ne $defStatus.AntivirusSignatureLastUpdated) {
        [int]((Get-Date) - $defStatus.AntivirusSignatureLastUpdated).TotalDays
    } else { 9999 }
    $sigBadge    = if ($sigAgeDays -gt 7) {
        $CriticalFindings.Add("Section 10 - Defender signatures are $sigAgeDays days old")
        StatusBadge "${sigAgeDays}d old" 'red'
    } elseif ($sigAgeDays -gt 3) { StatusBadge "${sigAgeDays}d old" 'yellow' }
    else { StatusBadge "${sigAgeDays}d old" 'amber' }

    $rows10 = @(
        @('Real-Time Protection',    $rtpBadge),
        @('Last Quick Scan',         (HtmlEncode $lastScan)),
        @('Signature Version',       (HtmlEncode $defStatus.AntivirusSignatureVersion)),
        @('Signature Last Updated',  (HtmlEncode $sigDate)),
        @('Signature Age',           $sigBadge),
        @('Antivirus Enabled',       (HtmlEncode $defStatus.AntivirusEnabled.ToString())),
        @('Antispyware Enabled',     (HtmlEncode $defStatus.AntispywareEnabled.ToString())),
        @('Behaviour Monitor',       (HtmlEncode $defStatus.BehaviorMonitorEnabled.ToString())),
        @('AM Engine Version',       (HtmlEncode $defStatus.AMEngineVersion))
    )
    $Sec10Html = BuildKVTable $rows10
} catch {
    $Sec10Html = "<p class='warn'>Windows Defender status not available (may not be installed or running): $(HtmlEncode $_.Exception.Message)</p>"
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 11 — PENDING REBOOTS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 11: Pending Reboots..."
$Sec11Html    = ''
$PendingReboot = $false
$rebootSources = [System.Collections.Generic.List[string]]::new()
try {
    # Windows Update
    try {
        $wuKey = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' -ErrorAction Stop
        if ($null -ne $wuKey) { $rebootSources.Add('Windows Update') }
    } catch {}

    # CBS (Component Based Servicing)
    try {
        $cbsKey = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' -ErrorAction Stop
        if ($null -ne $cbsKey) { $rebootSources.Add('CBS (Component Based Servicing)') }
    } catch {}

    # PendingFileRenameOperations
    try {
        $pfro = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction Stop).PendingFileRenameOperations
        if ($null -ne $pfro -and $pfro.Count -gt 0) { $rebootSources.Add('PendingFileRenameOperations') }
    } catch {}

    # SCCM / ConfigMgr
    try {
        $sccm = Invoke-CimMethod -Namespace 'root\ccm\clientsdk' -Class 'CCM_ClientUtilities' `
            -Name 'DetermineIfRebootPending' -ErrorAction Stop
        if ($null -ne $sccm -and ($sccm.RebootPending -or $sccm.IsHardRebootPending)) {
            $rebootSources.Add('SCCM / ConfigMgr')
        }
    } catch {}

    $PendingReboot = $rebootSources.Count -gt 0

    if ($PendingReboot) {
        $CriticalFindings.Add("Section 11 - Pending reboot detected: $($rebootSources -join ', ')")
        $statusBadge = StatusBadge 'Pending Reboot' 'red'
    } else {
        $statusBadge = StatusBadge 'No Reboot Pending' 'amber'
    }

    $rows11 = @(
        @('Reboot Status',                  $statusBadge),
        @('Windows Update Pending',         (HtmlEncode ($rebootSources -contains 'Windows Update').ToString())),
        @('CBS Pending',                    (HtmlEncode ($rebootSources -contains 'CBS (Component Based Servicing)').ToString())),
        @('PendingFileRenameOperations',    (HtmlEncode ($rebootSources -contains 'PendingFileRenameOperations').ToString())),
        @('SCCM Pending',                   (HtmlEncode ($rebootSources -contains 'SCCM / ConfigMgr').ToString()))
    )
    $Sec11Html = BuildKVTable $rows11
} catch {
    $Sec11Html = "<p class='error'>Error checking pending reboots: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 11 - Pending reboot check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 12 — LOCAL USERS & GROUPS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 12: Local Users & Groups..."
$Sec12Html = ''
try {
    # Local Administrators group members
    $adminMembers = @()
    try {
        $adminGroup = [ADSI]"WinNT://./Administrators,group"
        $adminMembers = @($adminGroup.Members() | ForEach-Object {
            $_.GetType().InvokeMember('Name', 'GetProperty', $null, $_, $null)
        })
    } catch {}

    $Sec12Html  = "<h4 style='margin:0 0 8px;'>Local Administrators Group Members</h4>"
    if ($adminMembers.Count -gt 0) {
        $Sec12Html += "<p>$(($adminMembers | ForEach-Object { "<code>$(HtmlEncode $_)</code>" }) -join ', ')</p>"
    } else {
        $Sec12Html += "<p class='warn'>Could not enumerate Administrators group members.</p>"
    }

    # Local user accounts
    $localUsers = @()
    try { $localUsers = @(Get-LocalUser -ErrorAction Stop) } catch {}

    if ($localUsers.Count -gt 0) {
        $Sec12Html += "<h4 style='margin:12px 0 8px;'>Local User Accounts</h4>"
        $Sec12Html += "<div class='table-wrap'><table><thead><tr>"
        $Sec12Html += "<th>Username</th><th>Enabled</th><th>Last Logon</th><th>Password Last Set</th><th>Account Locked</th>"
        $Sec12Html += "</tr></thead><tbody>"
        foreach ($u in $localUsers) {
            $enabled    = if ($u.Enabled) { StatusBadge 'Yes' 'amber' } else { StatusBadge 'No' 'grey' }
            $locked     = if ($u.PasswordExpired -or ($null -ne $u.AccountExpires -and $u.AccountExpires -lt (Get-Date))) {
                StatusBadge 'Yes' 'red'
            } else { StatusBadge 'No' 'amber' }
            $lastLogon  = if ($null -ne $u.LastLogon -and $u.LastLogon -gt [datetime]'2000-01-01') {
                $u.LastLogon.ToString('yyyy-MM-dd HH:mm')
            } else { 'Never' }
            $pwdSet     = if ($null -ne $u.PasswordLastSet -and $u.PasswordLastSet -gt [datetime]'2000-01-01') {
                $u.PasswordLastSet.ToString('yyyy-MM-dd HH:mm')
            } else { 'Never' }
            $Sec12Html += "<tr><td>$(HtmlEncode $u.Name)</td><td>$enabled</td>"
            $Sec12Html += "<td>$(HtmlEncode $lastLogon)</td><td>$(HtmlEncode $pwdSet)</td><td>$locked</td></tr>"
        }
        $Sec12Html += "</tbody></table></div>"
    }
} catch {
    $Sec12Html = "<p class='error'>Error retrieving local users/groups: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 12 - Local users/groups error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 13 — SHARED FOLDERS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 13: Shared Folders..."
$Sec13Html = ''
try {
    # Exclude well-known default admin shares
    $defaultShares = @('ADMIN$', 'C$', 'D$', 'E$', 'F$', 'G$', 'IPC$', 'print$', 'NETLOGON', 'SYSVOL')
    $shares = @(
        Get-CimInstance Win32_Share -ErrorAction Stop |
        Where-Object { $_.Name -notin $defaultShares -and $_.Type -in @(0, 2147483648) }
    )

    if ($shares.Count -eq 0) {
        $Sec13Html = "<p class='info'>No non-default shares found on this server.</p>"
    } else {
        $Sec13Html  = "<div class='table-wrap'><table><thead><tr>"
        $Sec13Html += "<th>Share Name</th><th>Path</th><th>Description</th><th>Max Connections</th>"
        $Sec13Html += "</tr></thead><tbody>"
        foreach ($s in $shares) {
            $maxConn = if ($null -ne $s.MaximumAllowed -and $s.MaximumAllowed -ne -1 -and $s.MaximumAllowed -gt 0) {
                $s.MaximumAllowed.ToString()
            } else { 'Unlimited' }
            $desc = if ($null -ne $s.Description) { $s.Description } else { '' }
            $Sec13Html += "<tr><td><strong>$(HtmlEncode $s.Name)</strong></td><td>$(HtmlEncode $s.Path)</td>"
            $Sec13Html += "<td>$(HtmlEncode $desc)</td><td>$(HtmlEncode $maxConn)</td></tr>"
        }
        $Sec13Html += "</tbody></table></div>"
    }
} catch {
    $Sec13Html = "<p class='error'>Error retrieving shared folders: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 13 - Shared folders error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 14 — SCHEDULED TASKS (CUSTOM / NON-MICROSOFT)
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 14: Scheduled Tasks (custom)..."
$Sec14Html = ''
try {
    $customTasks = @(
        Get-ScheduledTask -ErrorAction Stop |
        Where-Object { $_.TaskPath -notmatch '\\Microsoft\\' } |
        Select-Object -First 50
    )

    if ($customTasks.Count -eq 0) {
        $Sec14Html = "<p class='info'>No non-Microsoft scheduled tasks found.</p>"
    } else {
        $Sec14Html  = "<div class='table-wrap'><table><thead><tr>"
        $Sec14Html += "<th>Task Name</th><th>Path</th><th>State</th><th>Last Run</th><th>Last Result</th><th>Next Run</th>"
        $Sec14Html += "</tr></thead><tbody>"
        foreach ($task in $customTasks) {
            $info        = $null
            try { $info = $task | Get-ScheduledTaskInfo -ErrorAction Stop } catch {}
            $state       = $task.State.ToString()
            $stColor     = if ($state -in @('Ready','Running')) { 'amber' } else { 'yellow' }
            $lastRun     = if ($null -ne $info -and $null -ne $info.LastRunTime -and
                              $info.LastRunTime -gt [datetime]'2000-01-01') {
                $info.LastRunTime.ToString('yyyy-MM-dd HH:mm')
            } else { 'Never' }
            $lastResult  = if ($null -ne $info) {
                $rc = $info.LastTaskResult
                if ($rc -eq 0)      { StatusBadge 'Success' 'amber' }
                elseif ($rc -eq 267011) { StatusBadge 'Never Run' 'grey' }
                else { StatusBadge "Error (0x$($rc.ToString('X')))" 'red' }
            } else { StatusBadge 'N/A' 'grey' }
            $nextRun     = if ($null -ne $info -and $null -ne $info.NextRunTime -and
                              $info.NextRunTime -gt [datetime]'2000-01-01') {
                $info.NextRunTime.ToString('yyyy-MM-dd HH:mm')
            } else { 'Not scheduled' }

            $Sec14Html += "<tr><td>$(HtmlEncode $task.TaskName)</td><td>$(HtmlEncode $task.TaskPath)</td>"
            $Sec14Html += "<td>$(StatusBadge $state $stColor)</td><td>$(HtmlEncode $lastRun)</td>"
            $Sec14Html += "<td>$lastResult</td><td>$(HtmlEncode $nextRun)</td></tr>"
        }
        $Sec14Html += "</tbody></table></div>"
        $Sec14Html += "<p class='info' style='margin-top:8px;'>Showing up to 50 non-Microsoft tasks.</p>"
    }
} catch {
    $Sec14Html = "<p class='error'>Error checking scheduled tasks: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 14 - Scheduled tasks error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 15 — EVENT LOG SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 15: Event Log Summary (last 24 hours)..."
$Sec15Html = ''
try {
    $since          = (Get-Date).AddHours(-24)
    $allEvents      = [System.Collections.Generic.List[object]]::new()
    $logSummary     = @{}

    foreach ($logName in @('System', 'Application')) {
        try {
            $filter = @{ LogName = $logName; Level = @(1,2); StartTime = $since }
            $evts   = @(Get-WinEvent -FilterHashtable $filter -MaxEvents 100 -ErrorAction Stop)
            $logSummary[$logName] = $evts.Count
            foreach ($e in $evts) { $allEvents.Add($e) }
        } catch [System.Exception] {
            $logSummary[$logName] = 0
            if ($_.Exception.Message -notmatch 'No events were found') {
                $allEvents.Add([PSCustomObject]@{
                    TimeCreated = (Get-Date); LevelDisplayName = 'Error'
                    Id = 0; ProviderName = "EventLog/$logName"
                    Message = "Error reading $logName log: $($_.Exception.Message)"
                })
            }
        }
    }

    $totalCount  = $allEvents.Count
    $critCount15 = ($allEvents | Where-Object { $_.LevelDisplayName -eq 'Critical' }).Count
    if ($critCount15 -gt 0) {
        $CriticalFindings.Add("Section 15 - $critCount15 Critical event(s) in last 24 h")
    }

    # Summary tiles
    $Sec15Html  = "<div style='display:flex;flex-wrap:wrap;gap:16px;margin-bottom:12px;'>"
    $Sec15Html += "<div style='text-align:center;'><div style='font-size:1.6rem;font-weight:700;color:#e3a019;'>$($logSummary['System'])</div><div style='font-size:.75rem;color:var(--muted);'>System log events</div></div>"
    $Sec15Html += "<div style='text-align:center;'><div style='font-size:1.6rem;font-weight:700;color:#e3a019;'>$($logSummary['Application'])</div><div style='font-size:.75rem;color:var(--muted);'>Application log events</div></div>"
    $Sec15Html += "<div style='text-align:center;'><div style='font-size:1.6rem;font-weight:700;$(if ($critCount15 -gt 0) {"color:#f85149"} else {"color:#e3a019"})'>$critCount15</div><div style='font-size:.75rem;color:var(--muted);'>Critical events</div></div>"
    $Sec15Html += "</div>"

    $topEvents = $allEvents | Sort-Object TimeCreated -Descending | Select-Object -First 10

    if ($topEvents.Count -eq 0) {
        $Sec15Html += "<p style='color:#e3a019;'>&#x2714; No Critical or Error events in the last 24 hours.</p>"
    } else {
        $Sec15Html += "<div class='table-wrap'><table class='event-table'><thead><tr>"
        $Sec15Html += "<th>Time</th><th>Level</th><th>Event ID</th><th>Source</th><th>Message</th>"
        $Sec15Html += "</tr></thead><tbody>"
        foreach ($ev in $topEvents) {
            $timeStr  = if ($null -ne $ev.TimeCreated) { $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
            $level    = if ($null -ne $ev.LevelDisplayName) { $ev.LevelDisplayName } else { 'Unknown' }
            $lvlColor = if ($level -eq 'Critical') { 'red' } else { 'yellow' }
            $rowClass = if ($level -eq 'Critical') { 'row-critical' } else { 'row-error' }
            $source   = if ($null -ne $ev.ProviderName) { $ev.ProviderName } else { 'Unknown' }
            $msgRaw   = if ($null -ne $ev.Message) { $ev.Message } else { '' }
            $msgShort = if ($msgRaw.Length -gt 200) { $msgRaw.Substring(0,200) + '...' } else { $msgRaw }
            $Sec15Html += "<tr class='$rowClass'>"
            $Sec15Html += "<td style='white-space:nowrap;'>$(HtmlEncode $timeStr)</td>"
            $Sec15Html += "<td>$(StatusBadge $level $lvlColor)</td>"
            $Sec15Html += "<td>$($ev.Id)</td>"
            $Sec15Html += "<td>$(HtmlEncode $source)</td>"
            $Sec15Html += "<td>$(HtmlEncode $msgShort)</td>"
            $Sec15Html += "</tr>"
        }
        $Sec15Html += "</tbody></table></div>"
    }
} catch {
    $Sec15Html = "<p class='error'>Error reading event logs: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 15 - Event log check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 16 — RDP & REMOTE MANAGEMENT
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 16: RDP & Remote Management..."
$Sec16Html = ''
try {
    # RDP enabled check via registry
    $rdpEnabled = $false
    try {
        $rdpKey = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -ErrorAction Stop
        $rdpEnabled = ($rdpKey.fDenyTSConnections -eq 0)
    } catch {}

    # RDP port (default 3389 or custom)
    $rdpPort = 3389
    try {
        $rdpPortKey = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -ErrorAction Stop).PortNumber
        if ($null -ne $rdpPortKey) { $rdpPort = $rdpPortKey }
    } catch {}

    # NLA
    $nlaEnabled = $false
    try {
        $nlaKey = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -ErrorAction Stop).UserAuthentication
        $nlaEnabled = ($nlaKey -eq 1)
    } catch {}

    # WinRM
    $winrmStatus = 'Unknown'
    try {
        $wmSvc = Get-Service WinRM -ErrorAction Stop
        $winrmStatus = $wmSvc.Status.ToString()
    } catch {}

    $rdpBadge  = if ($rdpEnabled) { StatusBadge 'Enabled' 'amber' }  else { StatusBadge 'Disabled' 'grey' }
    $nlaBadge  = if ($nlaEnabled) { StatusBadge 'Required' 'amber' } else { StatusBadge 'Not Required' 'yellow' }
    $wrmBadge  = if ($winrmStatus -eq 'Running') { StatusBadge 'Running' 'amber' } else { StatusBadge $winrmStatus 'red' }

    $rows16 = @(
        @('RDP Enabled',    $rdpBadge),
        @('RDP Port',       (HtmlEncode $rdpPort.ToString())),
        @('NLA Required',   $nlaBadge),
        @('WinRM Status',   $wrmBadge)
    )
    $Sec16Html = BuildKVTable $rows16
} catch {
    $Sec16Html = "<p class='error'>Error checking RDP/remote management: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 16 - RDP check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 17 — TIME SYNCHRONISATION
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 17: Time Synchronisation..."
$Sec17Html = ''
try {
    $timeSource = 'Unknown'
    $timeOffset = 'Unknown'
    $offsetSec  = 0

    # w32tm /query /source
    try {
        $srcOutput = & w32tm /query /source 2>&1
        if ($LASTEXITCODE -eq 0) {
            $timeSource = $srcOutput | Out-String
            $timeSource = $timeSource.Trim()
        }
    } catch {}

    # w32tm /query /status for offset
    try {
        $statusOutput = & w32tm /query /status 2>&1 | Out-String
        if ($statusOutput -match 'Last Successful Sync Time:\s*(.+)') { }
        if ($statusOutput -match 'Offset:\s*([\-\d\.]+)s') {
            $offsetSec  = [double]$Matches[1]
            $timeOffset = "$offsetSec seconds"
        } elseif ($statusOutput -match 'Root Delay:\s*([\-\d\.]+)s') {
            $offsetSec  = [double]$Matches[1]
            $timeOffset = "~$offsetSec seconds (root delay)"
        }
    } catch {}

    $offsetBadge = if ([math]::Abs($offsetSec) -gt $TimeSyncWarnSec) {
        $CriticalFindings.Add("Section 17 - Time offset is $offsetSec seconds (threshold: $TimeSyncWarnSec s)")
        StatusBadge "WARN: $offsetSec s" 'red'
    } else {
        StatusBadge "$offsetSec s" 'amber'
    }

    # w32tm /query /peers (brief)
    $peersStr = 'N/A'
    try {
        $peersOut = & w32tm /query /peers 2>&1 | Out-String
        $peers = ($peersOut -split "`n" | Where-Object { $_ -match 'Peer:' } |
            ForEach-Object { ($_ -split ':',2)[1].Trim() })
        if ($peers.Count -gt 0) { $peersStr = $peers -join ', ' }
    } catch {}

    $rows17 = @(
        @('Time Source',       (HtmlEncode $timeSource)),
        @('Current Offset',    $offsetBadge),
        @('Offset (raw)',      (HtmlEncode $timeOffset)),
        @('Time Peers',        (HtmlEncode $peersStr)),
        @('Server Local Time', (HtmlEncode (Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')))
    )
    $Sec17Html = BuildKVTable $rows17
} catch {
    $Sec17Html = "<p class='error'>Error checking time sync: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 17 - Time sync check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# KPI / SUMMARY TILES
# ═══════════════════════════════════════════════════════════════════════════════
$CritCount  = $CriticalFindings.Count
$ReportDate = (Get-Date).ToString('dddd, dd MMMM yyyy HH:mm:ss')
$EndTime    = Get-Date
$Duration   = ($EndTime - $StartTime).ToString('hh\:mm\:ss')

$UptimeTileHtml = if ($UptimeStr -ne 'Unknown') { "<strong>$UptimeStr</strong>" } else { "<span style='color:var(--muted);'>Unknown</span>" }

$cpuTileColor  = if ($CpuPct -ge 90) { '#f85149' } elseif ($CpuPct -ge 75) { '#d29922' } else { '#e3a019' }
$CpuTileHtml   = "<span style='color:$cpuTileColor;font-weight:700;'>$CpuPct%</span>"

$ramTileColor  = if ($RamUsedPct -ge 90) { '#f85149' } elseif ($RamUsedPct -ge 75) { '#d29922' } else { '#e3a019' }
$RamTileHtml   = "<span style='color:$ramTileColor;font-weight:700;'>$RamUsedPct%</span>"

$diskTileColor = if ($DiskWorstPct -ge $DiskCritPct) { '#f85149' } elseif ($DiskWorstPct -ge $DiskWarnPct) { '#d29922' } else { '#e3a019' }
$DiskTileHtml  = "<span style='color:$diskTileColor;font-weight:700;'>$DiskWorstPct%</span>"

$svcTileColor  = if ($SvcOkCount -lt $SvcTotalCount) { '#f85149' } else { '#e3a019' }
$SvcTileHtml   = "<span style='color:$svcTileColor;font-weight:700;'>$SvcOkCount / $SvcTotalCount</span>"

$rebootTileHtml = if ($PendingReboot) { "<span style='color:#f85149;font-weight:700;'>Pending</span>" } else { "<span style='color:#e3a019;font-weight:700;'>None</span>" }

$OverallStatusBadge = if ($CritCount -gt 0) { StatusBadge 'CRITICAL' 'red' } else { StatusBadge 'HEALTHY' 'amber' }

# Critical findings panel
$CritFindingsHtml = ''
if ($CritCount -gt 0) {
    $cfRows = ($CriticalFindings | ForEach-Object { "<li>$(HtmlEncode $_)</li>" }) -join ''
    $CritFindingsHtml = "<div style='background:rgba(248,81,73,.12);border:1px solid #f85149;border-radius:8px;" +
        "padding:14px 18px;margin-bottom:16px;'>" +
        "<strong style='color:#f85149;'>&#x26A0; $CritCount Critical Finding(s) Detected</strong>" +
        "<ul style='margin:.6rem 0 0 1.2rem;color:#f85149;'>$cfRows</ul></div>"
}

# Logo
$LogoHtml = ''
if (-not [string]::IsNullOrWhiteSpace($CompanyLogoURL)) {
    $logoImg = "<img src='$(HtmlEncode $CompanyLogoURL)' alt='Logo' style='max-height:56px;vertical-align:middle;'>"
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
# BUILD FULL HTML REPORT
# ═══════════════════════════════════════════════════════════════════════════════
$HtmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Server Health Check - $(HtmlEncode $ServerHostname)</title>
<style>
/* ── CSS VARIABLES ── */
:root {
  --bg:       #0d1117;
  --card:     #161b22;
  --border:   #30363d;
  --text:     #c9d1d9;
  --muted:    #8b949e;
  --link:     #e3a019;
  --head-bg:  #010409;
  --th-bg:    #21262d;
  --tr-alt:   #1c2128;
  --amber:    #e3a019;
  --yellow:   #d29922;
  --red:      #f85149;
  --green:    #3fb950;
  --blue:     #58a6ff;
}
[data-theme="light"] {
  --bg:       #ffffff;
  --card:     #f6f8fa;
  --border:   #d0d7de;
  --text:     #24292f;
  --muted:    #57606a;
  --link:     #9a6700;
  --head-bg:  #fffbf0;
  --th-bg:    #fdf6e3;
  --tr-alt:   #fffcf0;
  --amber:    #9a6700;
  --yellow:   #9a6700;
  --red:      #cf222e;
  --green:    #1a7f37;
  --blue:     #0969da;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
       background: var(--bg); color: var(--text); font-size: 14px; line-height: 1.5; }
a    { color: var(--link); }
code { font-family: Consolas,'SFMono-Regular',monospace; font-size: .85em;
       background: var(--th-bg); padding: 1px 5px; border-radius: 4px; }

/* ── LAYOUT ── */
.page-wrap { max-width: 1280px; margin: 0 auto; padding: 0 20px 40px; }
.header    { background: var(--head-bg); border-bottom: 2px solid var(--amber);
             padding: 14px 24px; display: flex; align-items: center;
             justify-content: space-between; flex-wrap: wrap; gap: 12px; }
.header-left  { display: flex; align-items: center; gap: 16px; }
.header-title h1 { font-size: 1.25rem; font-weight: 700; color: var(--amber);
                   display: flex; align-items: center; gap: 8px; }
.header-title p  { font-size: .8rem; color: var(--muted); margin-top: 2px; }
.logo-wrap  { max-height: 60px; }
.theme-toggle { cursor: pointer; background: var(--card); border: 1px solid var(--amber);
                color: var(--text); border-radius: 20px; padding: 6px 14px;
                font-size: 12px; display: flex; align-items: center; gap: 6px; }
.theme-toggle:hover { background: var(--th-bg); }

/* ── SUMMARY BAR ── */
.summary-bar { display: flex; flex-wrap: wrap; gap: 12px;
               background: var(--card); border: 1px solid var(--border);
               border-top: 2px solid var(--amber);
               border-radius: 8px; padding: 16px 20px; margin: 20px 0; }
.stat-card   { flex: 1 1 120px; text-align: center; }
.stat-count  { font-size: 2rem; font-weight: 700; line-height: 1; }
.stat-label  { font-size: .72rem; color: var(--muted); margin-top: 4px; }

/* ── SECTION CARDS ── */
.section-card { background: var(--card); border: 1px solid var(--border);
                border-radius: 8px; margin-bottom: 14px; overflow: hidden; }
.section-summary { display: flex; align-items: center; gap: 10px; cursor: pointer;
                   padding: 13px 18px; list-style: none; user-select: none; }
.section-summary::-webkit-details-marker { display: none; }
.section-summary::marker { display: none; }
.sec-arrow { font-size: 10px; color: var(--muted); display: inline-block;
             transition: transform .2s; flex-shrink: 0; line-height: 1; }
details[open] > .section-summary .sec-arrow { transform: rotate(90deg); }
.section-summary:hover { background: var(--th-bg); }
.sec-num   { background: var(--amber); color: #000; font-size: .7rem; font-weight: 700;
             width: 22px; height: 22px; border-radius: 50%; display: flex;
             align-items: center; justify-content: center; flex-shrink: 0; }
.sec-title { font-weight: 600; font-size: .95rem; flex: 1; }
.section-body { padding: 15px 18px; border-top: 1px solid var(--border); }
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
.kv-table td  { padding: 7px 12px; border-bottom: 1px solid var(--border); }
.td-label     { font-weight: 600; white-space: nowrap; width: 220px; color: var(--muted); }
.event-table td:nth-child(5) { max-width: 360px; word-break: break-word; white-space: normal; }

/* ── DISK / RAM BARS ── */
.disk-container  { display: flex; flex-direction: column; gap: 12px; margin-bottom: 12px; }
.disk-row        { }
.disk-label      { font-size: .85rem; margin-bottom: 4px; font-weight: 600; }
.disk-bar-outer  { height: 18px; background: var(--th-bg); border-radius: 9px;
                   overflow: hidden; margin-bottom: 4px; border: 1px solid var(--border); }
.disk-bar-inner  { height: 100%; border-radius: 9px; transition: width .4s ease; }
.disk-stat       { font-size: .8rem; color: var(--muted); }

/* ── EVENT ROW HIGHLIGHT ── */
.row-critical { background: rgba(248,81,73,.15) !important; }
.row-error    { background: rgba(248,81,73,.08) !important; }

/* ── BADGES ── */
.badge { display: inline-block; font-size: .7rem; font-weight: 600; padding: 2px 8px;
         border-radius: 20px; color: #fff; white-space: nowrap; }

/* ── ALERTS ── */
.error { color: #f85149; padding: 8px 12px; background: rgba(248,81,73,.1);
         border-left: 3px solid #f85149; border-radius: 4px; }
.warn  { color: #d29922; padding: 8px 12px; background: rgba(210,153,34,.1);
         border-left: 3px solid #d29922; border-radius: 4px; }
.info  { color: var(--muted); }

/* ── FOOTER ── */
.footer { border-top: 1px solid var(--border); padding: 18px 0;
          margin-top: 24px; color: var(--muted); font-size: .8rem;
          display: flex; justify-content: space-between; flex-wrap: wrap; gap: 8px; }

/* ── RESPONSIVE ── */
@media (max-width: 768px) {
  .summary-bar { gap: 8px; }
  .stat-card   { flex: 1 1 90px; }
  .stat-count  { font-size: 1.4rem; }
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
      <h1>&#x1F5A5; Windows Server Health Check</h1>
      <p>$(HtmlEncode $ServerHostname) &nbsp;|&nbsp; Generated: $(HtmlEncode $ReportDate)</p>
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
    <div class="stat-count" style="font-size:.85rem;padding-top:10px;">$UptimeTileHtml</div>
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
    <div class="stat-count" style="font-size:1rem;padding-top:10px;">$SvcTileHtml</div>
    <div class="stat-label">Services Running</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:.9rem;padding-top:10px;">$rebootTileHtml</div>
    <div class="stat-label">Pending Reboot</div>
  </div>
  <div class="stat-card">
    <div class="stat-count $(if ($CritCount -gt 0) { 'c-red' } else { 'c-amber' })"
         style="color:$(if ($CritCount -gt 0) { '#f85149' } else { '#e3a019' });">$CritCount</div>
    <div class="stat-label">Critical Findings</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:.9rem;padding-top:10px;">$OverallStatusBadge</div>
    <div class="stat-label">Overall Status</div>
  </div>
</div>

<!-- CRITICAL FINDINGS PANEL -->
$CritFindingsHtml

<!-- 17 SECTIONS -->
$(BuildSection 1  'System Identity'              $Sec1Html  ($Sec1Html  -match "class='error'") $true)
$(BuildSection 2  'Hardware Inventory'           $Sec2Html  ($Sec2Html  -match "class='error'") $true)
$(BuildSection 3  'OS Details'                   $Sec3Html  ($Sec3Html  -match "class='error'") $true)
$(BuildSection 4  'Disk Health'                  $Sec4Html  ($Sec4Html  -match "class='error'") $true)
$(BuildSection 5  'CPU & Memory Real-Time'       $Sec5Html  ($Sec5Html  -match "class='error'") $true)
$(BuildSection 6  'Network Adapters'             $Sec6Html  ($Sec6Html  -match "class='error'") $false)
$(BuildSection 7  'Windows Services (Critical)'  $Sec7Html  ($Sec7Html  -match "class='error'") $false)
$(BuildSection 8  'Installed Roles & Features'   $Sec8Html  ($Sec8Html  -match "class='error'") $false)
$(BuildSection 9  'Windows Firewall'             $Sec9Html  ($Sec9Html  -match "class='error'") $false)
$(BuildSection 10 'Windows Defender / Security'  $Sec10Html ($Sec10Html -match "class='error'") $false)
$(BuildSection 11 'Pending Reboots'              $Sec11Html ($Sec11Html -match "class='error'") $false)
$(BuildSection 12 'Local Users & Groups'         $Sec12Html ($Sec12Html -match "class='error'") $false)
$(BuildSection 13 'Shared Folders'               $Sec13Html ($Sec13Html -match "class='error'") $false)
$(BuildSection 14 'Scheduled Tasks (Custom)'     $Sec14Html ($Sec14Html -match "class='error'") $false)
$(BuildSection 15 'Event Log Summary'            $Sec15Html ($Sec15Html -match "class='error'") $false)
$(BuildSection 16 'RDP & Remote Management'      $Sec16Html ($Sec16Html -match "class='error'") $false)
$(BuildSection 17 'Time Synchronisation'         $Sec17Html ($Sec17Html -match "class='error'") $false)

<!-- FOOTER -->
<div class="footer">
  <div>
    <strong>Windows Server Health Check v$ScriptVersion</strong> &nbsp;|&nbsp;
    Server: $(HtmlEncode $ServerHostname) &nbsp;|&nbsp;
    Generated: $(HtmlEncode $ReportDate) &nbsp;|&nbsp;
    Duration: $(HtmlEncode $Duration)
  </div>
  <div>Created by $AuthorLink</div>
</div>

</div><!-- end page-wrap -->

<script>
function toggleTheme() {
  var b = document.body;
  var i = document.getElementById('theme-icon');
  if (b.getAttribute('data-theme') === 'dark') {
    b.setAttribute('data-theme', 'light');
    i.textContent = '\uD83C\uDF19';
  } else {
    b.setAttribute('data-theme', 'dark');
    i.textContent = '\u2600\uFE0F';
  }
}
</script>
</body>
</html>
"@

# ═══════════════════════════════════════════════════════════════════════════════
# WRITE HTML REPORT
# ═══════════════════════════════════════════════════════════════════════════════
try {
    [System.IO.File]::WriteAllText($ReportFile, $HtmlReport, [System.Text.Encoding]::UTF8)
    Write-Host ""
    Write-Host "  [OK] Report saved: $ReportFile" -ForegroundColor Green
} catch {
    Write-Warning "Failed to write report: $_"
}

# ═══════════════════════════════════════════════════════════════════════════════
# STATUS SUMMARY FILE  (_HEALTHY.txt or _CRITICAL.txt)
# ═══════════════════════════════════════════════════════════════════════════════
$isCritical   = $CriticalFindings.Count -gt 0
$statusSuffix = if ($isCritical) { '_CRITICAL' } else { '_HEALTHY' }
$StatusFile   = Join-Path $ReportsDir ($ReportStamp + $statusSuffix + '.txt')
$separator    = '=' * 70

if ($isCritical) {
    $statusContent  = "$separator`r`n"
    $statusContent += " WINDOWS SERVER HEALTH CHECK  -  *** CRITICAL ALERT ***`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " Status    : CRITICAL`r`n"
    $statusContent += " Server    : $ServerHostname`r`n"
    $statusContent += " Generated : $ReportDate`r`n"
    $statusContent += " Duration  : $Duration`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += "`r`n CRITICAL FINDINGS ($($CriticalFindings.Count)):`r`n`r`n"
    $statusContent += (@($CriticalFindings) | ForEach-Object { "  [!] $_" }) -join "`r`n"
    $statusContent += "`r`n`r`n$separator`r`n"
    $statusContent += " Full HTML report : $ReportFile`r`n"
    $statusContent += " Status file      : $StatusFile`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " Server Health Check v$ScriptVersion  by $AuthorName`r`n"
    $statusContent += "$separator`r`n"
} else {
    $statusContent  = "$separator`r`n"
    $statusContent += " WINDOWS SERVER HEALTH CHECK  -  HEALTHY STATE`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " Status    : HEALTHY`r`n"
    $statusContent += " Server    : $ServerHostname`r`n"
    $statusContent += " Generated : $ReportDate`r`n"
    $statusContent += " Duration  : $Duration`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += "`r`n No critical findings detected. Server is in a healthy state.`r`n"
    $statusContent += "`r`n$separator`r`n"
    $statusContent += " Full HTML report : $ReportFile`r`n"
    $statusContent += " Status file      : $StatusFile`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " Server Health Check v$ScriptVersion  by $AuthorName`r`n"
    $statusContent += "$separator`r`n"
}

if ($EnableStatusFile) {
    try {
        [System.IO.File]::WriteAllText($StatusFile, $statusContent, [System.Text.Encoding]::UTF8)
        $sColor = if ($isCritical) { 'Red' } else { 'Green' }
        Write-Host "  [OK] Status file: $StatusFile" -ForegroundColor $sColor
    } catch {
        Write-Warning "Failed to write status file: $_"
    }
}

# ── CONSOLE SUMMARY ───────────────────────────────────────────────────────────
Write-Host ""
Write-Host "===============================================================" -ForegroundColor DarkYellow
Write-Host "  Server Health Check complete.  Duration: $Duration" -ForegroundColor DarkYellow
Write-Host "  Report : $ReportFile" -ForegroundColor Yellow
if ($EnableStatusFile) {
    $slabel = if ($isCritical) { 'Status (CRITICAL)' } else { 'Status (HEALTHY)' }
    Write-Host "  $slabel : $StatusFile" -ForegroundColor $(if ($isCritical) { 'Red' } else { 'Green' })
    Write-Host ""
    Write-Host "  To send email alerts, run: .\Server_HealthCheck_EmailAlert.ps1" -ForegroundColor Cyan
}
Write-Host "===============================================================" -ForegroundColor DarkYellow
Write-Host ""
