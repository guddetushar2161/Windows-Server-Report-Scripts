#Requires -Version 5.1
<#
.SYNOPSIS
    Comprehensive WSUS (Windows Server Update Services) Health Check Script

.DESCRIPTION
    Performs a thorough health check of the local WSUS server environment and
    exports the results to a single self-contained HTML file with a professional
    dark-themed dashboard UI.

    Checks performed:
      1.  WSUS Server Info
      2.  Synchronization Status
      3.  IIS Application Pool Health
      4.  Database Health
      5.  Client Reporting Summary
      6.  Update Compliance Overview
      7.  Top 10 Non-Compliant Computers
      8.  WSUS Services Status
      9.  Disk Space Check
      10. Event Log Check
      11. Cleanup Recommendations

.NOTES
    Version    : 1.0.0
    Author     : Tushar Gudde
    Requires   : PowerShell 5.1+, WSUS role installed on this server
    Permissions: Local Administrator / WSUS Administrators group
    Compatible : Windows Server 2016, 2019, 2022
#>

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
# Company Branding
$CompanyLogoURL = ''                         # URL/path to company logo (PNG/SVG). Leave blank to skip.
$CompanyWebsite = 'https://tushargudde.tech' # Company website URL for logo hyperlink.

# Author
$AuthorName = 'Tushar Gudde'                 # Author name shown in footer

# Disk-space warning thresholds (GB)
$DiskWarnGB     = 10
$DiskCritGB     = 5

# Client staleness thresholds (days)
$StaleDay7  = 7
$StaleDay14 = 14
$StaleDay30 = 30

# Compliance: flag computers with more than this many failed updates
$FailedUpdateThreshold = 5

# Cleanup: flag if WSUS cleanup was not run within this many days
$CleanupStaleDays = 30
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
$ReportFile = Join-Path $ReportsDir ("WSUS_Health_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

# ── HELPER FUNCTIONS ──────────────────────────────────────────────────────────
function HtmlEncode {
    param([string]$text)
    if ([string]::IsNullOrEmpty($text)) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function StatusBadge {
    param([string]$text, [string]$color)
    $map = @{
        green  = '#238636'
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
    param([array]$rows)   # each element: @('Label','Value')
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
        "<span style='color:#238636;'>&#x2714;</span>"
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

# Collect critical findings throughout all sections
$CriticalFindings = [System.Collections.Generic.List[string]]::new()

Write-Host ""
Write-Host "+==============================================================+" -ForegroundColor DarkCyan
Write-Host "|       WSUS Health Check  v$ScriptVersion                            |" -ForegroundColor DarkCyan
Write-Host "+==============================================================+" -ForegroundColor DarkCyan
Write-Host ""

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 0  -  SERVER DETAILS  (host running the script)
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 0: Gathering Server Details..."
$Sec0Html = ''
try {
    $cs   = Get-CimInstance Win32_ComputerSystem  -ErrorAction Stop
    $bios = Get-CimInstance Win32_BIOS            -ErrorAction Stop
    $os   = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $cpu  = @(Get-CimInstance Win32_Processor     -ErrorAction Stop)

    $totalRAM_GB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
    $isVirtual   = ($cs.Model        -match 'Virtual|VMware|VirtualBox|QEMU|KVM|Xen|HVM') -or
                   ($cs.Manufacturer -match 'VMware|QEMU|Xen|Parallels|innotek') -or
                   ($cs.Manufacturer -eq 'Microsoft Corporation' -and $cs.Model -match 'Virtual')
    $serverTypeBadge = if ($isVirtual) { StatusBadge 'Virtual Machine' 'blue' } else { StatusBadge 'Physical Server' 'green' }
    $cpuNames    = ($cpu | ForEach-Object { HtmlEncode (if ($null -ne $_.Name) { $_.Name.Trim() } else { 'Unknown' }) } | Select-Object -Unique) -join '; '

    $installDate  = if ($null -ne $os.InstallDate)    { $os.InstallDate.ToString('yyyy-MM-dd')            } else { 'N/A' }
    $lastBootTime = if ($null -ne $os.LastBootUpTime) { $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }

    $rows0 = @(
        @('Hostname',        (HtmlEncode $cs.Name)),
        @('Manufacturer',    (HtmlEncode $cs.Manufacturer)),
        @('Model',           (HtmlEncode $cs.Model)),
        @('Server Type',     $serverTypeBadge),
        @('Serial Number',   (HtmlEncode $bios.SerialNumber)),
        @('BIOS Version',    (HtmlEncode $bios.SMBIOSBIOSVersion)),
        @('Processors',      "$($cpu.Count) x $cpuNames"),
        @('Total RAM',       "$totalRAM_GB GB"),
        @('OS Name',         (HtmlEncode $os.Caption)),
        @('OS Version',      (HtmlEncode $os.Version)),
        @('OS Build',        (HtmlEncode $os.BuildNumber)),
        @('OS Install Date', (HtmlEncode $installDate)),
        @('Last Boot Time',  (HtmlEncode $lastBootTime))
    )

    $rows0Html = ($rows0 | ForEach-Object {
        "<tr><td class='td-label'>$($_[0])</td><td>$($_[1])</td></tr>"
    }) -join ''

    $Sec0Html = "<div class='table-wrap'><table class='kv-table'><tbody>$rows0Html</tbody></table></div>"
} catch {
    $Sec0Html = "<p class='error'>Error retrieving Server Details: $(HtmlEncode $_.Exception.Message)</p>"
}

# ── CHECK FOR WSUS MODULE ────────────────────────────────────────────────────
$WsusModuleAvailable = $false
$WsusServer          = $null
$WsusModuleError     = ''

try {
    # UpdateServices module ships with the WSUS role / RSAT-UpdateServices feature
    Import-Module UpdateServices -ErrorAction Stop
    $WsusModuleAvailable = $true
    Write-Progress2 "UpdateServices module loaded."
} catch {
    $WsusModuleError = $_.Exception.Message
    Write-Warning "UpdateServices module not found. WSUS role may not be installed. Error: $WsusModuleError"
}

if ($WsusModuleAvailable) {
    try {
        $WsusServer = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
        Write-Progress2 "Connected to WSUS server: $($WsusServer.Name)"
    } catch {
        $WsusModuleError = "Connected to UpdateServices module but could not reach WSUS server: $($_.Exception.Message)"
        $WsusModuleAvailable = $false
        Write-Warning $WsusModuleError
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1  -  WSUS SERVER INFO
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 1: WSUS Server Info..."
$Sec1Html = ''
$WsusServerName = $env:COMPUTERNAME

if (-not $WsusModuleAvailable) {
    $errMsg = if ([string]::IsNullOrEmpty($WsusModuleError)) { 'UpdateServices module not found or WSUS role not installed.' } else { $WsusModuleError }
    $Sec1Html = "<div class='error'><strong>WSUS Role Not Available</strong><br>$( HtmlEncode $errMsg )<br><br>To install WSUS: <code>Install-WindowsFeature -Name UpdateServices -IncludeManagementTools</code></div>"
    $CriticalFindings.Add("Section 1 - WSUS role not available: $errMsg")
} else {
    try {
        $WsusServerName   = HtmlEncode $WsusServer.Name
        $WsusVersion      = 'N/A'
        $WsusPort         = 'N/A'
        $WsusSsl          = 'N/A'
        $WsusDbType       = 'N/A'
        $WsusContentPath  = 'N/A'
        $WsusContentFree  = 'N/A'
        $WsusContentBadge = ''

        try { $WsusVersion = HtmlEncode $WsusServer.Version.ToString() } catch {}
        try { $WsusPort    = HtmlEncode $WsusServer.PortNumber.ToString() } catch {}
        try {
            $ssl = $WsusServer.UseSecureConnection
            $WsusSsl = if ($ssl) { StatusBadge 'SSL Enabled' 'green' } else { StatusBadge 'SSL Disabled' 'yellow' }
        } catch {}

        # Database type: WID vs SQL
        try {
            $config   = $WsusServer.GetConfiguration()
            $dbConn   = $config.MsSqlServerName
            if ($dbConn -match 'MICROSOFT##WID' -or [string]::IsNullOrEmpty($dbConn)) {
                $WsusDbType = StatusBadge 'Windows Internal Database (WID)' 'blue'
            } else {
                $WsusDbType = StatusBadge "SQL Server: $(HtmlEncode $dbConn)" 'blue'
            }
        } catch {}

        # Content directory and free disk space
        try {
            $contentDir  = $WsusServer.GetConfiguration().LocalContentCachePath
            $WsusContentPath = HtmlEncode $contentDir
            if (-not [string]::IsNullOrEmpty($contentDir)) {
                # Get just the drive letter for Get-PSDrive (no colon/backslash)
                $driveLetter = $contentDir.Substring(0, 1)
                $disk  = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue
                if ($null -eq $disk) {
                    # Fall back to WMI for UNC/non-standard drives
                    $deviceId = $driveLetter + ':'
                    $diskWmi = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$deviceId'" -ErrorAction SilentlyContinue
                    if ($diskWmi) {
                        $freeGB = [math]::Round($diskWmi.FreeSpace / 1GB, 2)
                        $totalGB = [math]::Round($diskWmi.Size / 1GB, 2)
                        $WsusContentFree = "$freeGB GB free of $totalGB GB"
                        $WsusContentBadge = if ($freeGB -lt $DiskCritGB) {
                            StatusBadge "CRITICAL: $freeGB GB free" 'red'
                        } elseif ($freeGB -lt $DiskWarnGB) {
                            StatusBadge "WARNING: $freeGB GB free" 'yellow'
                        } else {
                            StatusBadge "$freeGB GB free" 'green'
                        }
                    }
                } else {
                    $freeGB  = [math]::Round($disk.Free / 1GB, 2)
                    $totalGB = [math]::Round(($disk.Used + $disk.Free) / 1GB, 2)
                    $WsusContentFree = "$freeGB GB free of $totalGB GB"
                    if ($freeGB -lt $DiskCritGB) {
                        $CriticalFindings.Add("Section 1 - WSUS content drive critically low: $freeGB GB")
                        $WsusContentBadge = StatusBadge "CRITICAL: $freeGB GB free" 'red'
                    } elseif ($freeGB -lt $DiskWarnGB) {
                        $CriticalFindings.Add("Section 1 - WSUS content drive low: $freeGB GB")
                        $WsusContentBadge = StatusBadge "WARNING: $freeGB GB free" 'yellow'
                    } else {
                        $WsusContentBadge = StatusBadge "$freeGB GB free" 'green'
                    }
                }
            }
        } catch {
            $WsusContentFree = HtmlEncode "Error: $($_.Exception.Message)"
        }

        $rows1 = @(
            @('Server Name',        $WsusServerName),
            @('WSUS Version',       $WsusVersion),
            @('Port',               $WsusPort),
            @('SSL',                $WsusSsl),
            @('Database Type',      $WsusDbType),
            @('Content Directory',  $WsusContentPath),
            @('Content Drive Space',"$WsusContentFree $WsusContentBadge")
        )
        $Sec1Html = BuildKVTable $rows1
    } catch {
        $Sec1Html = "<p class='error'>Error reading WSUS server info: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("Section 1 - Error reading WSUS server info: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2  -  SYNCHRONIZATION STATUS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 2: Synchronization Status..."
$Sec2Html         = ''
$LastSyncBadgeKpi = StatusBadge 'N/A' 'grey'

if (-not $WsusModuleAvailable) {
    $Sec2Html = "<p class='warn'>WSUS not available - skipping synchronization check.</p>"
} else {
    try {
        $syncInfo = $WsusServer.GetSubscription()

        $lastSyncTime   = 'Never'
        $lastSyncResult = 'Unknown'
        $nextSyncTime   = 'Not scheduled'
        $totalUpdates   = 0
        $failedSyncs    = 0

        try { $lastSyncTime   = $syncInfo.LastSynchronizationTime.ToString('yyyy-MM-dd HH:mm:ss') } catch {}
        try {
            $rawResult      = $syncInfo.LastSynchronizationResult
            $lastSyncResult = $rawResult.ToString()
        } catch {}
        try {
            $nextSync = $syncInfo.NextScheduledSynchronizationTime
            $nextSyncTime = if ($nextSync -gt (Get-Date).AddYears(-1)) {
                $nextSync.ToString('yyyy-MM-dd HH:mm:ss')
            } else {
                'Not scheduled / manual'
            }
        } catch {}
        try { $totalUpdates = $WsusServer.GetUpdateCount() } catch {}
        try { $failedSyncs  = $syncInfo.NumberOfSyncFailures } catch {}

        $syncBadge = if ($lastSyncResult -match 'Succeeded') {
            StatusBadge 'Succeeded' 'green'
        } elseif ($lastSyncResult -match 'Failed') {
            $CriticalFindings.Add("Section 2 - Last WSUS synchronization FAILED")
            StatusBadge 'Failed' 'red'
        } else {
            StatusBadge $lastSyncResult 'yellow'
        }
        $LastSyncBadgeKpi = $syncBadge

        $failBadge = if ($failedSyncs -gt 0) {
            StatusBadge "$failedSyncs failed" 'red'
        } else {
            StatusBadge '0 failures' 'green'
        }

        $rows2 = @(
            @('Last Sync Time',      $(HtmlEncode $lastSyncTime)),
            @('Last Sync Result',    $syncBadge),
            @('Next Scheduled Sync', $(HtmlEncode $nextSyncTime)),
            @('Total Updates Synced',$(HtmlEncode $totalUpdates.ToString())),
            @('Sync Failures',       $failBadge)
        )
        $Sec2Html = BuildKVTable $rows2
    } catch {
        $Sec2Html = "<p class='error'>Error reading synchronization status: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("Section 2 - Sync status error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3  -  IIS APPLICATION POOL HEALTH
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 3: IIS Application Pool Health..."
$Sec3Html = ''
try {
    $webAdminAvailable = $false
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $webAdminAvailable = $true
    } catch {
        $Sec3Html = "<p class='warn'>WebAdministration module not available. IIS check skipped. $(HtmlEncode $_.Exception.Message)</p>"
    }

    if ($webAdminAvailable) {
        $poolRows = [System.Collections.Generic.List[string]]::new()

        # WSUSPool
        try {
            $pool = Get-Item 'IIS:\AppPools\WsusPool' -ErrorAction Stop
            $stateBadge = if ($pool.State -eq 'Started') {
                StatusBadge 'Started' 'green'
            } else {
                $CriticalFindings.Add("Section 3 - WSUSPool app pool is NOT started (state: $($pool.State))")
                StatusBadge $pool.State 'red'
            }
            $rapidFail = 'N/A'
            $recycle   = 'N/A'
            try { $rapidFail = HtmlEncode $pool.failure.rapidFailProtectionMaxCrashes.ToString() } catch {}
            try {
                $schedRecycle = $pool.recycling.periodicRestart.schedule.Collection
                $recycle = if ($schedRecycle -and $schedRecycle.Count -gt 0) {
                    ($schedRecycle | ForEach-Object { HtmlEncode $_.value }) -join ', '
                } else {
                    'No scheduled recycles'
                }
            } catch {}

            $poolRows.Add("<tr><td>WsusPool</td><td>$stateBadge</td><td>$rapidFail</td><td>$(HtmlEncode $recycle)</td></tr>")
        } catch {
            $poolRows.Add("<tr><td>WsusPool</td><td colspan='3'><em class='warn'>Not found or error: $(HtmlEncode $_.Exception.Message)</em></td></tr>")
            $CriticalFindings.Add("Section 3 - WSUSPool not found: $($_.Exception.Message)")
        }

        # WsusService binding (site)
        $bindingHtml = ''
        try {
            $site = Get-Website -Name 'WSUS Administration' -ErrorAction SilentlyContinue
            if ($null -eq $site) { $site = Get-Website | Where-Object { $_.Name -match 'WSUS' } | Select-Object -First 1 }
            if ($site) {
                $siteBadge = if ($site.State -eq 'Started') { StatusBadge 'Started' 'green' } else { StatusBadge $site.State 'red' }
                $bindings  = (Get-WebBinding -Name $site.Name -ErrorAction SilentlyContinue | ForEach-Object { HtmlEncode $_.bindingInformation }) -join ', '
                $bindingHtml = "<p><strong>WSUS IIS Site:</strong> $(HtmlEncode $site.Name) &mdash; $siteBadge &nbsp; Bindings: $bindings</p>"
            } else {
                $bindingHtml = "<p class='warn'>WSUS Administration IIS site not found.</p>"
            }
        } catch {
            $bindingHtml = "<p class='warn'>Could not query IIS site: $(HtmlEncode $_.Exception.Message)</p>"
        }

        $Sec3Html = @"
<div class='table-wrap'>
<table>
  <thead><tr><th>App Pool</th><th>State</th><th>Rapid Fail Max Crashes</th><th>Recycle Schedule</th></tr></thead>
  <tbody>$($poolRows -join '')</tbody>
</table>
</div>
$bindingHtml
"@
    }
} catch {
    $Sec3Html = "<p class='error'>IIS health check error: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 3 - IIS health check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4  -  DATABASE HEALTH
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 4: Database Health..."
$Sec4Html = ''
if (-not $WsusModuleAvailable) {
    $Sec4Html = "<p class='warn'>WSUS not available - skipping database check.</p>"
} else {
    try {
        $dbRows = [System.Collections.Generic.List[string]]::new()

        # Determine DB type
        $dbConnStr   = ''
        $isWID       = $true
        try {
            $cfg       = $WsusServer.GetConfiguration()
            $dbConnStr = $cfg.MsSqlServerName
            $isWID     = [string]::IsNullOrEmpty($dbConnStr) -or $dbConnStr -match 'MICROSOFT##WID'
        } catch {}

        if ($isWID) {
            # WID: check Windows Internal Database service
            $widSvc = Get-Service -Name 'MSSQL$MICROSOFT##WID' -ErrorAction SilentlyContinue
            if ($null -eq $widSvc) { $widSvc = Get-Service -Name 'MSSQL$MICROSOFT##SSEE' -ErrorAction SilentlyContinue }
            if ($widSvc) {
                $widBadge = if ($widSvc.Status -eq 'Running') { StatusBadge 'Running' 'green' } else {
                    $CriticalFindings.Add("Section 4 - WID service not running: $($widSvc.Status)")
                    StatusBadge $widSvc.Status 'red'
                }
                $dbRows.Add("<tr><td>Database Type</td><td>Windows Internal Database (WID)</td></tr>")
                $dbRows.Add("<tr><td>WID Service</td><td>$widBadge</td></tr>")
            } else {
                $dbRows.Add("<tr><td>Database Type</td><td>Windows Internal Database (WID)</td></tr>")
                $dbRows.Add("<tr><td>WID Service</td><td><em class='warn'>Service not found (MSSQL`$MICROSOFT##WID). May use a different instance name.</em></td></tr>")
            }
            # WID database file size
            try {
                $widDbPath = "$env:windir\WID\Data"
                if (Test-Path $widDbPath) {
                    $mdfFiles = Get-ChildItem -Path $widDbPath -Filter '*.mdf' -ErrorAction SilentlyContinue
                    if ($mdfFiles) {
                        $totalSize = ($mdfFiles | Measure-Object Length -Sum).Sum
                        $sizeGB    = [math]::Round($totalSize / 1GB, 3)
                        $dbRows.Add("<tr><td>WID Database Size</td><td>$(HtmlEncode $sizeGB) GB ($widDbPath)</td></tr>")
                    }
                }
            } catch {}
        } else {
            $dbRows.Add("<tr><td>Database Type</td><td>SQL Server: $(HtmlEncode $dbConnStr)</td></tr>")
            # Attempt SQL connection consistency-like check via WSUS API
            try {
                $updateScope  = New-Object Microsoft.UpdateServices.Administration.UpdateScope
                $updateCount  = $WsusServer.GetUpdateCount($updateScope)
                $dbRows.Add("<tr><td>DB Connectivity Check</td><td>$(StatusBadge 'OK - WSUS API query succeeded' 'green') (Updates in DB: $(HtmlEncode $updateCount.ToString()))</td></tr>")
            } catch {
                $dbRows.Add("<tr><td>DB Connectivity Check</td><td>$(StatusBadge 'Error' 'red') $(HtmlEncode $_.Exception.Message)</td></tr>")
                $CriticalFindings.Add("Section 4 - DB connectivity error: $($_.Exception.Message)")
            }
        }

        $Sec4Html = @"
<div class='table-wrap'>
<table class='kv-table'>
  <tbody>$($dbRows -join '')</tbody>
</table>
</div>
"@
    } catch {
        $Sec4Html = "<p class='error'>Database health check error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("Section 4 - Database health error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5  -  CLIENT REPORTING SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 5: Client Reporting Summary..."
$Sec5Html       = ''
$TotalClients   = 0
$PendingReboot  = 0

if (-not $WsusModuleAvailable) {
    $Sec5Html = "<p class='warn'>WSUS not available - skipping client reporting.</p>"
} else {
    try {
        $compScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope

        $allComputers  = @($WsusServer.GetComputerTargets($compScope))
        $TotalClients  = $allComputers.Count

        $now        = Get-Date
        $stale7     = 0
        $stale14    = 0
        $stale30    = 0
        $failedMany = 0
        $PendingReboot = 0

        foreach ($comp in $allComputers) {
            try {
                $lastReport = $null
                try { $lastReport = $comp.LastReportedStatusTime } catch {}
                if ($null -ne $lastReport -and $lastReport -gt (Get-Date '2000-01-01')) {
                    $daysSince = ($now - $lastReport).TotalDays
                    if ($daysSince -gt $StaleDay30) { $stale30++ }
                    elseif ($daysSince -gt $StaleDay14) { $stale14++ }
                    elseif ($daysSince -gt $StaleDay7)  { $stale7++  }
                }
            } catch {}

            try {
                $failCount = 0
                try { $failCount = $comp.GetUpdateInstallationSummary().FailedCount } catch {}
                if ($failCount -gt $FailedUpdateThreshold) { $failedMany++ }
            } catch {}

            try {
                $needsReboot = $false
                try {
                    $sum = $comp.GetUpdateInstallationSummary()
                    if ($sum.InstalledPendingRebootCount -gt 0) { $needsReboot = $true }
                } catch {}
                if ($needsReboot) { $PendingReboot++ }
            } catch {}
        }

        $rows5 = @(
            @('Total Registered Computers', $(HtmlEncode $TotalClients.ToString())),
            @("Computers with >&nbsp;$FailedUpdateThreshold Failed Updates",
              $(if ($failedMany -gt 0) { StatusBadge "$failedMany computers" 'red' } else { StatusBadge 'None' 'green' })),
            @("Not Reporting in $StaleDay7+ Days",
              $(if ($stale7 -gt 0) { StatusBadge "$stale7 computers" 'yellow' } else { StatusBadge 'None' 'green' })),
            @("Not Reporting in $StaleDay14+ Days",
              $(if ($stale14 -gt 0) { StatusBadge "$stale14 computers" 'yellow' } else { StatusBadge 'None' 'green' })),
            @("Not Reporting in $StaleDay30+ Days",
              $(if ($stale30 -gt 0) { StatusBadge "$stale30 computers" 'red' } else { StatusBadge 'None' 'green' })),
            @('Computers Pending Reboot',
              $(if ($PendingReboot -gt 0) { StatusBadge "$PendingReboot computers" 'yellow' } else { StatusBadge 'None' 'green' }))
        )
        $Sec5Html = BuildKVTable $rows5
    } catch {
        $Sec5Html = "<p class='error'>Client reporting error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("Section 5 - Client reporting error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6  -  UPDATE COMPLIANCE OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 6: Update Compliance Overview..."
$Sec6Html          = ''
$CompliancePct     = 0
$TotalUpdatesCount = 0

if (-not $WsusModuleAvailable) {
    $Sec6Html = "<p class='warn'>WSUS not available - skipping compliance overview.</p>"
} else {
    try {
        # Classifications
        $classMap = @{}
        try {
            $WsusServer.GetUpdateClassifications() | ForEach-Object {
                $classMap[$_.Id] = $_.Title
            }
        } catch {}

        $updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
        $allUpdates  = @($WsusServer.GetUpdates($updateScope))
        $TotalUpdatesCount = $allUpdates.Count

        $classCount  = @{}
        $approved    = 0
        $notApproved = 0
        $declined    = 0
        $expired     = 0
        $superseded  = 0

        foreach ($upd in $allUpdates) {
            try {
                $cTitle = 'Unknown'
                try { $cTitle = $upd.UpdateClassificationTitle } catch {}
                if ([string]::IsNullOrEmpty($cTitle)) { $cTitle = 'Unknown' }
                if ($classCount.ContainsKey($cTitle)) { $classCount[$cTitle]++ } else { $classCount[$cTitle] = 1 }
            } catch {}

            try {
                $action    = 'Unknown'
                $isDeclined = $false
                try { $action     = $upd.ApprovalAction.ToString() } catch {}
                try { $isDeclined = [bool]$upd.IsDeclined } catch {}

                if ($isDeclined) {
                    $declined++
                } else {
                    switch ($action) {
                        'Install'     { $approved++ }
                        'NotApproved' { $notApproved++ }
                        'Uninstall'   { }
                        default       { $notApproved++ }
                    }
                }
                try { if ($upd.IsSuperseded) { $superseded++ } } catch {}
                try { if ($upd.IsExpired)    { $expired++    } } catch {}
            } catch {}
        }

        # Compliance % (computers fully compliant = 0 needed updates)
        $compliantCount = 0
        try {
            if ($TotalClients -gt 0) {
                $compScope2 = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
                foreach ($comp in @($WsusServer.GetComputerTargets($compScope2))) {
                    try {
                        $sum = $comp.GetUpdateInstallationSummary()
                        if ($sum.NotInstalledCount -eq 0 -and $sum.FailedCount -eq 0 -and $sum.DownloadedCount -eq 0) {
                            $compliantCount++
                        }
                    } catch {}
                }
                $CompliancePct = [math]::Round(($compliantCount / $TotalClients) * 100, 1)
            }
        } catch {}

        # Classification table
        $classRows = foreach ($kvp in ($classCount.GetEnumerator() | Sort-Object Value -Descending)) {
            "<tr><td>$(HtmlEncode $kvp.Key)</td><td>$(HtmlEncode $kvp.Value.ToString())</td></tr>"
        }

        $pctColor   = if ($CompliancePct -ge 90) { '#3fb950' } elseif ($CompliancePct -ge 70) { '#d29922' } else { '#f85149' }
        $pctBarHtml = @"
<div style='margin:12px 0;'>
  <div style='display:flex;align-items:center;gap:12px;'>
    <span style='font-weight:600;min-width:180px;'>Overall Patch Compliance</span>
    <div style='flex:1;background:#21262d;border-radius:8px;height:18px;overflow:hidden;'>
      <div style='width:$($CompliancePct)%;background:$pctColor;height:100%;border-radius:8px;transition:width .5s;'></div>
    </div>
    <span style='font-weight:700;color:$pctColor;min-width:50px;'>$($CompliancePct)%</span>
  </div>
</div>
"@

        $Sec6Html = @"
$pctBarHtml
<div style='display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px;'>
  <div class='stat-card' style='flex:1 1 120px;text-align:center;border:1px solid var(--border);border-radius:8px;padding:12px;'>
    <div class='stat-count c-blue'>$(HtmlEncode $TotalUpdatesCount.ToString())</div><div class='stat-label'>Total Updates</div>
  </div>
  <div class='stat-card' style='flex:1 1 120px;text-align:center;border:1px solid var(--border);border-radius:8px;padding:12px;'>
    <div class='stat-count c-green'>$(HtmlEncode $approved.ToString())</div><div class='stat-label'>Approved</div>
  </div>
  <div class='stat-card' style='flex:1 1 120px;text-align:center;border:1px solid var(--border);border-radius:8px;padding:12px;'>
    <div class='stat-count c-yellow'>$(HtmlEncode $notApproved.ToString())</div><div class='stat-label'>Not Approved</div>
  </div>
  <div class='stat-card' style='flex:1 1 120px;text-align:center;border:1px solid var(--border);border-radius:8px;padding:12px;'>
    <div class='stat-count c-muted'>$(HtmlEncode $declined.ToString())</div><div class='stat-label'>Declined</div>
  </div>
  <div class='stat-card' style='flex:1 1 120px;text-align:center;border:1px solid var(--border);border-radius:8px;padding:12px;'>
    <div class='stat-count c-muted'>$(HtmlEncode $superseded.ToString())</div><div class='stat-label'>Superseded</div>
  </div>
  <div class='stat-card' style='flex:1 1 120px;text-align:center;border:1px solid var(--border);border-radius:8px;padding:12px;'>
    <div class='stat-count c-muted'>$(HtmlEncode $expired.ToString())</div><div class='stat-label'>Expired</div>
  </div>
</div>
<h3>Updates by Classification</h3>
<div class='table-wrap'>
<table>
  <thead><tr><th>Classification</th><th>Count</th></tr></thead>
  <tbody>$($classRows -join '')</tbody>
</table>
</div>
"@
    } catch {
        $Sec6Html = "<p class='error'>Compliance overview error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("Section 6 - Compliance overview error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7  -  TOP 10 NON-COMPLIANT COMPUTERS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 7: Top 10 Non-Compliant Computers..."
$Sec7Html = ''

if (-not $WsusModuleAvailable) {
    $Sec7Html = "<p class='warn'>WSUS not available - skipping non-compliant computers check.</p>"
} else {
    try {
        $compScope3  = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
        $allComps    = @($WsusServer.GetComputerTargets($compScope3))
        $compData    = [System.Collections.Generic.List[PSObject]]::new()

        foreach ($comp in $allComps) {
            try {
                $needed = 0
                try {
                    $sum    = $comp.GetUpdateInstallationSummary()
                    $needed = [int]$sum.NotInstalledCount + [int]$sum.FailedCount + [int]$sum.DownloadedCount
                } catch {}
                $compData.Add([PSCustomObject]@{
                    Name   = $comp.FullDomainName
                    Needed = $needed
                })
            } catch {}
        }

        $top10 = $compData | Sort-Object Needed -Descending | Select-Object -First 10

        if ($top10.Count -eq 0) {
            $Sec7Html = "<p class='info'>$(StatusBadge 'All computers are fully compliant' 'green')</p>"
        } else {
            $t7Rows = foreach ($c in $top10) {
                $neededBadge = if ($c.Needed -gt 50) {
                    StatusBadge $c.Needed.ToString() 'red'
                } elseif ($c.Needed -gt 10) {
                    StatusBadge $c.Needed.ToString() 'yellow'
                } else {
                    StatusBadge $c.Needed.ToString() 'grey'
                }
                "<tr><td>$(HtmlEncode $c.Name)</td><td>$neededBadge</td></tr>"
            }
            $Sec7Html = @"
<div class='table-wrap'>
<table>
  <thead><tr><th>Computer Name</th><th>Needed Updates</th></tr></thead>
  <tbody>$($t7Rows -join '')</tbody>
</table>
</div>
"@
        }
    } catch {
        $Sec7Html = "<p class='error'>Non-compliant computers error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("Section 7 - Non-compliant computers error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 8  -  WSUS SERVICES STATUS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 8: WSUS Services Status..."
$Sec8Html = ''
try {
    $serviceNames = @(
        @{ Name = 'WsusService';            Display = 'WSUS Service (WsusService)' },
        @{ Name = 'W3Svc';                  Display = 'IIS (W3Svc)' },
        @{ Name = 'MSSQL$MICROSOFT##WID';   Display = 'Windows Internal Database (WID)' },
        @{ Name = 'MSSQLSERVER';            Display = 'SQL Server (MSSQLSERVER)' },
        @{ Name = 'UpdateServicesDbServer'; Display = 'WSUS DB Server (UpdateServicesDbServer)' }
    )

    $svcRows = [System.Collections.Generic.List[string]]::new()
    foreach ($svcDef in $serviceNames) {
        $svc = Get-Service -Name $svcDef.Name -ErrorAction SilentlyContinue
        if ($null -eq $svc) {
            $svcRows.Add("<tr><td>$(HtmlEncode $svcDef.Display)</td><td>$(StatusBadge 'Not Installed' 'grey')</td><td>N/A</td></tr>")
        } else {
            $badge = if ($svc.Status -eq 'Running') {
                StatusBadge 'Running' 'green'
            } else {
                if ($svcDef.Name -in @('WsusService','W3Svc')) {
                    $CriticalFindings.Add("Section 8 - Critical service not running: $($svcDef.Display) ($($svc.Status))")
                }
                StatusBadge $svc.Status 'red'
            }
            $svcRows.Add("<tr><td>$(HtmlEncode $svcDef.Display)</td><td>$badge</td><td>$(HtmlEncode $svc.StartType.ToString())</td></tr>")
        }
    }

    $Sec8Html = @"
<div class='table-wrap'>
<table>
  <thead><tr><th>Service</th><th>Status</th><th>Start Type</th></tr></thead>
  <tbody>$($svcRows -join '')</tbody>
</table>
</div>
"@
} catch {
    $Sec8Html = "<p class='error'>Service status error: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 8 - Service status error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 9  -  DISK SPACE CHECK
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 9: Disk Space Check..."
$Sec9Html     = ''
$DiskSpaceKpi = 'N/A'
try {
    $contentDirForDisk = ''
    if ($WsusModuleAvailable -and $null -ne $WsusServer) {
        try { $contentDirForDisk = $WsusServer.GetConfiguration().LocalContentCachePath } catch {}
    }

    $diskRows = [System.Collections.Generic.List[string]]::new()
    $allDrives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue

    foreach ($d in $allDrives) {
        $freeGB  = [math]::Round($d.FreeSpace / 1GB, 2)
        $totalGB = [math]::Round($d.Size / 1GB, 2)
        $usedPct = if ($d.Size -gt 0) { [math]::Round((($d.Size - $d.FreeSpace) / $d.Size) * 100, 1) } else { 0 }

        $isContentDrive = (-not [string]::IsNullOrEmpty($contentDirForDisk)) -and
                          ($contentDirForDisk.StartsWith($d.DeviceID, [System.StringComparison]::OrdinalIgnoreCase))

        $badge = if ($freeGB -lt $DiskCritGB) {
            $CriticalFindings.Add("Section 9 - Drive $($d.DeviceID) critically low: $freeGB GB free")
            StatusBadge "CRITICAL: $freeGB GB free" 'red'
        } elseif ($freeGB -lt $DiskWarnGB) {
            $CriticalFindings.Add("Section 9 - Drive $($d.DeviceID) low disk space: $freeGB GB free")
            StatusBadge "WARNING: $freeGB GB free" 'yellow'
        } else {
            StatusBadge "$freeGB GB free" 'green'
        }

        $wsusLabel = if ($isContentDrive) { "<strong> (WSUS Content)</strong>" } else { "" }
        $pctBar = "<div style='width:100%;background:#21262d;border-radius:4px;height:10px;'><div style='width:$usedPct%;background:$(if($freeGB -lt $DiskCritGB){'#f85149'}elseif($freeGB -lt $DiskWarnGB){'#d29922'}else{'#3fb950'});height:100%;border-radius:4px;'></div></div>"

        $diskRows.Add("<tr><td>$(HtmlEncode $d.DeviceID)$wsusLabel</td><td>$(HtmlEncode $totalGB.ToString()) GB</td><td>$(HtmlEncode $freeGB.ToString()) GB</td><td>$usedPct%</td><td>$pctBar</td><td>$badge</td></tr>")

        if ($isContentDrive) { $DiskSpaceKpi = "$freeGB GB free" }
    }

    if ($diskRows.Count -eq 0) {
        $Sec9Html = "<p class='warn'>No fixed disk drives found.</p>"
    } else {
        $Sec9Html = @"
<div class='table-wrap'>
<table>
  <thead><tr><th>Drive</th><th>Total</th><th>Free</th><th>Used %</th><th>Usage Bar</th><th>Status</th></tr></thead>
  <tbody>$($diskRows -join '')</tbody>
</table>
</div>
<p class='info'>Warn below $DiskWarnGB GB &nbsp;|&nbsp; Critical below $DiskCritGB GB</p>
"@
    }
} catch {
    $Sec9Html = "<p class='error'>Disk space check error: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 9 - Disk space check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 10  -  EVENT LOG CHECK
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 10: Event Log Check..."
$Sec10Html = ''
try {
    $eventRows = [System.Collections.Generic.List[string]]::new()

    # Primary WSUS server log; fall back to Application if not available
    $logNames = @('Windows Server Update Services', 'Application')

    foreach ($logName in $logNames) {
        try {
            $events = Get-WinEvent -LogName $logName -MaxEvents 10 `
                        -FilterXPath "*[System[Level=1 or Level=2]]" `
                        -ErrorAction Stop
            foreach ($ev in $events) {
                $lvl       = switch ($ev.Level) { 1 { 'Critical' } 2 { 'Error' } default { $ev.LevelDisplayName } }
                $rowClass  = switch ($ev.Level) { 1 { 'row-critical' } 2 { 'row-error' } default { '' } }
                $lvlBadge  = switch ($ev.Level) { 1 { StatusBadge 'Critical' 'red' } 2 { StatusBadge 'Error' 'red' } default { StatusBadge $lvl 'grey' } }
                $timeStr   = HtmlEncode $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
                $rawMsg    = if ($null -ne $ev.Message) { $ev.Message } else { '' }
                $msgShort  = HtmlEncode ($rawMsg -replace '\r?\n', ' ' | ForEach-Object { if ($_.Length -gt 200) { $_.Substring(0,200) + '...' } else { $_ } })
                $eventRows.Add("<tr class='$rowClass'><td>$timeStr</td><td>$lvlBadge</td><td>$(HtmlEncode $ev.Id.ToString())</td><td>$(HtmlEncode $ev.ProviderName)</td><td>$msgShort</td></tr>")
            }
            break
        } catch {
            # log not found or no events matching filter - try next
        }
    }

    # Also scan System log for WSUS-related errors
    try {
        $sysEvents = Get-WinEvent -LogName 'System' -MaxEvents 200 -ErrorAction SilentlyContinue |
                     Where-Object { ($_.Level -le 2) -and ($_.ProviderName -match 'WSUS|Update') } |
                     Select-Object -First 5
        foreach ($ev in $sysEvents) {
            $lvlBadge  = if ($ev.Level -eq 1) { StatusBadge 'Critical' 'red' } else { StatusBadge 'Error' 'red' }
            $timeStr   = HtmlEncode $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            $rawMsg    = if ($null -ne $ev.Message) { $ev.Message } else { '' }
            $msgShort  = HtmlEncode ($rawMsg -replace '\r?\n', ' ' | ForEach-Object { if ($_.Length -gt 200) { $_.Substring(0,200) + '...' } else { $_ } })
            $eventRows.Add("<tr class='row-error'><td>$timeStr</td><td>$lvlBadge</td><td>$(HtmlEncode $ev.Id.ToString())</td><td>$(HtmlEncode $ev.ProviderName)</td><td>$msgShort</td></tr>")
        }
    } catch {}

    if ($eventRows.Count -eq 0) {
        $Sec10Html = "<p class='info'>$(StatusBadge 'No Critical/Error events found in WSUS event logs' 'green')</p>"
    } else {
        $Sec10Html = @"
<div class='table-wrap'>
<table class='event-table'>
  <thead><tr><th>Time</th><th>Level</th><th>Event ID</th><th>Source</th><th>Message</th></tr></thead>
  <tbody>$($eventRows -join '')</tbody>
</table>
</div>
"@
    }
} catch {
    $Sec10Html = "<p class='error'>Event log check error: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 10 - Event log check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 11  -  CLEANUP RECOMMENDATIONS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 11: Cleanup Recommendations..."
$Sec11Html      = ''
$CleanupRows    = [System.Collections.Generic.List[string]]::new()
$CleanupWarning = $false

try {
    # 1. Check registry for last cleanup run date
    $lastCleanupDate = $null
    $cleanupSource   = 'Registry'
    try {
        $regPath = 'HKLM:\SOFTWARE\Microsoft\Update Services\Server\Setup'
        $regKey  = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        if ($regKey -and $regKey.PSObject.Properties['LastCleanupWizardRun']) {
            $lastCleanupDate = [datetime]$regKey.LastCleanupWizardRun
        }
    } catch {}

    # 2. Fall back: check WSUS API if available
    if ($null -eq $lastCleanupDate -and $WsusModuleAvailable) {
        try {
            $cleanupMgr  = $WsusServer.GetCleanupManager()
            # GetCleanupStatus is not always available but try
            $cleanupStatus = $cleanupMgr.GetCleanupStatus()
            if ($cleanupStatus) {
                # No direct last-run-date property; use existence as indicator
                $cleanupSource = 'WSUS API'
            }
        } catch {}
    }

    # 3. Check WSUS cleanup via event log (Event 12032 = cleanup completed)
    if ($null -eq $lastCleanupDate) {
        try {
            $cleanupEvent = Get-WinEvent -LogName 'Application' -ErrorAction SilentlyContinue |
                            Where-Object { $_.Id -in @(12032,12033) -and $_.ProviderName -match 'WSUS' } |
                            Sort-Object TimeCreated -Descending |
                            Select-Object -First 1
            if ($cleanupEvent) {
                $lastCleanupDate = $cleanupEvent.TimeCreated
                $cleanupSource   = "Event Log (ID $($cleanupEvent.Id))"
            }
        } catch {}
    }

    if ($null -ne $lastCleanupDate) {
        $daysSinceCleanup = (Get-Date - $lastCleanupDate).TotalDays
        $cleanupBadge = if ($daysSinceCleanup -gt $CleanupStaleDays) {
            $CleanupWarning = $true
            $CriticalFindings.Add("Section 11 - WSUS cleanup not run in $([math]::Round($daysSinceCleanup,0)) days")
            StatusBadge "OVERDUE ($([math]::Round($daysSinceCleanup,0)) days ago)" 'red'
        } else {
            StatusBadge "$([math]::Round($daysSinceCleanup,0)) days ago" 'green'
        }
        $CleanupRows.Add("<tr><td>Last Cleanup Run</td><td>$(HtmlEncode $lastCleanupDate.ToString('yyyy-MM-dd HH:mm:ss'))</td><td>$cleanupBadge</td><td>$(HtmlEncode $cleanupSource)</td></tr>")
    } else {
        $CleanupWarning = $true
        $CriticalFindings.Add("Section 11 - WSUS cleanup last run date could not be determined")
        $CleanupRows.Add("<tr><td>Last Cleanup Run</td><td colspan='2'>$(StatusBadge 'Unknown - could not determine' 'yellow')</td><td>N/A</td></tr>")
    }

    # Cleanup recommendations
    $recommendations = @(
        @('Decline Expired Updates',     'Run WSUS Cleanup Wizard or: Invoke-WsusServerCleanup -DeclineExpiredUpdates'),
        @('Decline Superseded Updates',  'Run WSUS Cleanup Wizard or: Invoke-WsusServerCleanup -DeclineSupersededUpdates'),
        @('Delete Obsolete Updates',     'Run WSUS Cleanup Wizard or: Invoke-WsusServerCleanup -DeleteObsoleteUpdates -CleanupObsoleteComputers'),
        @('Compress Update Files',       'Run WSUS Cleanup Wizard or: Invoke-WsusServerCleanup -CompressUpdates'),
        @('Cleanup WID/SQL Database',    'Run: Invoke-WsusServerCleanup -CleanupObsoleteComputers -CleanupUnneededContentFiles')
    )

    foreach ($rec in $recommendations) {
        $CleanupRows.Add("<tr><td>$(HtmlEncode $rec[0])</td><td colspan='2'><code>$(HtmlEncode $rec[1])</code></td><td>Recommendation</td></tr>")
    }

    $Sec11Html = @"
<div class='table-wrap'>
<table>
  <thead><tr><th>Check</th><th>Details</th><th>Status</th><th>Source</th></tr></thead>
  <tbody>$($CleanupRows -join '')</tbody>
</table>
</div>
<p class='info' style='margin-top:8px;'>Flag threshold: cleanup not run within <strong>$CleanupStaleDays days</strong>. Run <code>Invoke-WsusServerCleanup</code> to perform cleanup.</p>
"@
} catch {
    $Sec11Html = "<p class='error'>Cleanup recommendations error: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 11 - Cleanup check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SUMMARY KPI VALUES
# ═══════════════════════════════════════════════════════════════════════════════
$CritCount    = $CriticalFindings.Count
$ReportDate   = (Get-Date).ToString('dddd, dd MMMM yyyy HH:mm:ss')
$EndTime      = Get-Date
$Duration     = ($EndTime - $StartTime).ToString('hh\:mm\:ss')

$LastSyncKpiHtml = if ($WsusModuleAvailable) { $LastSyncBadgeKpi } else { StatusBadge 'N/A' 'grey' }
$DiskSpaceKpiHtml = if ($DiskSpaceKpi -ne 'N/A') {
    if ($DiskSpaceKpi -match 'critical' -or ($DiskSpaceKpi -match '(\d+\.?\d*)\s*GB' -and [double]$Matches[1] -lt $DiskCritGB)) {
        StatusBadge $DiskSpaceKpi 'red'
    } elseif ($DiskSpaceKpi -match '(\d+\.?\d*)\s*GB' -and [double]$Matches[1] -lt $DiskWarnGB) {
        StatusBadge $DiskSpaceKpi 'yellow'
    } else {
        StatusBadge $DiskSpaceKpi 'green'
    }
} else { StatusBadge 'N/A' 'grey' }

Write-Progress2 "Building HTML report..."

# ── COMPANY LOGO HTML ─────────────────────────────────────────────────────────
$LogoHtml = ''
if (-not [string]::IsNullOrWhiteSpace($CompanyLogoURL)) {
    $logoImg = "<img src='$(HtmlEncode $CompanyLogoURL)' alt='Company Logo' style='max-height:60px;vertical-align:middle;'>"
    if (-not [string]::IsNullOrWhiteSpace($CompanyWebsite)) {
        $LogoHtml = "<a href='$(HtmlEncode $CompanyWebsite)' target='_blank'>$logoImg</a>"
    } else {
        $LogoHtml = $logoImg
    }
    $LogoHtml = "<div class='logo-wrap'>$LogoHtml</div>"
}

$AuthorLink = if (-not [string]::IsNullOrWhiteSpace($CompanyWebsite)) {
    "<a href='$(HtmlEncode $CompanyWebsite)' target='_blank' style='color:var(--link);'>$(HtmlEncode $AuthorName)</a>"
} else {
    HtmlEncode $AuthorName
}

# ── OVERALL STATUS FOR HEADER ─────────────────────────────────────────────────
$OverallStatusBadge = if ($CritCount -gt 0) { StatusBadge 'CRITICAL' 'red' } else { StatusBadge 'HEALTHY' 'green' }

# ═══════════════════════════════════════════════════════════════════════════════
# BUILD FULL HTML
# ═══════════════════════════════════════════════════════════════════════════════
$HtmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>WSUS Health Check Report - $(HtmlEncode $WsusServerName)</title>
<style>
/* ── CSS VARIABLES ── */
:root {
  --bg:       #0d1117;
  --card:     #161b22;
  --border:   #30363d;
  --text:     #c9d1d9;
  --muted:    #8b949e;
  --link:     #58a6ff;
  --head-bg:  #010409;
  --th-bg:    #21262d;
  --tr-alt:   #1c2128;
  --pre-bg:   #0d1117;
  --green:    #238636;
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
  --link:     #0969da;
  --head-bg:  #f6f8fa;
  --th-bg:    #eaeef2;
  --tr-alt:   #f6f8fa;
  --pre-bg:   #f6f8fa;
}

/* ── RESET & BASE ── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: var(--bg);
  color: var(--text);
  font-size: 14px;
  line-height: 1.6;
  transition: background .3s, color .3s;
}
a { color: var(--link); text-decoration: none; }
a:hover { text-decoration: underline; }
h3 { margin: 1rem 0 .5rem; font-size: 1rem; color: var(--text); }
h4 { margin: .75rem 0 .4rem; font-size: .9rem; color: var(--muted); }
p  { margin-bottom: .5rem; }
ul { margin: .4rem 0 .4rem 1.4rem; }
code { font-family: 'SFMono-Regular', Consolas, monospace; font-size: 12px;
       background: var(--th-bg); padding: 2px 6px; border-radius: 4px; }

/* ── LAYOUT ── */
.page-wrap  { max-width: 1400px; margin: 0 auto; padding: 0 16px 40px; }
.header     { background: var(--head-bg); border-bottom: 1px solid var(--border);
              padding: 16px 24px; display: flex; align-items: center;
              justify-content: space-between; flex-wrap: wrap; gap: 12px; }
.header-left  { display: flex; align-items: center; gap: 16px; }
.logo-wrap img { max-height: 52px; }
.header-title h1 { font-size: 1.4rem; color: var(--text); }
.header-title p  { font-size: .8rem; color: var(--muted); margin:0; }

/* ── THEME TOGGLE ── */
.theme-toggle { cursor: pointer; background: var(--card); border: 1px solid var(--border);
                color: var(--text); border-radius: 20px; padding: 6px 14px;
                font-size: 12px; display: flex; align-items: center; gap: 6px; }
.theme-toggle:hover { background: var(--th-bg); }

/* ── SUMMARY BAR ── */
.summary-bar { display: flex; flex-wrap: wrap; gap: 12px;
               background: var(--card); border: 1px solid var(--border);
               border-radius: 8px; padding: 16px 20px; margin: 20px 0; }
.stat-card   { flex: 1 1 140px; text-align: center; }
.stat-count  { font-size: 2rem; font-weight: 700; line-height: 1; }
.stat-label  { font-size: .75rem; color: var(--muted); margin-top: 4px; }
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
.sec-num   { background: var(--blue); color: #fff; font-size: .7rem; font-weight: 700;
             width: 22px; height: 22px; border-radius: 50%; display: flex;
             align-items: center; justify-content: center; flex-shrink: 0; }
.sec-title { font-weight: 600; font-size: .95rem; flex: 1; }
.section-body { padding: 16px 18px; border-top: 1px solid var(--border); }

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
.td-label   { font-weight: 600; white-space: nowrap; width: 220px; color: var(--muted); }

/* ── EVENT ROW COLORS ── */
.row-critical { background: rgba(218,54,51,.15) !important; }
.row-error    { background: rgba(218,54,51,.08) !important; }
.row-warning  { background: rgba(210,153,34,.12) !important; }

/* ── BADGES ── */
.badge { display: inline-block; font-size: .7rem; font-weight: 600; padding: 2px 8px;
         border-radius: 20px; color: #fff; white-space: nowrap; }

/* ── EVENT LOG TABLE ── */
.event-table td:nth-child(5) {
  max-width: 320px;
  word-break: break-word;
  white-space: normal;
}

/* ── MESSAGE CLASSES ── */
.error { color: #f85149; padding: 8px 12px; background: rgba(248,81,73,.1);
         border-left: 3px solid #f85149; border-radius: 4px; }
.warn  { color: #d29922; }
.info  { color: var(--muted); }

/* ── FOOTER ── */
.footer { border-top: 1px solid var(--border); padding: 20px 0;
          margin-top: 24px; color: var(--muted); font-size: .8rem;
          display: flex; justify-content: space-between; flex-wrap: wrap; gap: 8px; }

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
      <h1>&#x1F6E1; WSUS Health Check</h1>
      <p>$(HtmlEncode $WsusServerName) &nbsp;|&nbsp; Generated: $(HtmlEncode $ReportDate)</p>
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
    <div class="stat-count c-blue">$(HtmlEncode $TotalClients.ToString())</div>
    <div class="stat-label">Total Clients</div>
  </div>
  <div class="stat-card">
    <div class="stat-count $(if ($CompliancePct -ge 90) { 'c-green' } elseif ($CompliancePct -ge 70) { 'c-yellow' } else { 'c-red' })">$($CompliancePct)%</div>
    <div class="stat-label">Compliant</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:.9rem;padding-top:10px;">$LastSyncKpiHtml</div>
    <div class="stat-label">Last Sync Status</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:.9rem;padding-top:10px;">$DiskSpaceKpiHtml</div>
    <div class="stat-label">Content Disk Space</div>
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
$(if ($CritCount -gt 0) {
    $cfRows = ($CriticalFindings | ForEach-Object { "<li>$(HtmlEncode $_)</li>" }) -join ''
    "<div style='background:rgba(218,54,51,.12);border:1px solid #da3633;border-radius:8px;padding:14px 18px;margin-bottom:16px;'><strong style='color:#f85149;'>&#x26A0; $CritCount Critical Finding(s) Detected</strong><ul style='margin:.6rem 0 0 1.2rem;color:#f85149;'>$cfRows</ul></div>"
})

<!-- 12 SECTIONS -->
$(BuildSection 0  'Server Details'                  $Sec0Html  ($Sec0Html  -match 'error|Error')  $true)
$(BuildSection 1  'WSUS Server Info'                $Sec1Html  ($Sec1Html  -match 'error|Error')  $true)
$(BuildSection 2  'Synchronization Status'          $Sec2Html  ($Sec2Html  -match 'error|Error')  $true)
$(BuildSection 3  'IIS Application Pool Health'     $Sec3Html  ($Sec3Html  -match 'error|Error')  $false)
$(BuildSection 4  'Database Health'                 $Sec4Html  ($Sec4Html  -match 'error|Error')  $false)
$(BuildSection 5  'Client Reporting Summary'        $Sec5Html  ($Sec5Html  -match 'error|Error')  $false)
$(BuildSection 6  'Update Compliance Overview'      $Sec6Html  ($Sec6Html  -match 'error|Error')  $false)
$(BuildSection 7  'Top 10 Non-Compliant Computers'  $Sec7Html  ($Sec7Html  -match 'error|Error')  $false)
$(BuildSection 8  'WSUS Services Status'            $Sec8Html  ($Sec8Html  -match 'error|Error')  $false)
$(BuildSection 9  'Disk Space Check'                $Sec9Html  ($Sec9Html  -match 'error|Error')  $false)
$(BuildSection 10 'Event Log Check'                 $Sec10Html ($Sec10Html -match 'error|Error')  $false)
$(BuildSection 11 'Cleanup Recommendations'         $Sec11Html ($CleanupWarning -or ($Sec11Html -match 'error|Error')) $false)

<!-- FOOTER -->
<div class="footer">
  <div>
    <strong>WSUS Health Check v$ScriptVersion</strong> &nbsp;|&nbsp;
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
# WRITE REPORT FILE
# ═══════════════════════════════════════════════════════════════════════════════
try {
    [System.IO.File]::WriteAllText($ReportFile, $HtmlReport, [System.Text.Encoding]::UTF8)
    Write-Host ""
    Write-Host "  [OK] Report saved to: $ReportFile" -ForegroundColor Green
} catch {
    Write-Warning "Failed to write report: $_"
}

Write-Host ""
Write-Host "===============================================================" -ForegroundColor DarkCyan
Write-Host "  WSUS Health Check complete.  Duration: $Duration" -ForegroundColor DarkCyan
Write-Host "  Report  : $ReportFile" -ForegroundColor Yellow
$statusLabel = if ($CritCount -gt 0) { 'CRITICAL' } else { 'HEALTHY' }
$statusColor = if ($CritCount -gt 0) { 'Red' } else { 'Green' }
Write-Host "  Status  : $statusLabel ($CritCount critical findings)" -ForegroundColor $statusColor
Write-Host "===============================================================" -ForegroundColor DarkCyan
Write-Host ""
