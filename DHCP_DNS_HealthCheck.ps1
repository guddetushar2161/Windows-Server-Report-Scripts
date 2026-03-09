#Requires -Version 5.1
<#
.SYNOPSIS
    DHCP and DNS Server Health Check Script

.DESCRIPTION
    Performs a comprehensive health check of both the DHCP Server role and the
    DNS Server role on a Windows Server and exports all results to a single
    self-contained HTML dashboard file.

    Each role section gracefully shows "Role Not Installed" if the corresponding
    Windows feature / module is not present on the target server, rather than
    crashing. Every section is individually wrapped in Try/Catch.

    DHCP checks:
      1.  DHCP Service Status       (service state, AD authorization)
      2.  Scope Inventory           (all scopes with state, mask, range, lease duration)
      3.  Scope Utilization         (total/used/free IPs, % bars; warn >80%, crit >95%)
      4.  DHCP Failover             (partner, mode, state, % split; flag fault states)
      5.  DHCP Reservations         (count per scope)
      6.  Lease Statistics          (total active, offered, declined, NAKs)
      7.  DHCP Audit Log            (enabled/disabled, log path valid)
      8.  DHCP Database             (file location, size, backup path, last backup)
      9.  DHCP Event Log            (last 10 Critical/Error/Warning from DhcpAdminEvents)

    DNS checks:
      10. DNS Service Status        (service state, last start time)
      11. Zone Inventory            (all zones: type, dynamic update, DNSSEC, SOA serial)
      12. Zone Health               (SOA + NS records present, not paused)
      13. Forwarders                (list + port-53 reachability test)
      14. Root Hints                (populated count)
      15. DNS Resolution Tests      (local domain, external domain, reverse lookup)
      16. Aging & Scavenging        (enabled/disabled, refresh/no-refresh intervals)
      17. Conditional Forwarders    (list + reachability test)
      18. DNSSEC Status             (signed zones)
      19. DNS Cache                 (approximate count via Get-DnsServerCache)
      20. DNS Event Log             (last 10 Critical/Error from DNS Server log)

.NOTES
    Version    : 1.0.0
    Author     : Tushar Gudde
    Requires   : PowerShell 5.1+, Local Administrator rights
    Compatible : Windows Server 2012 R2, 2016, 2019, 2022
    Companion  : DHCP_DNS_HealthCheck_EmailAlert.ps1
#>

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
$CompanyLogoURL = ''
$CompanyWebsite = 'https://tushargudde.tech'
$AuthorName     = 'Tushar Gudde'

$ScopeUtilWarnPct  = 80
$ScopeUtilCritPct  = 95

$ForwarderTestPort    = 53
$ForwarderTestTimeoutSec = 3   # seconds

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
$ReportStamp = "DHCP_DNS_Health_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
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
        green  = '#3fb950'; yellow = '#d29922'; red    = '#f85149'
        blue   = '#58a6ff'; amber  = '#e3a019'; grey   = '#484f58'
        rose   = '#e05c7a'; pink   = '#db61a2'; purple = '#a371f7'
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

function UtilBar {
    param([double]$pct)
    $color = if ($pct -ge $ScopeUtilCritPct) { '#f85149' }
             elseif ($pct -ge $ScopeUtilWarnPct) { '#d29922' }
             else { '#3fb950' }
    $width = [math]::Min($pct, 100)
    return "<div class='disk-bar-outer' style='width:100%;'>" +
           "<div class='disk-bar-inner' style='width:${width}%;background:$color;'></div></div>" +
           "<span style='font-size:.75rem;color:$color;font-weight:600;'>$([math]::Round($pct,1))%</span>"
}

function BuildSection {
    param(
        [int]$num,
        [string]$title,
        [string]$body,
        [bool]$hasError = $false,
        [bool]$open     = $false,
        [string]$prefix = ''
    )
    $openAttr  = if ($open) { ' open' } else { '' }
    $indicator = if ($hasError) {
        "<span style='color:#f85149;'>&#x2716;</span>"
    } else {
        "<span style='color:#db61a2;'>&#x2714;</span>"
    }
    $numLabel = if ($prefix) { "$prefix.$num" } else { $num.ToString() }
    return @"
<details class='section-card'$openAttr>
  <summary class='section-summary'>
    <span class='sec-arrow'>&#9654;</span>
    <span class='sec-num'>$numLabel</span>
    <span class='sec-title'>$(HtmlEncode $title)</span>
    $indicator
  </summary>
  <div class='section-body'>
$body
  </div>
</details>
"@
}

function Test-Tcp {
    param([string]$Hostname, [int]$Port, [int]$TimeoutSec = 3)
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($Hostname, $Port, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne($TimeoutSec * 1000, $false)
        if ($ok -and $tcp.Connected) {
            $tcp.EndConnect($ar)
            $tcp.Close()
            return $true
        }
        $tcp.Close()
        return $false
    } catch {
        return $false
    }
}

# ── CRITICAL FINDINGS LIST ────────────────────────────────────────────────────
$CriticalFindings = [System.Collections.Generic.List[string]]::new()

# ── DEFAULT KPI VARS ─────────────────────────────────────────────────────────
$ServerHostname    = $env:COMPUTERNAME
$TotalScopes       = 0
$CriticalScopes    = 0
$TotalZones        = 0
$FailedForwarders  = 0
$IsDHCPInstalled   = $false
$IsDNSInstalled    = $false

Write-Host ""
Write-Host "+==============================================================+" -ForegroundColor DarkMagenta
Write-Host "|   DHCP + DNS Server Health Check  v$ScriptVersion                  |" -ForegroundColor DarkMagenta
Write-Host "|   Server: $($env:COMPUTERNAME)" -ForegroundColor DarkMagenta
Write-Host "+==============================================================+" -ForegroundColor DarkMagenta
Write-Host ""

# ── DETECT ROLES ─────────────────────────────────────────────────────────────
Write-Progress2 "Detecting installed roles..."

# DHCP detection
try {
    $dhcpSvc = Get-Service -Name DHCPServer -ErrorAction Stop
    $IsDHCPInstalled = $true
    Write-Progress2 "DHCP Server service (DHCPServer) detected."
} catch {
    $IsDHCPInstalled = $false
    Write-Progress2 "DHCP Server service not found."
}

# Try importing the DhcpServer module
$DhcpModuleOk = $false
if ($IsDHCPInstalled) {
    try {
        Import-Module DhcpServer -ErrorAction Stop
        $DhcpModuleOk = $true
    } catch {
        Write-Progress2 "DhcpServer module not available  -  some DHCP checks will use WMI/registry fallback."
    }
}

# DNS detection
try {
    $dnsSvc = Get-Service -Name DNS -ErrorAction Stop
    $IsDNSInstalled = $true
    Write-Progress2 "DNS Server service (DNS) detected."
} catch {
    $IsDNSInstalled = $false
    Write-Progress2 "DNS Server service not found."
}

$DnsModuleOk = $false
if ($IsDNSInstalled) {
    try {
        Import-Module DnsServer -ErrorAction Stop
        $DnsModuleOk = $true
    } catch {
        Write-Progress2 "DnsServer module not available  -  limited DNS checks will be performed."
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DHCP SECTION 1  -  DHCP SERVICE STATUS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DHCP-1: Service Status..."
$D1Html = ''
if (-not $IsDHCPInstalled) {
    $D1Html = "<div class='not-installed-card'><span class='ni-icon'>&#x1F4CB;</span><div><strong>DHCP Server Role Not Installed</strong><p>The DHCPServer service was not detected on this server.</p></div></div>"
} else {
    try {
        $dhcpSvc2     = Get-Service -Name DHCPServer -ErrorAction Stop
        $dhcpStatus   = $dhcpSvc2.Status.ToString()
        $dhcpStart    = $dhcpSvc2.StartType.ToString()
        $dhcpColor    = if ($dhcpStatus -eq 'Running') { 'green' } else { 'red' }
        if ($dhcpStatus -ne 'Running') { $CriticalFindings.Add("DHCP-1 - DHCPServer service is $dhcpStatus") }

        # AD Authorization
        $authStatus  = 'Not Checked'
        $authColor   = 'grey'
        if ($DhcpModuleOk) {
            try {
                $authServers = @(Get-DhcpServerInDC -ErrorAction Stop)
                $selfAuth    = $authServers | Where-Object { $_.DnsName -like "*$ServerHostname*" -or $null -ne $_.IPAddress }
                if ($selfAuth -and $selfAuth.Count -gt 0) {
                    $authStatus = 'Authorized in AD'
                    $authColor  = 'green'
                } else {
                    $authStatus = 'Not Authorized in AD'
                    $authColor  = 'red'
                    $CriticalFindings.Add("DHCP-1 - DHCP server is NOT authorized in Active Directory")
                }
            } catch {
                $authStatus = 'Not in AD domain or error'
                $authColor  = 'yellow'
            }
        }

        $rows = @(
            @('Service Name',    (HtmlEncode 'DHCPServer')),
            @('Service Status',  (StatusBadge $dhcpStatus $dhcpColor)),
            @('Start Type',      (HtmlEncode $dhcpStart)),
            @('AD Authorization',(StatusBadge $authStatus $authColor)),
            @('Module Available',(StatusBadge $(if ($DhcpModuleOk) { 'Yes' } else { 'No' }) $(if ($DhcpModuleOk) { 'green' } else { 'yellow' })))
        )
        $D1Html = BuildKVTable $rows
    } catch {
        $D1Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DHCP-1 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DHCP SECTION 2  -  SCOPE INVENTORY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DHCP-2: Scope Inventory..."
$D2Html = ''
$AllScopes = @()
if (-not $IsDHCPInstalled) {
    $D2Html = "<p class='warn'>DHCP not installed  -  skipped.</p>"
} elseif (-not $DhcpModuleOk) {
    $D2Html = "<p class='warn'>DhcpServer module not available  -  scope inventory requires the DHCP management tools.</p>"
} else {
    try {
        $AllScopes = @(Get-DhcpServerv4Scope -ErrorAction Stop)
        $TotalScopes = $AllScopes.Count
        if ($TotalScopes -eq 0) {
            $D2Html = "<p class='info'>No DHCP v4 scopes found on this server.</p>"
        } else {
            $D2Html  = "<div class='table-wrap'><table><thead><tr>"
            $D2Html += "<th>Scope ID</th><th>Name</th><th>Subnet Mask</th><th>Start IP</th><th>End IP</th><th>State</th><th>Lease Duration</th>"
            $D2Html += "</tr></thead><tbody>"
            foreach ($sc in $AllScopes) {
                $stColor = if ($sc.State -eq 'Active') { 'green' } else { 'yellow' }
                $lease   = if ($null -ne $sc.LeaseDuration) { $sc.LeaseDuration.ToString() } else { 'N/A' }
                $D2Html += "<tr>"
                $D2Html += "<td><code>$(HtmlEncode $sc.ScopeId.ToString())</code></td>"
                $D2Html += "<td>$(HtmlEncode $sc.Name)</td>"
                $D2Html += "<td>$(HtmlEncode $sc.SubnetMask.ToString())</td>"
                $D2Html += "<td>$(HtmlEncode $sc.StartRange.ToString())</td>"
                $D2Html += "<td>$(HtmlEncode $sc.EndRange.ToString())</td>"
                $D2Html += "<td>$(StatusBadge $sc.State.ToString() $stColor)</td>"
                $D2Html += "<td>$(HtmlEncode $lease)</td>"
                $D2Html += "</tr>"
            }
            $D2Html += "</tbody></table></div>"
        }
    } catch {
        $D2Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DHCP-2 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DHCP SECTION 3  -  SCOPE UTILIZATION
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DHCP-3: Scope Utilization..."
$D3Html = ''
if (-not $IsDHCPInstalled) {
    $D3Html = "<p class='warn'>DHCP not installed  -  skipped.</p>"
} elseif (-not $DhcpModuleOk -or $AllScopes.Count -eq 0) {
    $D3Html = "<p class='warn'>DhcpServer module not available or no scopes found  -  utilization check skipped.</p>"
} else {
    try {
        $D3Html  = "<div class='table-wrap'><table><thead><tr>"
        $D3Html += "<th>Scope ID</th><th>Name</th><th>Total IPs</th><th>In Use</th><th>Free</th><th>Utilization</th><th>Status</th>"
        $D3Html += "</tr></thead><tbody>"
        foreach ($sc in $AllScopes) {
            try {
                $stats    = Get-DhcpServerv4ScopeStatistics -ScopeId $sc.ScopeId -ErrorAction Stop
                $total    = $stats.Total
                $inUse    = $stats.InUse
                $free     = $stats.Free
                $pct      = if ($total -gt 0) { [math]::Round(($inUse / $total) * 100, 1) } else { 0 }
                $stLabel  = if ($pct -ge $ScopeUtilCritPct) { 'CRITICAL' }
                            elseif ($pct -ge $ScopeUtilWarnPct) { 'Warning' }
                            else { 'OK' }
                $stColor  = if ($pct -ge $ScopeUtilCritPct) { 'red' }
                            elseif ($pct -ge $ScopeUtilWarnPct) { 'yellow' }
                            else { 'green' }
                if ($pct -ge $ScopeUtilCritPct) {
                    $CriticalScopes++
                    $CriticalFindings.Add("DHCP-3 - Scope $($sc.ScopeId) utilization is ${pct}% (CRITICAL)")
                } elseif ($pct -ge $ScopeUtilWarnPct) {
                    $CriticalFindings.Add("DHCP-3 - Scope $($sc.ScopeId) utilization is ${pct}% (Warning)")
                }
                $D3Html += "<tr>"
                $D3Html += "<td><code>$(HtmlEncode $sc.ScopeId.ToString())</code></td>"
                $D3Html += "<td>$(HtmlEncode $sc.Name)</td>"
                $D3Html += "<td>$total</td><td>$inUse</td><td>$free</td>"
                $D3Html += "<td>$(UtilBar $pct)</td>"
                $D3Html += "<td>$(StatusBadge $stLabel $stColor)</td>"
                $D3Html += "</tr>"
            } catch {
                $D3Html += "<tr><td><code>$(HtmlEncode $sc.ScopeId.ToString())</code></td><td>$(HtmlEncode $sc.Name)</td>"
                $D3Html += "<td colspan='5'><span class='error'>Error: $(HtmlEncode $_.Exception.Message)</span></td></tr>"
            }
        }
        $D3Html += "</tbody></table></div>"
    } catch {
        $D3Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DHCP-3 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DHCP SECTION 4  -  FAILOVER
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DHCP-4: Failover..."
$D4Html = ''
if (-not $IsDHCPInstalled -or -not $DhcpModuleOk) {
    $D4Html = "<p class='warn'>$(if (-not $IsDHCPInstalled) { 'DHCP not installed' } else { 'DhcpServer module not available' })  -  skipped.</p>"
} else {
    try {
        $failovers = @(Get-DhcpServerv4Failover -ErrorAction Stop)
        if ($failovers.Count -eq 0) {
            $D4Html = "<p class='info'>No DHCP failover relationships configured on this server.</p>"
        } else {
            $D4Html  = "<div class='table-wrap'><table><thead><tr>"
            $D4Html += "<th>Name</th><th>Partner Server</th><th>Mode</th><th>State</th><th>Split %</th><th>Status</th>"
            $D4Html += "</tr></thead><tbody>"
            foreach ($fo in $failovers) {
                $foState   = $fo.State.ToString()
                $faultStates = @('CommunicationInterrupted','ConflictDone','PartnerDown','RecoverWait','Recover')
                $isFault   = $faultStates -contains $foState
                $foColor   = if ($foState -eq 'Normal') { 'green' } elseif ($isFault) { 'red' } else { 'yellow' }
                if ($isFault) { $CriticalFindings.Add("DHCP-4 - Failover '$($fo.Name)' is in fault state: $foState") }
                $split     = if ($fo.LoadBalancePercent) { "$($fo.LoadBalancePercent)%" } else { 'N/A' }
                $D4Html += "<tr>"
                $D4Html += "<td>$(HtmlEncode $fo.Name)</td>"
                $D4Html += "<td>$(HtmlEncode $fo.PartnerServer)</td>"
                $D4Html += "<td>$(HtmlEncode $fo.Mode.ToString())</td>"
                $D4Html += "<td>$(StatusBadge $foState $foColor)</td>"
                $D4Html += "<td>$(HtmlEncode $split)</td>"
                $D4Html += "<td>$(if ($isFault) { StatusBadge 'FAULT' 'red' } else { StatusBadge 'OK' 'green' })</td>"
                $D4Html += "</tr>"
            }
            $D4Html += "</tbody></table></div>"
        }
    } catch {
        $D4Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DHCP-4 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DHCP SECTION 5  -  RESERVATIONS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DHCP-5: Reservations..."
$D5Html = ''
if (-not $IsDHCPInstalled -or -not $DhcpModuleOk -or $AllScopes.Count -eq 0) {
    $D5Html = "<p class='warn'>$(if (-not $IsDHCPInstalled) { 'DHCP not installed' } elseif (-not $DhcpModuleOk) { 'DhcpServer module not available' } else { 'No scopes found' })  -  skipped.</p>"
} else {
    try {
        $D5Html  = "<div class='table-wrap'><table><thead><tr>"
        $D5Html += "<th>Scope ID</th><th>Scope Name</th><th>Reservation Count</th>"
        $D5Html += "</tr></thead><tbody>"
        $totalRes = 0
        foreach ($sc in $AllScopes) {
            try {
                $res   = @(Get-DhcpServerv4Reservation -ScopeId $sc.ScopeId -ErrorAction Stop)
                $count = $res.Count
                $totalRes += $count
                $D5Html += "<tr><td><code>$(HtmlEncode $sc.ScopeId.ToString())</code></td>"
                $D5Html += "<td>$(HtmlEncode $sc.Name)</td>"
                $D5Html += "<td>$count</td></tr>"
            } catch {
                $D5Html += "<tr><td><code>$(HtmlEncode $sc.ScopeId.ToString())</code></td>"
                $D5Html += "<td>$(HtmlEncode $sc.Name)</td>"
                $D5Html += "<td><span class='warn'>Error: $(HtmlEncode $_.Exception.Message)</span></td></tr>"
            }
        }
        $D5Html += "</tbody></table></div>"
        $D5Html += "<p style='margin-top:8px;font-size:.85rem;'>Total reservations across all scopes: <strong>$totalRes</strong></p>"
    } catch {
        $D5Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DHCP-5 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DHCP SECTION 6  -  LEASE STATISTICS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DHCP-6: Lease Statistics..."
$D6Html = ''
if (-not $IsDHCPInstalled -or -not $DhcpModuleOk) {
    $D6Html = "<p class='warn'>$(if (-not $IsDHCPInstalled) { 'DHCP not installed' } else { 'DhcpServer module not available' })  -  skipped.</p>"
} else {
    try {
        $stats = Get-DhcpServerv4Statistics -ErrorAction Stop
        $rows6 = @(
            @('Total Scopes',         (HtmlEncode $stats.TotalScopes.ToString())),
            @('Total Addresses',      (HtmlEncode $stats.TotalAddresses.ToString())),
            @('Addresses In Use',     (HtmlEncode $stats.AddressesInUse.ToString())),
            @('Addresses Available',  (HtmlEncode $stats.AddressesAvailable.ToString())),
            @('Total Leases',         (HtmlEncode $stats.TotalLeases.ToString())),
            @('Pending Offers',       (HtmlEncode $stats.PendingOffers.ToString())),
            @('Requests',             (HtmlEncode $stats.Requests.ToString())),
            @('Declines',             (HtmlEncode $stats.Declines.ToString())),
            @('NAKs',                 (HtmlEncode $stats.Naks.ToString())),
            @('Releases',             (HtmlEncode $stats.Releases.ToString())),
            @('Server Start Time',    (HtmlEncode $(if ($null -ne $stats.ServerStartTime) { $stats.ServerStartTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' })))
        )
        $D6Html = BuildKVTable $rows6
    } catch {
        $D6Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DHCP-6 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DHCP SECTION 7  -  AUDIT LOG
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DHCP-7: Audit Log..."
$D7Html = ''
if (-not $IsDHCPInstalled -or -not $DhcpModuleOk) {
    $D7Html = "<p class='warn'>$(if (-not $IsDHCPInstalled) { 'DHCP not installed' } else { 'DhcpServer module not available' })  -  skipped.</p>"
} else {
    try {
        $auditCfg  = Get-DhcpServerAuditLog -ErrorAction Stop
        $auditEnabled = $auditCfg.Enable
        $auditPath    = $auditCfg.Path
        $pathExists   = if (-not [string]::IsNullOrEmpty($auditPath)) { Test-Path $auditPath } else { $false }
        $enableColor  = if ($auditEnabled) { 'green' } else { 'yellow' }
        $pathColor    = if ($pathExists) { 'green' } else { 'red' }
        if (-not $auditEnabled) { $CriticalFindings.Add("DHCP-7 - DHCP audit logging is disabled") }
        if ($auditEnabled -and -not $pathExists) { $CriticalFindings.Add("DHCP-7 - DHCP audit log path does not exist: $auditPath") }
        $rows7 = @(
            @('Audit Logging Enabled', (StatusBadge $(if ($auditEnabled) { 'Enabled' } else { 'Disabled' }) $enableColor)),
            @('Log Path',              (HtmlEncode $auditPath)),
            @('Log Path Exists',       (StatusBadge $(if ($pathExists) { 'Yes' } else { 'No' }) $pathColor))
        )
        $D7Html = BuildKVTable $rows7
    } catch {
        $D7Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DHCP-7 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DHCP SECTION 8  -  DATABASE
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DHCP-8: Database..."
$D8Html = ''
if (-not $IsDHCPInstalled -or -not $DhcpModuleOk) {
    $D8Html = "<p class='warn'>$(if (-not $IsDHCPInstalled) { 'DHCP not installed' } else { 'DhcpServer module not available' })  -  skipped.</p>"
} else {
    try {
        $dbCfg  = Get-DhcpServerDatabase -ErrorAction Stop
        $dbFile = Join-Path $dbCfg.FileName 'dhcp.mdb'
        if (-not (Test-Path $dbFile)) { $dbFile = $dbCfg.FileName }
        $dbSizeMB = 'N/A'
        if (Test-Path $dbFile) {
            $dbSizeMB = [math]::Round((Get-Item $dbFile).Length / 1MB, 2)
        }
        $backupPath     = $dbCfg.BackupPath
        $backupInterval = if ($null -ne $dbCfg.BackupInterval) { "$($dbCfg.BackupInterval) minutes" } else { 'N/A' }
        $backupExists   = if (-not [string]::IsNullOrEmpty($backupPath)) { Test-Path $backupPath } else { $false }
        $rows8 = @(
            @('Database Path',      (HtmlEncode $dbCfg.FileName)),
            @('Database Size',      (HtmlEncode "$dbSizeMB MB")),
            @('Backup Path',        (HtmlEncode $backupPath)),
            @('Backup Path Exists', (StatusBadge $(if ($backupExists) { 'Yes' } else { 'No' }) $(if ($backupExists) { 'green' } else { 'yellow' }))),
            @('Backup Interval',    (HtmlEncode $backupInterval)),
            @('Cleanup Interval',   (HtmlEncode $(if ($null -ne $dbCfg.CleanupInterval) { "$($dbCfg.CleanupInterval) minutes" } else { 'N/A' })))
        )
        $D8Html = BuildKVTable $rows8
    } catch {
        $D8Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DHCP-8 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DHCP SECTION 9  -  EVENT LOG
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DHCP-9: Event Log..."
$D9Html = ''
try {
    if (-not $IsDHCPInstalled) {
        $D9Html = "<p class='warn'>DHCP not installed  -  skipped.</p>"
    } else {
        $dhcpEvents = [System.Collections.Generic.List[object]]::new()
        $since48h   = (Get-Date).AddHours(-48)

        $dhcpLogs = @('DhcpAdminEvents', 'Microsoft-Windows-Dhcp-Server/Operational', 'System')
        foreach ($log in $dhcpLogs) {
            try {
                $filter = @{ LogName = $log; Level = @(1,2,3); StartTime = $since48h }
                $evts   = @(Get-WinEvent -FilterHashtable $filter -MaxEvents 30 -ErrorAction Stop |
                    Where-Object { $log -ne 'System' -or $_.ProviderName -match 'DhcpServer|Dhcp' })
                foreach ($e in $evts) { $dhcpEvents.Add($e) }
            } catch {}
        }

        $topDhcpEvents = $dhcpEvents | Sort-Object TimeCreated -Descending | Select-Object -Unique -First 10

        if ($topDhcpEvents.Count -eq 0) {
            $D9Html = "<p style='color:#3fb950;'>&#x2714; No Critical/Error/Warning events found in DHCP logs (last 48 hours).</p>"
        } else {
            $D9Html  = "<div class='table-wrap'><table><thead><tr>"
            $D9Html += "<th>Time</th><th>Level</th><th>Event ID</th><th>Source</th><th>Message</th>"
            $D9Html += "</tr></thead><tbody>"
            foreach ($ev in $topDhcpEvents) {
                $timeStr  = if ($null -ne $ev.TimeCreated) { $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
                $level    = if ($null -ne $ev.LevelDisplayName) { $ev.LevelDisplayName } else { 'Unknown' }
                $lvlColor = switch ($level) { 'Critical' { 'red' } 'Error' { 'red' } 'Warning' { 'yellow' } default { 'grey' } }
                $msgRaw   = if ($null -ne $ev.Message) { $ev.Message } else { '' }
                $msgShort = if ($msgRaw.Length -gt 180) { $msgRaw.Substring(0,180) + '...' } else { $msgRaw }
                $D9Html += "<tr><td style='white-space:nowrap;'>$(HtmlEncode $timeStr)</td>"
                $D9Html += "<td>$(StatusBadge $level $lvlColor)</td>"
                $D9Html += "<td>$($ev.Id)</td>"
                $D9Html += "<td>$(HtmlEncode $(if ($null -ne $ev.ProviderName) { $ev.ProviderName } else { 'N/A' }))</td>"
                $D9Html += "<td>$(HtmlEncode $msgShort)</td></tr>"
            }
            $D9Html += "</tbody></table></div>"
        }
    }
} catch {
    $D9Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("DHCP-9 error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# DNS SECTION 10  -  SERVICE STATUS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DNS-10: Service Status..."
$N10Html = ''
if (-not $IsDNSInstalled) {
    $N10Html = "<div class='not-installed-card'><span class='ni-icon'>&#x1F4CB;</span><div><strong>DNS Server Role Not Installed</strong><p>The DNS service was not detected on this server.</p></div></div>"
} else {
    try {
        $dnsSvc2      = Get-Service -Name DNS -ErrorAction Stop
        $dnsStatus    = $dnsSvc2.Status.ToString()
        $dnsStartType = $dnsSvc2.StartType.ToString()
        $dnsColor     = if ($dnsStatus -eq 'Running') { 'green' } else { 'red' }
        if ($dnsStatus -ne 'Running') { $CriticalFindings.Add("DNS-10 - DNS service is $dnsStatus") }

        $rows10 = @(
            @('Service Name',    (HtmlEncode 'DNS')),
            @('Service Status',  (StatusBadge $dnsStatus $dnsColor)),
            @('Start Type',      (HtmlEncode $dnsStartType)),
            @('Module Available',(StatusBadge $(if ($DnsModuleOk) { 'Yes' } else { 'No' }) $(if ($DnsModuleOk) { 'green' } else { 'yellow' })))
        )
        $N10Html = BuildKVTable $rows10
    } catch {
        $N10Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DNS-10 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DNS SECTION 11  -  ZONE INVENTORY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DNS-11: Zone Inventory..."
$N11Html  = ''
$AllZones = @()
if (-not $IsDNSInstalled) {
    $N11Html = "<p class='warn'>DNS not installed  -  skipped.</p>"
} elseif (-not $DnsModuleOk) {
    $N11Html = "<p class='warn'>DnsServer module not available  -  zone inventory skipped.</p>"
} else {
    try {
        $AllZones  = @(Get-DnsServerZone -ErrorAction Stop)
        $TotalZones = $AllZones.Count
        if ($TotalZones -eq 0) {
            $N11Html = "<p class='info'>No DNS zones found.</p>"
        } else {
            $N11Html  = "<div class='table-wrap'><table><thead><tr>"
            $N11Html += "<th>Zone Name</th><th>Zone Type</th><th>Replication</th><th>Dynamic Update</th><th>AD Integrated</th><th>Paused</th>"
            $N11Html += "</tr></thead><tbody>"
            foreach ($z in $AllZones) {
                $zType   = $z.ZoneType.ToString()
                $zDyn    = if ($null -ne $z.DynamicUpdate) { $z.DynamicUpdate.ToString() } else { 'N/A' }
                $zAdInt  = if ($null -ne $z.IsDsIntegrated) { if ($z.IsDsIntegrated) { 'Yes' } else { 'No' } } else { 'N/A' }
                $zPaused = if ($null -ne $z.IsPaused) { if ($z.IsPaused) { 'Yes' } else { 'No' } } else { 'N/A' }
                $pauseColor = if ($z.IsPaused) { 'red' } else { 'green' }
                $replScope  = if ($null -ne $z.ReplicationScope) { $z.ReplicationScope.ToString() } else { 'N/A' }
                $N11Html += "<tr>"
                $N11Html += "<td>$(HtmlEncode $z.ZoneName)</td>"
                $N11Html += "<td>$(StatusBadge $zType 'blue')</td>"
                $N11Html += "<td>$(HtmlEncode $replScope)</td>"
                $N11Html += "<td>$(HtmlEncode $zDyn)</td>"
                $N11Html += "<td>$(HtmlEncode $zAdInt)</td>"
                $N11Html += "<td>$(StatusBadge $zPaused $pauseColor)</td>"
                $N11Html += "</tr>"
            }
            $N11Html += "</tbody></table></div>"
        }
    } catch {
        $N11Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DNS-11 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DNS SECTION 12  -  ZONE HEALTH
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DNS-12: Zone Health..."
$N12Html = ''
if (-not $IsDNSInstalled -or -not $DnsModuleOk -or $AllZones.Count -eq 0) {
    $N12Html = "<p class='warn'>$(if (-not $IsDNSInstalled) { 'DNS not installed' } elseif (-not $DnsModuleOk) { 'DnsServer module not available' } else { 'No zones found' })  -  skipped.</p>"
} else {
    try {
        $primaryZones = @($AllZones | Where-Object { $_.ZoneType -eq 'Primary' -and $_.ZoneName -ne '.' })
        if ($primaryZones.Count -eq 0) {
            $N12Html = "<p class='info'>No primary zones found for health check.</p>"
        } else {
            $N12Html  = "<div class='table-wrap'><table><thead><tr>"
            $N12Html += "<th>Zone</th><th>Paused</th><th>SOA Present</th><th>NS Present</th><th>Status</th>"
            $N12Html += "</tr></thead><tbody>"
            foreach ($z in $primaryZones) {
                $soaOk = $false; $nsOk = $false
                try {
                    $soaRec = Get-DnsServerResourceRecord -ZoneName $z.ZoneName -RRType Soa -ErrorAction Stop
                    $soaOk  = ($null -ne $soaRec -and @($soaRec).Count -gt 0)
                } catch {}
                try {
                    $nsRec = Get-DnsServerResourceRecord -ZoneName $z.ZoneName -RRType Ns -ErrorAction Stop
                    $nsOk  = ($null -ne $nsRec -and @($nsRec).Count -gt 0)
                } catch {}
                $paused   = if ($null -ne $z.IsPaused -and $z.IsPaused) { 'Yes' } else { 'No' }
                $pauseClr = if ($paused -eq 'Yes') { 'red' } else { 'green' }
                $allGood  = $soaOk -and $nsOk -and ($paused -eq 'No')
                $statusBadge = if ($allGood) { StatusBadge 'OK' 'green' } else { StatusBadge 'Issue' 'red' }
                if (-not $allGood) { $CriticalFindings.Add("DNS-12 - Zone '$($z.ZoneName)': SOA=$soaOk, NS=$nsOk, Paused=$paused") }
                $N12Html += "<tr>"
                $N12Html += "<td>$(HtmlEncode $z.ZoneName)</td>"
                $N12Html += "<td>$(StatusBadge $paused $pauseClr)</td>"
                $N12Html += "<td>$(StatusBadge $(if ($soaOk) { 'Yes' } else { 'No' }) $(if ($soaOk) { 'green' } else { 'red' }))</td>"
                $N12Html += "<td>$(StatusBadge $(if ($nsOk) { 'Yes' } else { 'No' }) $(if ($nsOk) { 'green' } else { 'red' }))</td>"
                $N12Html += "<td>$statusBadge</td>"
                $N12Html += "</tr>"
            }
            $N12Html += "</tbody></table></div>"
        }
    } catch {
        $N12Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DNS-12 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DNS SECTION 13  -  FORWARDERS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DNS-13: Forwarders..."
$N13Html = ''
if (-not $IsDNSInstalled -or -not $DnsModuleOk) {
    $N13Html = "<p class='warn'>$(if (-not $IsDNSInstalled) { 'DNS not installed' } else { 'DnsServer module not available' })  -  skipped.</p>"
} else {
    try {
        $fwdCfg  = Get-DnsServerForwarder -ErrorAction Stop
        $fwdIPs  = @($fwdCfg.IPAddress)
        if ($fwdIPs.Count -eq 0) {
            $N13Html = "<p class='info'>No forwarders configured (root hints may be in use).</p>"
        } else {
            $N13Html  = "<div class='table-wrap'><table><thead><tr>"
            $N13Html += "<th>Forwarder IP</th><th>Port 53 Reachable</th><th>Status</th>"
            $N13Html += "</tr></thead><tbody>"
            foreach ($ip in $fwdIPs) {
                $ipStr  = $ip.ToString()
                $reach  = Test-Tcp -Hostname $ipStr -Port $ForwarderTestPort -TimeoutSec $ForwarderTestTimeoutSec
                $rLabel = if ($reach) { 'Reachable &#x2714;' } else { 'Unreachable &#x2716;' }
                $rColor = if ($reach) { 'green' } else { 'red' }
                if (-not $reach) {
                    $FailedForwarders++
                    $CriticalFindings.Add("DNS-13 - Forwarder $ipStr is unreachable on port 53")
                }
                $N13Html += "<tr><td><code>$(HtmlEncode $ipStr)</code></td>"
                $N13Html += "<td>$(StatusBadge $rLabel $rColor)</td>"
                $N13Html += "<td>$(StatusBadge $(if ($reach) { 'OK' } else { 'FAILED' }) $rColor)</td></tr>"
            }
            $N13Html += "</tbody></table></div>"
            $N13Html += "<p style='margin-top:8px;font-size:.85rem;'>Use Root Hints: $(HtmlEncode $fwdCfg.UseRootHint.ToString())</p>"
        }
    } catch {
        $N13Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DNS-13 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DNS SECTION 14  -  ROOT HINTS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DNS-14: Root Hints..."
$N14Html = ''
if (-not $IsDNSInstalled -or -not $DnsModuleOk) {
    $N14Html = "<p class='warn'>$(if (-not $IsDNSInstalled) { 'DNS not installed' } else { 'DnsServer module not available' })  -  skipped.</p>"
} else {
    try {
        $rootHints = @(Get-DnsServerRootHint -ErrorAction Stop)
        $rhCount   = $rootHints.Count
        $rhColor   = if ($rhCount -ge 5) { 'green' } elseif ($rhCount -gt 0) { 'yellow' } else { 'red' }
        if ($rhCount -eq 0) { $CriticalFindings.Add("DNS-14 - No root hints configured and forwarders may also be absent") }
        $rows14 = @(
            @('Root Hint Servers', (StatusBadge "$rhCount server(s)" $rhColor))
        )
        $N14Html = BuildKVTable $rows14
        if ($rhCount -gt 0) {
            $N14Html += "<div class='table-wrap' style='margin-top:10px;'><table><thead><tr><th>Name Server</th><th>IP Addresses</th></tr></thead><tbody>"
            foreach ($rh in ($rootHints | Select-Object -First 15)) {
                $nsName = if ($null -ne $rh.NameServer) { $rh.NameServer.RecordData.NameServer } else { 'N/A' }
                $ips    = if ($null -ne $rh.IPAddress) {
                    ($rh.IPAddress | ForEach-Object { $_.RecordData.IPv4Address.ToString() }) -join ', '
                } else { 'N/A' }
                $N14Html += "<tr><td>$(HtmlEncode $nsName)</td><td>$(HtmlEncode $ips)</td></tr>"
            }
            $N14Html += "</tbody></table></div>"
        }
    } catch {
        $N14Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DNS-14 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DNS SECTION 15  -  RESOLUTION TESTS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DNS-15: Resolution Tests..."
$N15Html = ''
if (-not $IsDNSInstalled) {
    $N15Html = "<p class='warn'>DNS not installed  -  skipped.</p>"
} else {
    try {
        # Get local domain
        $localDomain = $env:USERDNSDOMAIN
        if ([string]::IsNullOrEmpty($localDomain)) {
            try { $localDomain = (Get-WmiObject Win32_ComputerSystem -ErrorAction Stop).Domain } catch {}
        }

        $tests = [System.Collections.Generic.List[object]]::new()

        # 1 - Local domain
        if (-not [string]::IsNullOrEmpty($localDomain)) {
            try {
                $resolved = [System.Net.Dns]::GetHostAddresses($localDomain)
                $tests.Add([pscustomobject]@{ Test='Local Domain'; Target=$localDomain; Result='Resolved'; IPs=($resolved | ForEach-Object { $_.ToString() }) -join ', '; Status='OK'; Color='green' })
            } catch {
                $tests.Add([pscustomobject]@{ Test='Local Domain'; Target=$localDomain; Result='Failed'; IPs=$_.Exception.Message; Status='FAILED'; Color='red' })
                $CriticalFindings.Add("DNS-15 - Failed to resolve local domain: $localDomain")
            }
        } else {
            $tests.Add([pscustomobject]@{ Test='Local Domain'; Target='(not joined)'; Result='Skipped'; IPs='N/A'; Status='N/A'; Color='grey' })
        }

        # 2 - External domain
        try {
            $extResolved = [System.Net.Dns]::GetHostAddresses('microsoft.com')
            $tests.Add([pscustomobject]@{ Test='External Domain'; Target='microsoft.com'; Result='Resolved'; IPs=($extResolved | Select-Object -First 3 | ForEach-Object { $_.ToString() }) -join ', '; Status='OK'; Color='green' })
        } catch {
            $tests.Add([pscustomobject]@{ Test='External Domain'; Target='microsoft.com'; Result='Failed'; IPs='Forwarder may be unreachable'; Status='Warning'; Color='yellow' })
        }

        # 3 - Reverse lookup for server own IP
        try {
            $ownIP  = (([System.Net.Dns]::GetHostEntry($env:COMPUTERNAME)).AddressList |
                Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).ToString()
            $revHost = [System.Net.Dns]::GetHostEntry($ownIP).HostName
            $tests.Add([pscustomobject]@{ Test='Reverse Lookup (self)'; Target=$ownIP; Result='Resolved'; IPs=$revHost; Status='OK'; Color='green' })
        } catch {
            $tests.Add([pscustomobject]@{ Test='Reverse Lookup (self)'; Target=$env:COMPUTERNAME; Result='Failed'; IPs=$_.Exception.Message; Status='Warning'; Color='yellow' })
        }

        $N15Html  = "<div class='table-wrap'><table><thead><tr>"
        $N15Html += "<th>Test</th><th>Target</th><th>Result</th><th>IPs / Hostname</th><th>Status</th>"
        $N15Html += "</tr></thead><tbody>"
        foreach ($t in $tests) {
            $N15Html += "<tr><td>$(HtmlEncode $t.Test)</td><td><code>$(HtmlEncode $t.Target)</code></td>"
            $N15Html += "<td>$(HtmlEncode $t.Result)</td><td>$(HtmlEncode $t.IPs)</td>"
            $N15Html += "<td>$(StatusBadge $t.Status $t.Color)</td></tr>"
        }
        $N15Html += "</tbody></table></div>"
    } catch {
        $N15Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DNS-15 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DNS SECTION 16  -  AGING & SCAVENGING
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DNS-16: Aging & Scavenging..."
$N16Html = ''
if (-not $IsDNSInstalled -or -not $DnsModuleOk -or $AllZones.Count -eq 0) {
    $N16Html = "<p class='warn'>$(if (-not $IsDNSInstalled) { 'DNS not installed' } elseif (-not $DnsModuleOk) { 'DnsServer module not available' } else { 'No zones found' })  -  skipped.</p>"
} else {
    try {
        # Server-level scavenging
        $scavCfg = Get-DnsServerScavenging -ErrorAction Stop
        $scavEnabled = $scavCfg.ScavengingState
        $scavColor   = if ($scavEnabled) { 'green' } else { 'yellow' }

        $rows16 = @(
            @('Server Scavenging Enabled', (StatusBadge $(if ($scavEnabled) { 'Enabled' } else { 'Disabled' }) $scavColor)),
            @('Scavenging Interval',       (HtmlEncode $(if ($null -ne $scavCfg.ScavengingInterval) { $scavCfg.ScavengingInterval.ToString() } else { 'N/A' }))),
            @('Refresh Interval',          (HtmlEncode $(if ($null -ne $scavCfg.RefreshInterval) { $scavCfg.RefreshInterval.ToString() } else { 'N/A' }))),
            @('No-Refresh Interval',       (HtmlEncode $(if ($null -ne $scavCfg.NoRefreshInterval) { $scavCfg.NoRefreshInterval.ToString() } else { 'N/A' })))
        )
        $N16Html = BuildKVTable $rows16

        # Per-zone aging
        $primaryZones16 = @($AllZones | Where-Object { $_.ZoneType -eq 'Primary' -and $_.ZoneName -ne '.' })
        if ($primaryZones16.Count -gt 0) {
            $N16Html += "<h4 style='margin:12px 0 8px;'>Per-Zone Aging Configuration</h4>"
            $N16Html += "<div class='table-wrap'><table><thead><tr>"
            $N16Html += "<th>Zone</th><th>Aging Enabled</th><th>Refresh Interval</th><th>No-Refresh Interval</th>"
            $N16Html += "</tr></thead><tbody>"
            foreach ($z in ($primaryZones16 | Select-Object -First 20)) {
                try {
                    $za = Get-DnsServerZoneAging -ZoneName $z.ZoneName -ErrorAction Stop
                    $zaEnabled = $za.AgingEnabled
                    $zaColor   = if ($zaEnabled) { 'green' } else { 'grey' }
                    $N16Html += "<tr><td>$(HtmlEncode $z.ZoneName)</td>"
                    $N16Html += "<td>$(StatusBadge $(if ($zaEnabled) { 'Yes' } else { 'No' }) $zaColor)</td>"
                    $N16Html += "<td>$(HtmlEncode $(if ($null -ne $za.RefreshInterval) { $za.RefreshInterval.ToString() } else { 'N/A' }))</td>"
                    $N16Html += "<td>$(HtmlEncode $(if ($null -ne $za.NoRefreshInterval) { $za.NoRefreshInterval.ToString() } else { 'N/A' }))</td></tr>"
                } catch {
                    $N16Html += "<tr><td>$(HtmlEncode $z.ZoneName)</td><td colspan='3'><span class='warn'>$(HtmlEncode $_.Exception.Message)</span></td></tr>"
                }
            }
            $N16Html += "</tbody></table></div>"
        }
    } catch {
        $N16Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DNS-16 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DNS SECTION 17  -  CONDITIONAL FORWARDERS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DNS-17: Conditional Forwarders..."
$N17Html = ''
if (-not $IsDNSInstalled -or -not $DnsModuleOk) {
    $N17Html = "<p class='warn'>$(if (-not $IsDNSInstalled) { 'DNS not installed' } else { 'DnsServer module not available' })  -  skipped.</p>"
} else {
    try {
        $condFwds = @($AllZones | Where-Object { $_.ZoneType -eq 'Forwarder' })
        if ($condFwds.Count -eq 0) {
            $N17Html = "<p class='info'>No conditional forwarders configured.</p>"
        } else {
            $N17Html  = "<div class='table-wrap'><table><thead><tr>"
            $N17Html += "<th>Domain</th><th>Forwarder IPs</th><th>Reachability</th>"
            $N17Html += "</tr></thead><tbody>"
            foreach ($cf in $condFwds) {
                try {
                    $cfZone = Get-DnsServerZone -Name $cf.ZoneName -ErrorAction Stop
                    $cfIPs  = @()
                    try {
                        $cfFwd = Get-DnsServerZone -Name $cf.ZoneName -ErrorAction Stop
                        if ($null -ne $cfFwd.MasterServers) { $cfIPs = @($cfFwd.MasterServers | ForEach-Object { $_.ToString() }) }
                    } catch {}
                    if ($cfIPs.Count -eq 0) { $cfIPs = @('N/A') }
                    $reach = @()
                    foreach ($ip in ($cfIPs | Where-Object { $_ -ne 'N/A' })) {
                        $ok = Test-Tcp -Hostname $ip -Port 53 -TimeoutSec $ForwarderTestTimeoutSec
                        $reach += if ($ok) { "<span style='color:#3fb950;'>$ip &#x2714;</span>" }
                                  else { "<span style='color:#f85149;'>$ip &#x2716;</span>" }
                        if (-not $ok) { $CriticalFindings.Add("DNS-17 - Conditional forwarder for '$($cf.ZoneName)' IP $ip unreachable") }
                    }
                    $N17Html += "<tr><td>$(HtmlEncode $cf.ZoneName)</td>"
                    $N17Html += "<td>$(HtmlEncode ($cfIPs -join ', '))</td>"
                    $N17Html += "<td>$(if ($reach.Count -gt 0) { $reach -join ' ' } else { 'N/A' })</td></tr>"
                } catch {
                    $N17Html += "<tr><td>$(HtmlEncode $cf.ZoneName)</td><td colspan='2'><span class='warn'>$(HtmlEncode $_.Exception.Message)</span></td></tr>"
                }
            }
            $N17Html += "</tbody></table></div>"
        }
    } catch {
        $N17Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DNS-17 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DNS SECTION 18  -  DNSSEC STATUS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DNS-18: DNSSEC..."
$N18Html = ''
if (-not $IsDNSInstalled -or -not $DnsModuleOk -or $AllZones.Count -eq 0) {
    $N18Html = "<p class='warn'>$(if (-not $IsDNSInstalled) { 'DNS not installed' } elseif (-not $DnsModuleOk) { 'DnsServer module not available' } else { 'No zones found' })  -  skipped.</p>"
} else {
    try {
        $signedZones = @($AllZones | Where-Object { $null -ne $_.IsSigned -and $_.IsSigned -eq $true })
        if ($signedZones.Count -eq 0) {
            $N18Html = "<p class='info'>No DNSSEC-signed zones found. DNSSEC is not in use on this server.</p>"
        } else {
            $N18Html  = "<div class='table-wrap'><table><thead><tr>"
            $N18Html += "<th>Zone Name</th><th>Zone Type</th><th>DNSSEC Signed</th>"
            $N18Html += "</tr></thead><tbody>"
            foreach ($z in $signedZones) {
                $N18Html += "<tr><td>$(HtmlEncode $z.ZoneName)</td>"
                $N18Html += "<td>$(HtmlEncode $z.ZoneType.ToString())</td>"
                $N18Html += "<td>$(StatusBadge 'Signed' 'green')</td></tr>"
            }
            $N18Html += "</tbody></table></div>"
        }
    } catch {
        $N18Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DNS-18 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DNS SECTION 19  -  DNS CACHE
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DNS-19: DNS Cache..."
$N19Html = ''
if (-not $IsDNSInstalled -or -not $DnsModuleOk) {
    $N19Html = "<p class='warn'>$(if (-not $IsDNSInstalled) { 'DNS not installed' } else { 'DnsServer module not available' })  -  skipped.</p>"
} else {
    try {
        $cacheRoot  = Get-DnsServerCache -ErrorAction Stop
        $cacheEntries = 0
        try {
            $cacheZone    = Get-DnsServerZone -Name '..cache' -ErrorAction Stop
            $cacheRecords = @(Get-DnsServerResourceRecord -ZoneName '..cache' -ErrorAction Stop)
            $cacheEntries = $cacheRecords.Count
        } catch {}

        $maxTtl   = if ($null -ne $cacheRoot.MaxTTL) { $cacheRoot.MaxTTL.ToString() } else { 'N/A' }
        $maxNegTtl = if ($null -ne $cacheRoot.MaxNegativeTTL) { $cacheRoot.MaxNegativeTTL.ToString() } else { 'N/A' }
        $rows19 = @(
            @('Cache Entries (approx)', (HtmlEncode $cacheEntries.ToString())),
            @('Max TTL',               (HtmlEncode $maxTtl)),
            @('Max Negative TTL',      (HtmlEncode $maxNegTtl))
        )
        $N19Html = BuildKVTable $rows19
    } catch {
        $N19Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("DNS-19 error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DNS SECTION 20  -  EVENT LOG
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "DNS-20: Event Log..."
$N20Html = ''
try {
    if (-not $IsDNSInstalled) {
        $N20Html = "<p class='warn'>DNS not installed  -  skipped.</p>"
    } else {
        $dnsEvents = [System.Collections.Generic.List[object]]::new()
        $since48h  = (Get-Date).AddHours(-48)
        $dnsLogs   = @('DNS Server', 'Microsoft-Windows-DNSServer/Operational')
        foreach ($log in $dnsLogs) {
            try {
                $filter = @{ LogName = $log; Level = @(1,2); StartTime = $since48h }
                $evts   = @(Get-WinEvent -FilterHashtable $filter -MaxEvents 30 -ErrorAction Stop)
                foreach ($e in $evts) { $dnsEvents.Add($e) }
            } catch {}
        }
        # Also check System log for DNS source
        try {
            $filter2 = @{ LogName = 'System'; ProviderName = 'DNS'; Level = @(1,2); StartTime = $since48h }
            $evts2   = @(Get-WinEvent -FilterHashtable $filter2 -MaxEvents 20 -ErrorAction Stop)
            foreach ($e in $evts2) { $dnsEvents.Add($e) }
        } catch {}

        $topDnsEvents = $dnsEvents | Sort-Object TimeCreated -Descending | Select-Object -Unique -First 10

        if ($topDnsEvents.Count -eq 0) {
            $N20Html = "<p style='color:#3fb950;'>&#x2714; No Critical or Error events found in DNS logs (last 48 hours).</p>"
        } else {
            $N20Html  = "<div class='table-wrap'><table><thead><tr>"
            $N20Html += "<th>Time</th><th>Level</th><th>Event ID</th><th>Source</th><th>Message</th>"
            $N20Html += "</tr></thead><tbody>"
            foreach ($ev in $topDnsEvents) {
                $timeStr  = if ($null -ne $ev.TimeCreated) { $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
                $level    = if ($null -ne $ev.LevelDisplayName) { $ev.LevelDisplayName } else { 'Unknown' }
                $lvlColor = if ($level -eq 'Critical') { 'red' } else { 'yellow' }
                $msgRaw   = if ($null -ne $ev.Message) { $ev.Message } else { '' }
                $msgShort = if ($msgRaw.Length -gt 180) { $msgRaw.Substring(0,180) + '...' } else { $msgRaw }
                $N20Html += "<tr><td style='white-space:nowrap;'>$(HtmlEncode $timeStr)</td>"
                $N20Html += "<td>$(StatusBadge $level $lvlColor)</td>"
                $N20Html += "<td>$($ev.Id)</td>"
                $N20Html += "<td>$(HtmlEncode $(if ($null -ne $ev.ProviderName) { $ev.ProviderName } else { 'N/A' }))</td>"
                $N20Html += "<td>$(HtmlEncode $msgShort)</td></tr>"
            }
            $N20Html += "</tbody></table></div>"
        }
    }
} catch {
    $N20Html = "<p class='error'>Error: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("DNS-20 error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# KPI TILES & FINAL ASSEMBLY
# ═══════════════════════════════════════════════════════════════════════════════
$CritCount  = $CriticalFindings.Count
$ReportDate = (Get-Date).ToString('dddd, dd MMMM yyyy HH:mm:ss')
$EndTime    = Get-Date
$Duration   = ($EndTime - $StartTime).ToString('hh\:mm\:ss')

$OverallStatusBadge = if ($CritCount -gt 0) { StatusBadge 'CRITICAL' 'red' } else { StatusBadge 'HEALTHY' 'green' }
$dhcpBadge          = if ($IsDHCPInstalled) { StatusBadge 'Installed' 'green' } else { StatusBadge 'Not Installed' 'grey' }
$dnsBadge           = if ($IsDNSInstalled)  { StatusBadge 'Installed' 'green' } else { StatusBadge 'Not Installed' 'grey' }

$critScopesColor = if ($CriticalScopes -gt 0) { '#f85149' } else { '#3fb950' }
$failFwdColor    = if ($FailedForwarders -gt 0) { '#f85149' } else { '#3fb950' }
$critFindColor   = if ($CritCount -gt 0) { '#f85149' } else { '#3fb950' }

# Critical findings panel
$CritFindingsHtml = ''
if ($CritCount -gt 0) {
    $cfRows = ($CriticalFindings | ForEach-Object { "<li>$(HtmlEncode $_)</li>" }) -join ''
    $CritFindingsHtml = "<div style='background:rgba(219,97,162,.12);border:1px solid #db61a2;border-radius:8px;" +
        "padding:14px 18px;margin-bottom:16px;'>" +
        "<strong style='color:#db61a2;'>&#x26A0; $CritCount Critical Finding(s) Detected</strong>" +
        "<ul style='margin:.6rem 0 0 1.2rem;color:#db61a2;'>$cfRows</ul></div>"
}

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

$HtmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DHCP + DNS Health Check - $(HtmlEncode $ServerHostname)</title>
<style>
:root {
  --bg:      #0d1117;
  --card:    #161b22;
  --border:  #30363d;
  --text:    #c9d1d9;
  --muted:   #8b949e;
  --link:    #db61a2;
  --head-bg: #12091a;
  --th-bg:   #21262d;
  --tr-alt:  #1c2128;
  --accent:  #db61a2;
}
[data-theme="light"] {
  --bg:      #ffffff;
  --card:    #f6f8fa;
  --border:  #d0d7de;
  --text:    #24292f;
  --muted:   #57606a;
  --link:    #a020a0;
  --head-bg: #fdf0f9;
  --th-bg:   #fdf2fa;
  --tr-alt:  #fff8fd;
  --accent:  #a020a0;
}
* { box-sizing:border-box; margin:0; padding:0; }
body { font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
       background:var(--bg); color:var(--text); font-size:14px; line-height:1.5; }
a { color:var(--link); }
code { font-family:Consolas,'SFMono-Regular',monospace; font-size:.85em;
       background:var(--th-bg); padding:1px 5px; border-radius:4px; }
.page-wrap { max-width:1280px; margin:0 auto; padding:0 20px 40px; }
.header    { background:var(--head-bg); border-bottom:2px solid var(--accent);
             padding:14px 24px; display:flex; align-items:center;
             justify-content:space-between; flex-wrap:wrap; gap:12px; }
.header-left  { display:flex; align-items:center; gap:16px; }
.header-title h1 { font-size:1.25rem; font-weight:700; color:var(--accent);
                   display:flex; align-items:center; gap:8px; }
.header-title p  { font-size:.8rem; color:var(--muted); margin-top:2px; }
.logo-wrap { max-height:60px; }
.theme-toggle { cursor:pointer; background:var(--card); border:1px solid var(--accent);
                color:var(--text); border-radius:20px; padding:6px 14px;
                font-size:12px; display:flex; align-items:center; gap:6px; }
.theme-toggle:hover { background:var(--th-bg); }
.role-strip { display:flex; gap:12px; margin:14px 0 0;
              background:var(--card); border:1px solid var(--border);
              border-radius:8px; padding:12px 20px; flex-wrap:wrap; }
.role-item  { display:flex; align-items:center; gap:8px; font-size:.85rem; }
.summary-bar { display:flex; flex-wrap:wrap; gap:12px;
               background:var(--card); border:1px solid var(--border);
               border-top:2px solid var(--accent);
               border-radius:8px; padding:16px 20px; margin:14px 0; }
.stat-card   { flex:1 1 110px; text-align:center; }
.stat-count  { font-size:2rem; font-weight:700; line-height:1; }
.stat-label  { font-size:.72rem; color:var(--muted); margin-top:4px; }
.group-heading { font-size:1rem; font-weight:700; color:var(--accent);
                 margin:22px 0 10px; padding-left:4px;
                 border-left:3px solid var(--accent); }
.section-card { background:var(--card); border:1px solid var(--border);
                border-radius:8px; margin-bottom:10px; overflow:hidden; }
.section-summary { display:flex; align-items:center; gap:10px; cursor:pointer;
                   padding:11px 18px; list-style:none; user-select:none; }
.section-summary::-webkit-details-marker { display:none; }
.section-summary::marker { display:none; }
.sec-arrow { font-size:10px; color:var(--muted); display:inline-block;
             transition:transform .2s; flex-shrink:0; line-height:1; }
details[open] > .section-summary .sec-arrow { transform:rotate(90deg); }
.section-summary:hover { background:var(--th-bg); }
.sec-num { background:var(--accent); color:#fff; font-size:.7rem; font-weight:700;
           min-width:32px; height:22px; border-radius:11px; display:flex;
           align-items:center; justify-content:center; flex-shrink:0; padding:0 6px; }
.sec-title { font-weight:600; font-size:.9rem; flex:1; }
.section-body { padding:14px 18px; border-top:1px solid var(--border); }
h4 { font-size:.9rem; font-weight:600; }
.table-wrap { overflow-x:auto; border-radius:6px; border:1px solid var(--border); }
table { width:100%; border-collapse:collapse; font-size:13px; }
thead tr { background:var(--th-bg); }
th { padding:8px 12px; text-align:left; font-weight:600;
     border-bottom:1px solid var(--border); white-space:nowrap; }
td { padding:7px 12px; border-bottom:1px solid var(--border); vertical-align:top; }
tbody tr:nth-child(even) { background:var(--tr-alt); }
tbody tr:hover { background:var(--th-bg); }
.kv-table td { padding:7px 12px; border-bottom:1px solid var(--border); }
.td-label { font-weight:600; white-space:nowrap; width:220px; color:var(--muted); }
.disk-bar-outer { height:16px; background:var(--th-bg); border-radius:8px;
                  overflow:hidden; margin-bottom:2px; border:1px solid var(--border); }
.disk-bar-inner { height:100%; border-radius:8px; transition:width .4s ease; }
.badge { display:inline-block; font-size:.7rem; font-weight:600; padding:2px 8px;
         border-radius:20px; color:#fff; white-space:nowrap; }
.not-installed-card { display:flex; align-items:flex-start; gap:14px;
                      background:rgba(88,166,255,.08); border:1px solid #58a6ff;
                      border-radius:8px; padding:16px 20px; }
.ni-icon { font-size:1.8rem; flex-shrink:0; }
.error { color:#f85149; padding:8px 12px; background:rgba(248,81,73,.1);
         border-left:3px solid #f85149; border-radius:4px; }
.warn  { color:#d29922; padding:8px 12px; background:rgba(210,153,34,.1);
         border-left:3px solid #d29922; border-radius:4px; }
.info  { color:var(--muted); }
.footer { border-top:1px solid var(--border); padding:18px 0;
          margin-top:24px; color:var(--muted); font-size:.8rem;
          display:flex; justify-content:space-between; flex-wrap:wrap; gap:8px; }
@media (max-width:768px) {
  .summary-bar { gap:8px; }
  .stat-card { flex:1 1 90px; }
  .stat-count { font-size:1.4rem; }
  .header { flex-direction:column; align-items:flex-start; }
}
</style>
</head>
<body data-theme="dark">

<div class="header">
  <div class="header-left">
    $LogoHtml
    <div class="header-title">
      <h1>&#x1F5A7; DHCP + DNS Health Check</h1>
      <p>$(HtmlEncode $ServerHostname) &nbsp;|&nbsp; Generated: $(HtmlEncode $ReportDate)</p>
    </div>
  </div>
  <button class="theme-toggle" onclick="toggleTheme()" title="Toggle Dark/Light mode">
    <span id="theme-icon">&#x2600;&#xFE0F;</span> Toggle Theme
  </button>
</div>

<div class="page-wrap">

<div class="role-strip">
  <div class="role-item"><strong>DHCP Server:</strong> $dhcpBadge</div>
  <div class="role-item"><strong>DNS Server:</strong> $dnsBadge</div>
  <div class="role-item"><strong>DhcpServer Module:</strong> $(StatusBadge $(if ($DhcpModuleOk) { 'Available' } else { 'Not Available' }) $(if ($DhcpModuleOk) { 'green' } else { 'yellow' }))</div>
  <div class="role-item"><strong>DnsServer Module:</strong> $(StatusBadge $(if ($DnsModuleOk) { 'Available' } else { 'Not Available' }) $(if ($DnsModuleOk) { 'green' } else { 'yellow' }))</div>
</div>

<div class="summary-bar">
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.8rem; color:#3fb950;">$TotalScopes</div>
    <div class="stat-label">Total Scopes</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.8rem; color:$critScopesColor;">$CriticalScopes</div>
    <div class="stat-label">Critical Scopes (&gt;$ScopeUtilCritPct%)</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.8rem; color:#3fb950;">$TotalZones</div>
    <div class="stat-label">Total DNS Zones</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.8rem; color:$failFwdColor;">$FailedForwarders</div>
    <div class="stat-label">Failed Forwarders</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.8rem; color:$critFindColor;">$CritCount</div>
    <div class="stat-label">Critical Findings</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:.85rem; padding-top:10px;">$OverallStatusBadge</div>
    <div class="stat-label">Overall Status</div>
  </div>
</div>

$CritFindingsHtml

<div class="group-heading">&#x1F4E1; DHCP Server Checks</div>

$(BuildSection 1 'DHCP Service Status'   $D1Html ($D1Html -match "class='error'") $true  'D')
$(BuildSection 2 'Scope Inventory'        $D2Html ($D2Html -match "class='error'") $true  'D')
$(BuildSection 3 'Scope Utilization'      $D3Html ($D3Html -match "class='error'") $true  'D')
$(BuildSection 4 'DHCP Failover'          $D4Html ($D4Html -match "class='error'") $false 'D')
$(BuildSection 5 'Reservations'           $D5Html ($D5Html -match "class='error'") $false 'D')
$(BuildSection 6 'Lease Statistics'       $D6Html ($D6Html -match "class='error'") $false 'D')
$(BuildSection 7 'Audit Log'              $D7Html ($D7Html -match "class='error'") $false 'D')
$(BuildSection 8 'DHCP Database'          $D8Html ($D8Html -match "class='error'") $false 'D')
$(BuildSection 9 'DHCP Event Log'         $D9Html ($D9Html -match "class='error'") $false 'D')

<div class="group-heading">&#x1F310; DNS Server Checks</div>

$(BuildSection 10 'DNS Service Status'      $N10Html ($N10Html -match "class='error'") $true  'N')
$(BuildSection 11 'Zone Inventory'           $N11Html ($N11Html -match "class='error'") $true  'N')
$(BuildSection 12 'Zone Health'              $N12Html ($N12Html -match "class='error'") $true  'N')
$(BuildSection 13 'Forwarders'               $N13Html ($N13Html -match "class='error'") $true  'N')
$(BuildSection 14 'Root Hints'               $N14Html ($N14Html -match "class='error'") $false 'N')
$(BuildSection 15 'DNS Resolution Tests'     $N15Html ($N15Html -match "class='error'") $false 'N')
$(BuildSection 16 'Aging &amp; Scavenging'   $N16Html ($N16Html -match "class='error'") $false 'N')
$(BuildSection 17 'Conditional Forwarders'   $N17Html ($N17Html -match "class='error'") $false 'N')
$(BuildSection 18 'DNSSEC Status'            $N18Html ($N18Html -match "class='error'") $false 'N')
$(BuildSection 19 'DNS Cache'                $N19Html ($N19Html -match "class='error'") $false 'N')
$(BuildSection 20 'DNS Event Log'            $N20Html ($N20Html -match "class='error'") $false 'N')

<div class="footer">
  <div>
    <strong>DHCP + DNS Health Check v$ScriptVersion</strong> &nbsp;|&nbsp;
    Server: $(HtmlEncode $ServerHostname) &nbsp;|&nbsp;
    Generated: $(HtmlEncode $ReportDate) &nbsp;|&nbsp;
    Duration: $(HtmlEncode $Duration)
  </div>
  <div>Created by $AuthorLink</div>
</div>
</div>

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

# ── WRITE REPORT ─────────────────────────────────────────────────────────────
try {
    [System.IO.File]::WriteAllText($ReportFile, $HtmlReport, [System.Text.Encoding]::UTF8)
    Write-Host ""
    Write-Host "  [OK] Report saved: $ReportFile" -ForegroundColor Green
} catch {
    Write-Warning "Failed to write report: $_"
}

# ── STATUS FILE ───────────────────────────────────────────────────────────────
$isCritical   = $CriticalFindings.Count -gt 0
$statusSuffix = if ($isCritical) { '_CRITICAL' } else { '_HEALTHY' }
$StatusFile   = Join-Path $ReportsDir ($ReportStamp + $statusSuffix + '.txt')
$sep          = '=' * 70

if ($isCritical) {
    $sc  = "$sep`r`n DHCP + DNS HEALTH CHECK  -  *** CRITICAL ALERT ***`r`n$sep`r`n"
    $sc += " Status    : CRITICAL`r`n Server    : $ServerHostname`r`n"
    $sc += " Generated : $ReportDate`r`n Duration  : $Duration`r`n$sep`r`n"
    $sc += "`r`n CRITICAL FINDINGS ($($CriticalFindings.Count)):`r`n`r`n"
    $sc += (@($CriticalFindings) | ForEach-Object { "  [!] $_" }) -join "`r`n"
    $sc += "`r`n`r`n$sep`r`n Full HTML report : $ReportFile`r`n Status file      : $StatusFile`r`n$sep`r`n"
    $sc += " DHCP + DNS Health Check v$ScriptVersion  by $AuthorName`r`n$sep`r`n"
} else {
    $sc  = "$sep`r`n DHCP + DNS HEALTH CHECK  -  HEALTHY STATE`r`n$sep`r`n"
    $sc += " Status    : HEALTHY`r`n Server    : $ServerHostname`r`n"
    $sc += " Generated : $ReportDate`r`n Duration  : $Duration`r`n$sep`r`n"
    $sc += "`r`n No critical findings detected. DHCP and DNS are in a healthy state.`r`n"
    $sc += "`r`n$sep`r`n Full HTML report : $ReportFile`r`n Status file      : $StatusFile`r`n$sep`r`n"
    $sc += " DHCP + DNS Health Check v$ScriptVersion  by $AuthorName`r`n$sep`r`n"
}

if ($EnableStatusFile) {
    try {
        [System.IO.File]::WriteAllText($StatusFile, $sc, [System.Text.Encoding]::UTF8)
        $sColor = if ($isCritical) { 'Red' } else { 'Green' }
        Write-Host "  [OK] Status file: $StatusFile" -ForegroundColor $sColor
    } catch {
        Write-Warning "Failed to write status file: $_"
    }
}

Write-Host ""
Write-Host "===============================================================" -ForegroundColor DarkMagenta
Write-Host "  DHCP + DNS Health Check complete.  Duration: $Duration" -ForegroundColor DarkMagenta
Write-Host "  Report : $ReportFile" -ForegroundColor Yellow
if ($EnableStatusFile) {
    $slabel = if ($isCritical) { 'Status (CRITICAL)' } else { 'Status (HEALTHY)' }
    Write-Host "  $slabel : $StatusFile" -ForegroundColor $(if ($isCritical) { 'Red' } else { 'Green' })
    Write-Host ""
    Write-Host "  To send email alerts, run: .\DHCP_DNS_HealthCheck_EmailAlert.ps1" -ForegroundColor Cyan
}
Write-Host "===============================================================" -ForegroundColor DarkMagenta
Write-Host ""
