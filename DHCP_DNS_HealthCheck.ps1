#Requires -Version 5.1
<#
.SYNOPSIS
    Comprehensive DHCP and DNS Health Check Script

.DESCRIPTION
    Performs a thorough health check of the Windows DHCP Server and DNS Server
    services and exports the results to a single self-contained HTML file with a
    professional dark/light-themed dashboard UI.

.NOTES
    Version    : 1.0.0
    Author     : Tushar Gudde
    Requires   : PowerShell 5.1+, DhcpServer module (RSAT DHCP), DnsServer module (RSAT DNS)
    Permissions: Local Admin / DHCP Admin / DNS Admin
#>

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
# Company Branding
$CompanyLogoURL    = ''                          # URL/path to company logo (PNG/SVG). Leave blank to skip.
$CompanyWebsite    = ''                          # Company website URL for logo hyperlink.

# Author
$AuthorName        = 'Tushar Gudde'              # Author name shown in footer

# Event Log Settings
$EventLogHours     = 24                          # How many hours back to scan DHCP/DNS events

# DHCP scope utilisation thresholds (percent)
$DhcpWarnPct       = 80                          # Warn when scope utilisation exceeds this %
$DhcpCritPct       = 90                          # Critical when scope utilisation exceeds this %

# DNS resolution test host (used in Section 8)
$DnsTestExternalHost = 'www.microsoft.com'       # External hostname to test DNS resolution

# Email Alert Configuration
$EnableEmailAlert       = $false
$SMTPServer             = 'smtp.yourdomain.com'
$SMTPPort               = 587
$SMTPFrom               = 'dhcp-dns-healthcheck@yourdomain.com'
$SMTPTo                 = @('admin@yourdomain.com')
$SMTPSubject            = 'DHCP/DNS Health Check - CRITICAL ALERT'
$SMTPUseSSL             = $true
$SMTPCredentialUser     = ''                     # Leave blank for anonymous relay
$SMTPCredentialPass     = ''                     # Leave blank for anonymous relay
# SECURITY NOTE: For production use, avoid storing credentials in the script.
# Use Windows Credential Manager, a secrets vault, or prompt at runtime via Get-Credential.
# ──────────────────────────────────────────────────────────────────────────────

$ScriptVersion  = '1.0.0'
$StartTime      = Get-Date
$ScriptDir      = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ([string]::IsNullOrEmpty($ScriptDir)) { $ScriptDir = $PWD.Path }

# ── HELPER FUNCTIONS ──────────────────────────────────────────────────────────
function HtmlEncode {
    param([string]$text)
    if ([string]::IsNullOrEmpty($text)) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function TruncateMessage {
    param([string]$text, [int]$maxLen = 500)
    if ([string]::IsNullOrEmpty($text)) { return '' }
    $text = $text.Trim()
    if ($text.Length -gt $maxLen) { return ($text.Substring(0, $maxLen) + '...') }
    return $text
}

function StatusBadge {
    param([string]$text, [string]$color)
    # color: green | yellow | red | blue | grey
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

# Collect critical findings for optional email
$CriticalFindings = [System.Collections.Generic.List[string]]::new()

# ── REPORTS FOLDER ────────────────────────────────────────────────────────────
$ReportsDir = Join-Path $ScriptDir 'Reports'
if (-not (Test-Path $ReportsDir)) {
    try { New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null }
    catch { Write-Warning "Could not create Reports folder: $_" }
}
$ReportFile = Join-Path $ReportsDir ("DHCP_DNS_Health_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

# ── DETECT MODULES ────────────────────────────────────────────────────────────
$DhcpInstalled   = $false
$DhcpModuleOk    = $false
$DnsInstalled    = $false
$DnsModuleOk     = $false

try {
    $dhcpSvc = Get-Service -Name 'DHCPServer' -ErrorAction Stop
    $DhcpInstalled = ($dhcpSvc.Status -eq 'Running')
} catch { $DhcpInstalled = $false }

try {
    Import-Module DhcpServer -ErrorAction Stop
    $DhcpModuleOk = $true
} catch { $DhcpModuleOk = $false }

try {
    $dnsSvc = Get-Service -Name 'DNS' -ErrorAction Stop
    $DnsInstalled = ($dnsSvc.Status -eq 'Running')
} catch { $DnsInstalled = $false }

try {
    Import-Module DnsServer -ErrorAction Stop
    $DnsModuleOk = $true
} catch { $DnsModuleOk = $false }

Write-Host ""
Write-Host "+==============================================================+" -ForegroundColor DarkCyan
Write-Host "|       DHCP / DNS Health Check  v$ScriptVersion                       |" -ForegroundColor DarkCyan
Write-Host "+==============================================================+" -ForegroundColor DarkCyan
Write-Host ""

if (-not $DhcpInstalled) {
    Write-Warning "DHCP Server service not found or not running. DHCP sections will report N/A."
}
if (-not $DnsInstalled) {
    Write-Warning "DNS Server service not found or not running. DNS sections will report N/A."
}

# =============================================================================
# SECTION 1 - DHCP SERVER SERVICE STATUS
# =============================================================================
Write-Progress2 "Section 1: DHCP Server Service Status..."
$Sec1Html  = ''
$Sec1Error = $false
try {
    $svc = Get-Service -Name 'DHCPServer' -ErrorAction Stop
    $statusClr = switch ($svc.Status) {
        'Running' { 'green' }
        'Stopped' { 'red'   }
        default   { 'yellow' }
    }
    if ($svc.Status -ne 'Running') {
        $CriticalFindings.Add("Section 1 - DHCPServer service is $($svc.Status)")
        $Sec1Error = $true
    }
    $startClr = if ($svc.StartType -eq 'Automatic') { 'green' } else { 'yellow' }

    $rows = @(
        @('Service Name',  (HtmlEncode $svc.Name))
        @('Display Name',  (HtmlEncode $svc.DisplayName))
        @('Status',        (StatusBadge $svc.Status.ToString() $statusClr))
        @('Start Type',    (StatusBadge $svc.StartType.ToString() $startClr))
    )
    $rowsHtml = ($rows | ForEach-Object {
        "<tr><td class='td-label'>$($_[0])</td><td>$($_[1])</td></tr>"
    }) -join ''
    $Sec1Html = "<table class='kv-table'><tbody>$rowsHtml</tbody></table>"
} catch {
    $Sec1Html  = "<p class='error'>DHCP Server service not found - DHCP is not installed or has been removed: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec1Error = $true
    $CriticalFindings.Add("Section 1 - DHCPServer service not found: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 2 - DHCP SCOPES AND UTILISATION
# =============================================================================
Write-Progress2 "Section 2: DHCP Scopes and Utilisation..."
$Sec2Html  = ''
$Sec2Error = $false
$ScopeCount = 0
try {
    if (-not $DhcpInstalled) {
        $Sec2Html = "<p class='warn'>DHCP not installed - scope check skipped.</p>"
    } elseif (-not $DhcpModuleOk) {
        $Sec2Html = "<p class='warn'>DhcpServer PowerShell module not available. Install RSAT DHCP Tools.</p>"
    } else {
        $scopes = Get-DhcpServerv4Scope -ErrorAction Stop
        $ScopeCount = @($scopes).Count

        if ($ScopeCount -eq 0) {
            $Sec2Html = "<p class='info'>No IPv4 DHCP scopes configured on this server.</p>"
        } else {
            $rows = foreach ($scope in $scopes | Sort-Object ScopeId) {
                $stats = try {
                    Get-DhcpServerv4ScopeStatistics -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                } catch { $null }

                $total    = if ($stats) { $stats.AddressesFree + $stats.AddressesInUse } else { 0 }
                $inUse    = if ($stats) { $stats.AddressesInUse } else { 0 }
                $free     = if ($stats) { $stats.AddressesFree  } else { 0 }
                $pct      = if ($total -gt 0) { [int](($inUse / $total) * 100) } else { 0 }

                $pctClr = if ($pct -ge $DhcpCritPct) { 'red' } elseif ($pct -ge $DhcpWarnPct) { 'yellow' } else { 'green' }
                if ($pct -ge $DhcpCritPct) {
                    $CriticalFindings.Add("Section 2 - Scope $($scope.ScopeId) utilisation $pct% (critical threshold $DhcpCritPct%)")
                    $Sec2Error = $true
                } elseif ($pct -ge $DhcpWarnPct) {
                    $CriticalFindings.Add("Section 2 - Scope $($scope.ScopeId) utilisation $pct% (warning threshold $DhcpWarnPct%)")
                }

                $stateClr = if ($scope.State -eq 'Active') { 'green' } else { 'yellow' }
                $scopeId  = HtmlEncode $scope.ScopeId.IPAddressToString
                $name     = HtmlEncode $scope.Name
                $start    = HtmlEncode $scope.StartRange.IPAddressToString
                $end      = HtmlEncode $scope.EndRange.IPAddressToString
                $mask     = HtmlEncode $scope.SubnetMask.IPAddressToString

                "<tr><td>$scopeId</td><td>$name</td><td>$start - $end</td><td>$mask</td>" +
                "<td>$(StatusBadge $scope.State.ToString() $stateClr)</td>" +
                "<td>$inUse / $total</td><td>$(StatusBadge "$pct%" $pctClr)</td></tr>"
            }

            $Sec2Html = @"
<div class='table-wrap'>
<table>
  <thead><tr>
    <th>Scope ID</th><th>Name</th><th>Range</th><th>Mask</th>
    <th>State</th><th>In Use / Total</th><th>Utilisation</th>
  </tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
        }
    }
} catch {
    $Sec2Html  = "<p class='error'>Error retrieving DHCP scopes: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec2Error = $true
    $CriticalFindings.Add("Section 2 - DHCP scopes error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 3 - DHCP FAILOVER RELATIONSHIPS
# =============================================================================
Write-Progress2 "Section 3: DHCP Failover Relationships..."
$Sec3Html  = ''
$Sec3Error = $false
try {
    if (-not $DhcpInstalled) {
        $Sec3Html = "<p class='warn'>DHCP not installed - failover check skipped.</p>"
    } elseif (-not $DhcpModuleOk) {
        $Sec3Html = "<p class='warn'>DhcpServer PowerShell module not available.</p>"
    } else {
        $failovers = try { Get-DhcpServerv4Failover -ErrorAction SilentlyContinue } catch { @() }

        if (-not $failovers -or @($failovers).Count -eq 0) {
            $Sec3Html = "<p class='info'>No DHCP failover relationships configured.</p>"
        } else {
            $rows = foreach ($fo in $failovers) {
                $stateClr = switch ($fo.State) {
                    'Normal'   { 'green'  }
                    'Recovery' { 'yellow' }
                    default    { 'red'    }
                }
                if ($fo.State -ne 'Normal') {
                    $CriticalFindings.Add("Section 3 - DHCP failover '$($fo.Name)' state is $($fo.State)")
                    $Sec3Error = $true
                }
                $foName    = HtmlEncode $fo.Name
                $partner   = HtmlEncode $fo.PartnerServer
                $mode      = HtmlEncode $fo.Mode.ToString()
                $scopeIds  = HtmlEncode ($fo.ScopeId -join ', ')
                "<tr><td>$foName</td><td>$partner</td><td>$mode</td>" +
                "<td>$(StatusBadge $fo.State.ToString() $stateClr)</td><td>$scopeIds</td></tr>"
            }

            $Sec3Html = @"
<div class='table-wrap'>
<table>
  <thead><tr>
    <th>Failover Name</th><th>Partner Server</th><th>Mode</th><th>State</th><th>Scopes</th>
  </tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
        }
    }
} catch {
    $Sec3Html  = "<p class='error'>Error retrieving DHCP failover: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec3Error = $true
    $CriticalFindings.Add("Section 3 - DHCP failover error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 4 - DHCP LEASE SUMMARY
# =============================================================================
Write-Progress2 "Section 4: DHCP Lease Summary..."
$Sec4Html  = ''
$Sec4Error = $false
try {
    if (-not $DhcpInstalled) {
        $Sec4Html = "<p class='warn'>DHCP not installed - lease summary skipped.</p>"
    } elseif (-not $DhcpModuleOk) {
        $Sec4Html = "<p class='warn'>DhcpServer PowerShell module not available.</p>"
    } else {
        $scopes = try { Get-DhcpServerv4Scope -ErrorAction SilentlyContinue } catch { @() }
        if (-not $scopes -or @($scopes).Count -eq 0) {
            $Sec4Html = "<p class='info'>No DHCP scopes found. No leases to report.</p>"
        } else {
            $rows = foreach ($scope in $scopes | Sort-Object ScopeId) {
                $leases = try {
                    Get-DhcpServerv4Lease -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue |
                    Where-Object { $_.AddressState -eq 'Active' } |
                    Select-Object -First 5
                } catch { @() }

                foreach ($lease in $leases) {
                    $clientName = HtmlEncode (TruncateMessage ($lease.HostName -as [string]) 40)
                    $ip         = HtmlEncode $lease.IPAddress.IPAddressToString
                    $mac        = HtmlEncode $lease.ClientId
                    $expires    = HtmlEncode $lease.LeaseExpiryTime.ToString('yyyy-MM-dd HH:mm')
                    $scopeIdStr = HtmlEncode $scope.ScopeId.IPAddressToString
                    "<tr><td>$scopeIdStr</td><td>$ip</td><td>$clientName</td><td>$mac</td><td>$expires</td></tr>"
                }
            }

            if ($rows) {
                $Sec4Html = @"
<p class='info'>Showing up to 5 active leases per scope.</p>
<div class='table-wrap'>
<table>
  <thead><tr>
    <th>Scope</th><th>IP Address</th><th>Host Name</th><th>MAC</th><th>Lease Expires</th>
  </tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
            } else {
                $Sec4Html = "<p class='info'>No active leases found across all scopes.</p>"
            }
        }
    }
} catch {
    $Sec4Html  = "<p class='error'>Error retrieving DHCP leases: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec4Error = $true
    $CriticalFindings.Add("Section 4 - DHCP lease error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 5 - DNS SERVER SERVICE STATUS
# =============================================================================
Write-Progress2 "Section 5: DNS Server Service Status..."
$Sec5Html  = ''
$Sec5Error = $false
try {
    $svc = Get-Service -Name 'DNS' -ErrorAction Stop
    $statusClr = switch ($svc.Status) {
        'Running' { 'green'  }
        'Stopped' { 'red'    }
        default   { 'yellow' }
    }
    if ($svc.Status -ne 'Running') {
        $CriticalFindings.Add("Section 5 - DNS service is $($svc.Status)")
        $Sec5Error = $true
    }
    $startClr = if ($svc.StartType -eq 'Automatic') { 'green' } else { 'yellow' }

    # DNS Server version via registry
    $dnsVersion = ''
    try {
        $dnsRegPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\DNS\Parameters'
        if (Test-Path $dnsRegPath) {
            $dnsVer = (Get-ItemProperty -Path $dnsRegPath -ErrorAction SilentlyContinue).Version
            if ($dnsVer) { $dnsVersion = $dnsVer.ToString() }
        }
    } catch {}

    $rows = @(
        @('Service Name',  (HtmlEncode $svc.Name))
        @('Display Name',  (HtmlEncode $svc.DisplayName))
        @('Status',        (StatusBadge $svc.Status.ToString() $statusClr))
        @('Start Type',    (StatusBadge $svc.StartType.ToString() $startClr))
        @('Version',       (HtmlEncode $dnsVersion))
    )
    $rowsHtml = ($rows | ForEach-Object {
        "<tr><td class='td-label'>$($_[0])</td><td>$($_[1])</td></tr>"
    }) -join ''
    $Sec5Html = "<table class='kv-table'><tbody>$rowsHtml</tbody></table>"
} catch {
    $Sec5Html  = "<p class='error'>DNS Server service not found - DNS is not installed or has been removed: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec5Error = $true
    $CriticalFindings.Add("Section 5 - DNS service not found: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 6 - DNS ZONES
# =============================================================================
Write-Progress2 "Section 6: DNS Zones..."
$Sec6Html  = ''
$Sec6Error = $false
$ZoneCount = 0
try {
    if (-not $DnsInstalled) {
        $Sec6Html = "<p class='warn'>DNS not installed - zone check skipped.</p>"
    } elseif (-not $DnsModuleOk) {
        $Sec6Html = "<p class='warn'>DnsServer PowerShell module not available. Install RSAT DNS Tools.</p>"
    } else {
        $zones = Get-DnsServerZone -ErrorAction Stop
        $ZoneCount = @($zones).Count

        if ($ZoneCount -eq 0) {
            $Sec6Html = "<p class='info'>No DNS zones found on this server.</p>"
        } else {
            $rows = foreach ($zone in $zones | Sort-Object ZoneName) {
                $zType    = HtmlEncode $zone.ZoneType.ToString()
                $zName    = HtmlEncode $zone.ZoneName
                $dynamic  = HtmlEncode $zone.DynamicUpdate.ToString()
                $adInt    = if ($zone.IsAutoCreated) { StatusBadge 'Auto' 'grey' } else { StatusBadge 'Manual' 'blue' }
                $signed   = if ($zone.IsDsIntegrated) { StatusBadge 'AD-Integrated' 'green' } else { StatusBadge 'File-based' 'yellow' }
                $paused   = if ($zone.ZonePaused) {
                    $CriticalFindings.Add("Section 6 - DNS zone '$($zone.ZoneName)' is paused")
                    $Sec6Error = $true
                    StatusBadge 'Paused' 'red'
                } else { StatusBadge 'Active' 'green' }
                "<tr><td>$zName</td><td>$zType</td><td>$signed</td><td>$dynamic</td><td>$adInt</td><td>$paused</td></tr>"
            }

            $Sec6Html = @"
<div class='table-wrap'>
<table>
  <thead><tr>
    <th>Zone Name</th><th>Type</th><th>Integration</th>
    <th>Dynamic Update</th><th>Origin</th><th>Status</th>
  </tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
        }
    }
} catch {
    $Sec6Html  = "<p class='error'>Error retrieving DNS zones: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec6Error = $true
    $CriticalFindings.Add("Section 6 - DNS zones error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 7 - DNS FORWARDERS AND ROOT HINTS
# =============================================================================
Write-Progress2 "Section 7: DNS Forwarders and Root Hints..."
$Sec7Html  = ''
$Sec7Error = $false
try {
    if (-not $DnsInstalled) {
        $Sec7Html = "<p class='warn'>DNS not installed - forwarder check skipped.</p>"
    } elseif (-not $DnsModuleOk) {
        $Sec7Html = "<p class='warn'>DnsServer PowerShell module not available.</p>"
    } else {
        # Forwarders
        $fwdRows = ''
        try {
            $forwarders = Get-DnsServerForwarder -ErrorAction SilentlyContinue
            if ($forwarders -and $forwarders.IPAddress) {
                foreach ($ip in $forwarders.IPAddress) {
                    $ipStr = HtmlEncode $ip.IPAddressToString
                    $fwdRows += "<tr><td>$ipStr</td><td>$(StatusBadge 'Configured' 'blue')</td></tr>"
                }
            } else {
                $fwdRows = "<tr><td colspan='2'><em>No forwarders configured</em></td></tr>"
            }
        } catch {
            $fwdRows = "<tr><td colspan='2' class='error'>Error: $(HtmlEncode $_.Exception.Message)</td></tr>"
        }

        # Root hints count
        $rootHintCount = 0
        try {
            $rootHints = Get-DnsServerRootHint -ErrorAction SilentlyContinue
            $rootHintCount = if ($rootHints) { @($rootHints).Count } else { 0 }
        } catch {}

        $Sec7Html = @"
<h3>Forwarders</h3>
<div class='table-wrap'>
<table>
  <thead><tr><th>Forwarder IP</th><th>Status</th></tr></thead>
  <tbody>$fwdRows</tbody>
</table>
</div>
<h3>Root Hints &nbsp;$(StatusBadge "$rootHintCount root servers" 'blue')</h3>
<p class='info'>Root hints define fallback resolution when forwarders are unavailable.</p>
"@
    }
} catch {
    $Sec7Html  = "<p class='error'>Error retrieving DNS forwarders: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec7Error = $true
    $CriticalFindings.Add("Section 7 - DNS forwarders error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 8 - DNS SERVER DIAGNOSTICS
# =============================================================================
Write-Progress2 "Section 8: DNS Server Diagnostics..."
$Sec8Html  = ''
$Sec8Error = $false
try {
    if (-not $DnsInstalled) {
        $Sec8Html = "<p class='warn'>DNS not installed - diagnostics skipped.</p>"
    } elseif (-not $DnsModuleOk) {
        $Sec8Html = "<p class='warn'>DnsServer PowerShell module not available.</p>"
    } else {
        # DNS Server settings
        $serverSettings = try { Get-DnsServer -ErrorAction SilentlyContinue } catch { $null }

        $recEnabled    = 'N/A'
        $cacheSettings = 'N/A'
        $serverName    = $env:COMPUTERNAME

        if ($serverSettings) {
            $recEnabled = if ($serverSettings.ServerRecursion -and $serverSettings.ServerRecursion.Enable) {
                StatusBadge 'Enabled' 'green'
            } else {
                StatusBadge 'Disabled' 'yellow'
            }
        }

        # Test basic DNS resolution
        $testResults = @()
        foreach ($testHost in @($DnsTestExternalHost, $env:USERDNSDOMAIN)) {
            if ([string]::IsNullOrWhiteSpace($testHost)) { continue }
            try {
                $res = Resolve-DnsName -Name $testHost -Server 127.0.0.1 -ErrorAction Stop | Select-Object -First 1
                $testResults += @{Name=$testHost; Status='Resolved'; Detail=$res.IPAddress; Color='green'}
            } catch {
                $testResults += @{Name=$testHost; Status='Failed'; Detail=$_.Exception.Message; Color='red'}
                $CriticalFindings.Add("Section 8 - DNS resolution failed for '$testHost': $($_.Exception.Message)")
                $Sec8Error = $true
            }
        }

        $dnsTestRows = ($testResults | ForEach-Object {
            $n = HtmlEncode $_['Name']
            $d = HtmlEncode (TruncateMessage $_['Detail'] 60)
            "<tr><td>$n</td><td>$(StatusBadge $_['Status'] $_['Color'])</td><td>$d</td></tr>"
        }) -join ''

        $rows = @(
            @('Server Name',    (HtmlEncode $serverName))
            @('Recursion',      $recEnabled)
        )
        $rowsHtml = ($rows | ForEach-Object {
            "<tr><td class='td-label'>$($_[0])</td><td>$($_[1])</td></tr>"
        }) -join ''

        $Sec8Html = @"
<table class='kv-table'>
  <tbody>$rowsHtml</tbody>
</table>
<h3>DNS Resolution Tests</h3>
<div class='table-wrap'>
<table>
  <thead><tr><th>Query</th><th>Result</th><th>Detail</th></tr></thead>
  <tbody>$dnsTestRows</tbody>
</table>
</div>
"@
    }
} catch {
    $Sec8Html  = "<p class='error'>Error running DNS diagnostics: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec8Error = $true
    $CriticalFindings.Add("Section 8 - DNS diagnostics error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 9 - DHCP AND DNS EVENT LOG
# =============================================================================
Write-Progress2 "Section 9: DHCP and DNS Event Log..."
$Sec9Html  = ''
$Sec9Error = $false

$dhcpDnsEventSources = @('Microsoft-Windows-DHCP-Server', 'Microsoft-Windows-DNSServer', 'DNS', 'DhcpServer')

# WinEvent level constants
$EvtLevelCritical = 1
$EvtLevelError    = 2
$EvtLevelWarning  = 3

try {
    $since = (Get-Date).AddHours(-$EventLogHours)
    $events = @()

    foreach ($src in $dhcpDnsEventSources) {
        try {
            $evts = Get-WinEvent -ProviderName $src -ErrorAction SilentlyContinue |
                    Where-Object { $_.TimeCreated -ge $since -and $_.Level -in $EvtLevelCritical,$EvtLevelError,$EvtLevelWarning } |
                    Select-Object -First 50
            if ($evts) { $events += $evts }
        } catch {}
    }

    # Also check System log for DHCP/DNS service events
    try {
        $sysEvts = Get-WinEvent -LogName 'System' -ErrorAction SilentlyContinue |
                   Where-Object {
                       $_.TimeCreated -ge $since -and
                       $_.Level -in $EvtLevelCritical,$EvtLevelError,$EvtLevelWarning -and
                       ($_.ProviderName -match 'DHCP|DNS')
                   } | Select-Object -First 50
        if ($sysEvts) { $events += $sysEvts }
    } catch {}

    $events = $events | Sort-Object TimeCreated -Descending | Select-Object -First 100

    if ($events.Count -eq 0) {
        $Sec9Html = "<p class='info'>No DHCP or DNS warning/error events found in the last $EventLogHours hour(s).</p>"
    } else {
        $rows = foreach ($ev in $events) {
            $lvl = switch ($ev.Level) {
                $EvtLevelCritical { @('Critical', 'red') }
                $EvtLevelError    { @('Error', 'red') }
                $EvtLevelWarning  { @('Warning', 'yellow') }
                default           { @('Info', 'blue') }
            }
            if ($ev.Level -in $EvtLevelCritical,$EvtLevelError) {
                $msgSnippet = if ($ev.Message) { $ev.Message.Substring(0, [Math]::Min(100, $ev.Message.Length)) } else { '(no message)' }
                $CriticalFindings.Add("Section 9 - $($ev.ProviderName) event $($ev.Id): $msgSnippet")
                $Sec9Error = $true
            }
            $rowClass = switch ($ev.Level) {
                $EvtLevelCritical { 'row-critical' }
                $EvtLevelError    { 'row-error'    }
                $EvtLevelWarning  { 'row-warning'  }
                default           { '' }
            }
            $timeStr  = HtmlEncode $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            $provider = HtmlEncode $ev.ProviderName
            $evId     = HtmlEncode $ev.Id.ToString()
            $msgShort = HtmlEncode (TruncateMessage $ev.Message 120)
            "<tr class='$rowClass'><td>$timeStr</td><td>$provider</td><td>$evId</td><td>$(StatusBadge $lvl[0] $lvl[1])</td><td>$msgShort</td></tr>"
        }

        $Sec9Html = @"
<div class='table-wrap'>
<table>
  <thead><tr>
    <th>Time</th><th>Source</th><th>Event ID</th><th>Level</th><th>Message</th>
  </tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
    }
} catch {
    $Sec9Html  = "<p class='error'>Error reading DHCP/DNS event log: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec9Error = $true
    $CriticalFindings.Add("Section 9 - Event log error: $($_.Exception.Message)")
}

# =============================================================================
# SUMMARY COUNTS
# =============================================================================
$CritCount = $CriticalFindings.Count

Write-Progress2 "Building HTML report..."

# =============================================================================
# COMPANY LOGO HTML
# =============================================================================
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

# =============================================================================
# HELPER: Build collapsible section
# =============================================================================
function BuildSection {
    param(
        [int]$num,
        [string]$title,
        [string]$body,
        [bool]$hasError = $false,
        [bool]$open = $false
    )
    $openAttr  = if ($open) { ' open' } else { '' }
    $indicator = if ($hasError) { "<span style='color:#da3633;'>&#x2716;</span>" } else { "<span style='color:#238636;'>&#x2714;</span>" }
    return @"
<details class='section-card'$openAttr>
  <summary class='section-summary'>
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

$EndTime    = Get-Date
$Duration   = ($EndTime - $StartTime).ToString('hh\:mm\:ss')
$ReportDate = $EndTime.ToString('dddd, dd MMMM yyyy HH:mm:ss')

# =============================================================================
# BUILD FULL HTML REPORT
# =============================================================================
$HtmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DHCP/DNS Health Check Report - $($env:COMPUTERNAME)</title>
<style>
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
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: var(--bg); color: var(--text);
  font-size: 14px; line-height: 1.6;
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
em { font-style: italic; }
.page-wrap  { max-width: 1400px; margin: 0 auto; padding: 0 16px 40px; }
.header     { background: var(--head-bg); border-bottom: 1px solid var(--border);
              padding: 16px 24px; display: flex; align-items: center;
              justify-content: space-between; flex-wrap: wrap; gap: 12px; }
.header-left  { display: flex; align-items: center; gap: 16px; }
.logo-wrap img { max-height: 52px; }
.header-title h1 { font-size: 1.4rem; color: var(--text); }
.header-title p  { font-size: .8rem; color: var(--muted); margin: 0; }
.theme-toggle { cursor: pointer; background: var(--card); border: 1px solid var(--border);
                color: var(--text); border-radius: 20px; padding: 6px 14px;
                font-size: 12px; display: flex; align-items: center; gap: 6px; }
.theme-toggle:hover { background: var(--th-bg); }
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
.section-card { background: var(--card); border: 1px solid var(--border);
                border-radius: 8px; margin-bottom: 16px; overflow: hidden; }
.section-summary { display: flex; align-items: center; gap: 10px; cursor: pointer;
                   padding: 14px 18px; list-style: none; user-select: none; }
.section-summary::-webkit-details-marker { display: none; }
.section-summary::before { content: '\25B6'; font-size: 10px; color: var(--muted);
                            transition: transform .2s; }
details[open] > .section-summary::before { transform: rotate(90deg); }
.section-summary:hover { background: var(--th-bg); }
.sec-num   { background: var(--blue); color: #fff; font-size: .7rem; font-weight: 700;
             width: 22px; height: 22px; border-radius: 50%; display: flex;
             align-items: center; justify-content: center; flex-shrink: 0; }
.sec-title { font-weight: 600; font-size: .95rem; flex: 1; }
.section-body { padding: 16px 18px; border-top: 1px solid var(--border); }
.table-wrap { overflow-x: auto; border-radius: 6px; border: 1px solid var(--border); }
table       { width: 100%; border-collapse: collapse; font-size: 13px; }
thead tr    { background: var(--th-bg); position: sticky; top: 0; z-index: 1; }
th          { padding: 8px 12px; text-align: left; font-weight: 600;
              border-bottom: 1px solid var(--border); white-space: nowrap; }
td          { padding: 7px 12px; border-bottom: 1px solid var(--border); vertical-align: top; }
tbody tr:nth-child(even) { background: var(--tr-alt); }
tbody tr:hover { background: var(--th-bg); }
.kv-table    { width: 100%; border-collapse: collapse; font-size: 13px; }
.kv-table td { padding: 7px 12px; border-bottom: 1px solid var(--border); }
.td-label    { font-weight: 600; white-space: nowrap; width: 220px; color: var(--muted); }
.row-critical { background: rgba(218,54,51,.15) !important; }
.row-error    { background: rgba(218,54,51,.08) !important; }
.row-warning  { background: rgba(210,153,34,.12) !important; }
.badge { display: inline-block; font-size: .7rem; font-weight: 600; padding: 2px 8px;
         border-radius: 20px; color: #fff; white-space: nowrap; }
.code-block { background: var(--pre-bg); border: 1px solid var(--border); border-radius: 6px;
              padding: 12px; overflow-x: auto; font-size: 12px;
              font-family: 'SFMono-Regular', Consolas, monospace;
              white-space: pre; color: var(--text); line-height: 1.5; }
.error { color: #f85149; padding: 8px 12px; background: rgba(248,81,73,.1);
         border-left: 3px solid #f85149; border-radius: 4px; }
.warn  { color: #d29922; }
.info  { color: var(--muted); }
.footer { border-top: 1px solid var(--border); padding: 20px 0;
          margin-top: 24px; color: var(--muted); font-size: .8rem;
          display: flex; justify-content: space-between; flex-wrap: wrap; gap: 8px; }
@media (max-width: 768px) {
  .summary-bar { gap: 8px; }
  .stat-card   { flex: 1 1 100px; }
  .stat-count  { font-size: 1.5rem; }
  .header      { flex-direction: column; align-items: flex-start; }
}
</style>
</head>
<body data-theme="dark">

<div class="header">
  <div class="header-left">
    $LogoHtml
    <div class="header-title">
      <h1>&#x1F4C1; DHCP / DNS Health Check</h1>
      <p>$($env:COMPUTERNAME) &nbsp;|&nbsp; Generated: $(HtmlEncode $ReportDate)</p>
    </div>
  </div>
  <button class="theme-toggle" onclick="toggleTheme()" title="Toggle Dark/Light mode">
    <span id="theme-icon">&#x2600;&#xFE0F;</span> Toggle Theme
  </button>
</div>

<div class="page-wrap">
<div class="summary-bar">
  <div class="stat-card">
    <div class="stat-count $(if ($DhcpInstalled) { 'c-green' } else { 'c-red' })">$(if ($DhcpInstalled) { 'Running' } else { 'Stopped' })</div>
    <div class="stat-label">DHCP Service</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-blue">$ScopeCount</div>
    <div class="stat-label">DHCP Scopes</div>
  </div>
  <div class="stat-card">
    <div class="stat-count $(if ($DnsInstalled) { 'c-green' } else { 'c-red' })">$(if ($DnsInstalled) { 'Running' } else { 'Stopped' })</div>
    <div class="stat-label">DNS Service</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-blue">$ZoneCount</div>
    <div class="stat-label">DNS Zones</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-red">$CritCount</div>
    <div class="stat-label">Critical Findings</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-muted" style="font-size:1rem; padding-top:8px;">$($env:COMPUTERNAME)</div>
    <div class="stat-label">Server</div>
  </div>
</div>

$(BuildSection 1 'DHCP Server Service Status'       $Sec1Html $Sec1Error $true)
$(BuildSection 2 'DHCP Scopes and Utilisation'      $Sec2Html $Sec2Error $true)
$(BuildSection 3 'DHCP Failover Relationships'      $Sec3Html $Sec3Error $false)
$(BuildSection 4 'DHCP Lease Summary'               $Sec4Html $Sec4Error $false)
$(BuildSection 5 'DNS Server Service Status'        $Sec5Html $Sec5Error $true)
$(BuildSection 6 'DNS Zones'                        $Sec6Html $Sec6Error $true)
$(BuildSection 7 'DNS Forwarders and Root Hints'    $Sec7Html $Sec7Error $false)
$(BuildSection 8 'DNS Server Diagnostics'           $Sec8Html $Sec8Error $false)
$(BuildSection 9 'DHCP and DNS Event Log'           $Sec9Html $Sec9Error $false)

<div class="footer">
  <div>
    <strong>DHCP/DNS Health Check v$ScriptVersion</strong> &nbsp;|&nbsp;
    Generated: $(HtmlEncode $ReportDate) &nbsp;|&nbsp;
    Duration: $(HtmlEncode $Duration)
  </div>
  <div>Created by $(HtmlEncode $AuthorName)</div>
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

# =============================================================================
# WRITE REPORT FILE
# =============================================================================
try {
    [System.IO.File]::WriteAllText($ReportFile, $HtmlReport, [System.Text.Encoding]::UTF8)
    Write-Host ""
    Write-Host "  [OK] Report saved to: $ReportFile" -ForegroundColor Green
} catch {
    Write-Warning "Failed to write report: $_"
}

# =============================================================================
# EMAIL ALERT (OPTIONAL)
# =============================================================================
if ($EnableEmailAlert -and $CriticalFindings.Count -gt 0) {
    Write-Progress2 "Sending critical alert email..."
    try {
        $emailBody  = "DHCP/DNS Health Check detected $($CriticalFindings.Count) critical finding(s):`r`n`r`n"
        $emailBody += ($CriticalFindings | ForEach-Object { "* $_" }) -join "`r`n"
        $emailBody += "`r`n`r`nPlease review the full report: $ReportFile"
        $emailBody += "`r`n`r`n-- DHCP/DNS Health Check v$ScriptVersion by $AuthorName"

        $mailParams = @{
            SmtpServer  = $SMTPServer
            Port        = $SMTPPort
            From        = $SMTPFrom
            To          = $SMTPTo
            Subject     = $SMTPSubject
            Body        = $emailBody
            UseSsl      = $SMTPUseSSL
            ErrorAction = 'Stop'
        }

        if (-not [string]::IsNullOrWhiteSpace($SMTPCredentialUser)) {
            $secPass   = ConvertTo-SecureString $SMTPCredentialPass -AsPlainText -Force
            $cred      = New-Object System.Management.Automation.PSCredential($SMTPCredentialUser, $secPass)
            $mailParams['Credential'] = $cred
        }

        Send-MailMessage @mailParams
        Write-Host "  [OK] Alert email sent to: $($SMTPTo -join ', ')" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to send alert email: $_"
    }
} elseif ($EnableEmailAlert -and $CriticalFindings.Count -eq 0) {
    Write-Progress2 "No critical findings - email alert skipped."
}

Write-Host ""
Write-Host "===============================================================" -ForegroundColor DarkCyan
Write-Host "  DHCP/DNS Health Check complete.  Duration: $Duration" -ForegroundColor DarkCyan
Write-Host "  Report: $ReportFile"                                   -ForegroundColor Yellow
Write-Host "===============================================================" -ForegroundColor DarkCyan
Write-Host ""
