#Requires -Version 5.1
<#
.SYNOPSIS
    Comprehensive Active Directory Health Check Script

.DESCRIPTION
    Performs a thorough health check of the Active Directory environment and
    exports the results to a single self-contained HTML file with a professional
    dark/light-themed dashboard UI.

.NOTES
    Version    : 2.1.0
    Author     : Tushar Gudde
    Requires   : PowerShell 5.1+, RSAT AD DS Tools (ActiveDirectory module)
    Permissions: Domain Admin or equivalent
#>

# -- CONFIGURATION -------------------------------------------------------------
# Company Branding
$CompanyLogoURL    = ''                          # URL/path to company logo (PNG/SVG). Leave blank to skip.
$CompanyWebsite    = 'https://tushargudde.tech'                          # Company website URL for logo hyperlink.

# Author
$AuthorName        = 'Tushar Gudde'              # Author name shown in footer

# Event Log Settings
$EventLogHours     = 2                           # How many hours back to scan events

# Stale Object Threshold
$StaleThresholdDays = 90                         # Days of inactivity before flagged stale

# Output Settings
# -- Status Summary File (saved next to the HTML report) --
$EnableStatusFile       = $true                  # Write a plain-text status summary file alongside the HTML report
# ------------------------------------------------------------------------------
# NOTE: For email notifications, run AD_HealthCheck_EmailAlert.ps1 after this script.

$ScriptVersion  = '2.1.0'
$StartTime      = Get-Date
$ScriptDir      = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ([string]::IsNullOrEmpty($ScriptDir)) { $ScriptDir = $PWD.Path }

# -- HELPER FUNCTIONS ----------------------------------------------------------
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

function ConvertADMode {
    # Converts raw ADForestMode/ADDomainMode enum strings to human-readable names.
    param([string]$raw)
    if ([string]::IsNullOrEmpty($raw)) { return 'Unknown' }
    switch -Regex ($raw) {
        'Windows2000'        { return 'Windows 2000' }
        'Windows2003Interim' { return 'Windows Server 2003 Interim' }
        'Windows2003'        { return 'Windows Server 2003' }
        'Windows2008R2'      { return 'Windows Server 2008 R2' }
        'Windows2008'        { return 'Windows Server 2008' }
        'Windows2012R2'      { return 'Windows Server 2012 R2' }
        'Windows2012'        { return 'Windows Server 2012' }
        'Windows2016'        { return 'Windows Server 2016' }
        'Windows2019'        { return 'Windows Server 2019' }
        'Windows2025'        { return 'Windows Server 2025' }
        default              { return $raw }
    }
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

# Collect critical findings throughout all sections
$CriticalFindings = [System.Collections.Generic.List[string]]::new()

# -- REPORTS FOLDER ------------------------------------------------------------
$ReportsDir = Join-Path $ScriptDir 'Reports'
if (-not (Test-Path $ReportsDir)) {
    try { New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null }
    catch { Write-Warning "Could not create Reports folder: $_" }
}
$ReportFile = Join-Path $ReportsDir ("AD_Health_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
# Status summary file path - suffix determined after health evaluation (_HEALTHY or _CRITICAL)
$ReportStamp = [System.IO.Path]::GetFileNameWithoutExtension($ReportFile) # e.g. AD_Health_20260308_181050

# -- IMPORT ACTIVE DIRECTORY MODULE -------------------------------------------
$ADModuleAvailable = $false
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $ADModuleAvailable = $true
    Write-Progress2 "ActiveDirectory module loaded."
} catch {
    Write-Warning "ActiveDirectory module not found. Install RSAT AD DS Tools. Some checks will be skipped."
}

Write-Host ""
Write-Host "+==============================================================+" -ForegroundColor DarkCyan
Write-Host "|        Active Directory Health Check  v$ScriptVersion              |" -ForegroundColor DarkCyan
Write-Host "+==============================================================+" -ForegroundColor DarkCyan
Write-Host ""

# ===============================================================================
# SECTION 1  -  SERVER DETAILS  (host running the script)
# ===============================================================================
Write-Progress2 "Section 1: Gathering Server Details..."
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
    $cpuNames    = ($cpu | ForEach-Object { HtmlEncode $_.Name.Trim() } | Select-Object -Unique) -join '; '

    $rows0 = @(
        @('Hostname',         (HtmlEncode $cs.Name)),
        @('Manufacturer',     (HtmlEncode $cs.Manufacturer)),
        @('Model',            (HtmlEncode $cs.Model)),
        @('Server Type',      $serverTypeBadge),
        @('Serial Number',    (HtmlEncode $bios.SerialNumber)),
        @('BIOS Version',     (HtmlEncode $bios.SMBIOSBIOSVersion)),
        @('Processors',       "$($cpu.Count) x $cpuNames"),
        @('Total RAM',        "$totalRAM_GB GB"),
        @('OS Name',          (HtmlEncode $os.Caption)),
        @('OS Version',       (HtmlEncode $os.Version)),
        @('OS Build',         (HtmlEncode $os.BuildNumber)),
        @('OS Install Date',  (HtmlEncode $os.InstallDate.ToString('yyyy-MM-dd'))),
        @('Last Boot Time',   (HtmlEncode $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss')))
    )

    $rows0Html = ($rows0 | ForEach-Object {
        "<tr><td class='td-label'>$($_[0])</td><td>$($_[1])</td></tr>"
    }) -join ''

    $Sec0Html = @"
<table class='kv-table'>
  <tbody>$rows0Html</tbody>
</table>
"@
} catch {
    $Sec0Html = "<p class='error'>Error retrieving Server Details: $(HtmlEncode $_.Exception.Message)</p>"
}

# ===============================================================================
# SECTION 2  -  DOMAIN & FOREST INFO
# ===============================================================================
Write-Progress2 "Section 2: Gathering Domain & Forest Info..."
$Sec1Html = ''
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $domain  = Get-ADDomain -ErrorAction Stop
    $forest  = Get-ADForest  -ErrorAction Stop

    $domainsInForest = if ($forest.Domains) { (@($forest.Domains) -join ', ') } else { 'N/A' }
    $sitesInForest   = if ($forest.Sites)   { (@($forest.Sites)   -join ', ') } else { 'N/A' }

    $rows = @(
        @('Domain Name',            (HtmlEncode $domain.DNSRoot)),
        @('NetBIOS Name',           (HtmlEncode $domain.NetBIOSName)),
        @('Domain DN',              (HtmlEncode $domain.DistinguishedName)),
        @('Forest Name',            (HtmlEncode $forest.Name)),
        @('Forest Functional Level',(HtmlEncode (ConvertADMode ([string]$forest.ForestMode)))),
        @('Domain Functional Level',(HtmlEncode (ConvertADMode ([string]$domain.DomainMode)))),
        @('PDC Emulator',           (HtmlEncode ([string]$domain.PDCEmulator))),
        @('RID Master',             (HtmlEncode ([string]$domain.RIDMaster))),
        @('Infrastructure Master',  (HtmlEncode ([string]$domain.InfrastructureMaster))),
        @('Schema Master',          (HtmlEncode ([string]$forest.SchemaMaster))),
        @('Domain Naming Master',   (HtmlEncode ([string]$forest.DomainNamingMaster))),
        @('Domains in Forest',      (HtmlEncode $domainsInForest)),
        @('Sites',                  (HtmlEncode $sitesInForest))
    )

    $rowsHtml = ($rows | ForEach-Object {
        "<tr><td class='td-label'>$($_[0])</td><td>$($_[1])</td></tr>"
    }) -join ''

    $Sec1Html = @"
<table class='kv-table'>
  <tbody>$rowsHtml</tbody>
</table>
"@
} catch {
    $Sec1Html = "<p class='error'>Error retrieving Domain/Forest info: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 2  -  Domain/Forest Info error: $($_.Exception.Message)")
}

# ===============================================================================
# SECTION 3  -  DOMAIN CONTROLLER INVENTORY
# ===============================================================================
Write-Progress2 "Section 3: Domain Controller Inventory..."
$Sec2Html   = ''
$AllDCs     = @()
$DCCount    = 0
$HealthyDCs = 0
$WarnDCs    = 0
$CritDCs    = 0
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $AllDCs = @(Get-ADDomainController -Filter * -ErrorAction Stop | Sort-Object Name)
    $DCCount = $AllDCs.Count

    $rows = foreach ($dc in $AllDCs) {
        $ip  = if ($dc.IPv4Address) { HtmlEncode $dc.IPv4Address } else { '<em>N/A</em>' }
        $os  = HtmlEncode $dc.OperatingSystem
        $site= HtmlEncode $dc.Site
        $gc  = if ($dc.IsGlobalCatalog) { StatusBadge 'Yes' 'green' } else { StatusBadge 'No' 'grey' }
        $ro  = if ($dc.IsReadOnly)      { StatusBadge 'RODC' 'yellow' } else { StatusBadge 'Writable' 'blue' }
        $en  = if ($dc.Enabled)         { StatusBadge 'Enabled' 'green' } else { StatusBadge 'Disabled' 'red' }
        "<tr><td>$(HtmlEncode $dc.Name)</td><td>$ip</td><td>$os</td><td>$site</td><td>$gc</td><td>$ro</td><td>$en</td></tr>"
    }
    $HealthyDCs = $DCCount  # refined in later sections

    $Sec2Html = @"
<div class='table-wrap'>
<table>
  <thead><tr>
    <th>Name</th><th>IP Address</th><th>OS Version</th><th>Site</th>
    <th>Global Catalog</th><th>Type</th><th>Status</th>
  </tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
} catch {
    $Sec2Html = "<p class='error'>Error retrieving DC Inventory: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 3  -  DC Inventory error: $($_.Exception.Message)")
}

# ===============================================================================
# SECTION 4  -  AD SERVICES STATUS PER DC
# ===============================================================================
Write-Progress2 "Section 4: AD Services Status per DC..."
$Sec3Html     = ''
$ServicesToCheck = @('NTDS','NETLOGON','W32Time','DNS','KDC')
try {
    if ($AllDCs.Count -eq 0) { throw "No Domain Controllers found or AD module unavailable." }

    $headerCols = ($ServicesToCheck | ForEach-Object { "<th>$_</th>" }) -join ''
    $rows = foreach ($dc in $AllDCs) {
        $dcName = $dc.Name
        $cols = foreach ($svc in $ServicesToCheck) {
            try {
                $s = Get-Service -ComputerName $dcName -Name $svc -ErrorAction SilentlyContinue
                if ($null -eq $s) {
                    "<td>$(StatusBadge 'N/A' 'grey')</td>"
                } elseif ($s.Status -eq 'Running') {
                    "<td>$(StatusBadge 'Running' 'green')</td>"
                } else {
                    $CriticalFindings.Add("Section 4  -  DC $dcName service $svc is $($s.Status)")
                    "<td>$(StatusBadge $s.Status.ToString() 'red')</td>"
                }
            } catch {
                "<td>$(StatusBadge 'Error' 'red')</td>"
            }
        }
        "<tr><td><strong>$(HtmlEncode $dcName)</strong></td>$($cols -join '')</tr>"
    }

    $Sec3Html = @"
<div class='table-wrap'>
<table>
  <thead><tr><th>Domain Controller</th>$headerCols</tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
} catch {
    $Sec3Html = "<p class='error'>Error checking AD Services: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 4  -  AD Services error: $($_.Exception.Message)")
}

# ===============================================================================
# SECTION 5  -  REPLICATION HEALTH
# ===============================================================================
Write-Progress2 "Section 5: Replication Health..."
$Sec4Html = ''
try {
    # repadmin /replsummary
    $replSummaryRaw = & repadmin /replsummary 2>&1
    $replSummaryText = if ($replSummaryRaw) {
        ($replSummaryRaw | ForEach-Object { $_.ToString() }) -join "`n"
    } else { 'No output from repadmin /replsummary' }

    # repadmin /showrepl
    $replShowRaw = & repadmin /showrepl 2>&1
    $replShowText = if ($replShowRaw) {
        ($replShowRaw | ForEach-Object { $_.ToString() }) -join "`n"
    } else { 'No output from repadmin /showrepl' }

    # Detect failures
    $hasFailures = ($replSummaryText -match 'fail|error' -or $replShowText -match 'fail|error')
    if ($hasFailures) { $CriticalFindings.Add("Section 5  -  Replication failures detected by repadmin.") }

    $summaryEncoded = HtmlEncode $replSummaryText
    $showreplEncoded = HtmlEncode $replShowText

    # Color-highlight lines containing failure keywords
    $summaryLines = $summaryEncoded -split "`n" | ForEach-Object {
        if ($_ -match 'fail|error') {
            "<span class='repl-fail'>$_</span>"
        } else { $_ }
    }
    $showreplLines = $showreplEncoded -split "`n" | ForEach-Object {
        if ($_ -match 'fail|error') {
            "<span class='repl-fail'>$_</span>"
        } else { $_ }
    }

    $Sec4Html = @"
<h3>repadmin /replsummary</h3>
<pre class='code-block'>$($summaryLines -join "`n")</pre>
<h3>repadmin /showrepl</h3>
<pre class='code-block'>$($showreplLines -join "`n")</pre>
"@
} catch {
    $Sec4Html = "<p class='error'>Error running repadmin: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 5  -  Replication check error: $($_.Exception.Message)")
}

# ===============================================================================
# SECTION 6  -  SYSVOL & NETLOGON SHARE
# ===============================================================================
Write-Progress2 "Section 6: SYSVOL & Netlogon Share..."
$Sec5Html = ''
try {
    if ($AllDCs.Count -eq 0) { throw "No Domain Controllers found." }

    $rows = foreach ($dc in $AllDCs) {
        $dcName = $dc.Name
        foreach ($share in @('SYSVOL','NETLOGON')) {
            $path = "\\$dcName\$share"
            try {
                $accessible = Test-Path $path -ErrorAction SilentlyContinue
                if ($accessible) {
                    $badge = StatusBadge 'Accessible' 'green'
                } else {
                    $badge = StatusBadge 'Not Accessible' 'red'
                    $CriticalFindings.Add("Section 6  -  $path is not accessible.")
                }
            } catch {
                $badge = StatusBadge 'Error' 'red'
            }
            "<tr><td>$(HtmlEncode $dcName)</td><td>$(HtmlEncode $share)</td><td>$(HtmlEncode $path)</td><td>$badge</td></tr>"
        }
    }

    $Sec5Html = @"
<div class='table-wrap'>
<table>
  <thead><tr><th>Domain Controller</th><th>Share</th><th>Path</th><th>Status</th></tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
} catch {
    $Sec5Html = "<p class='error'>Error checking SYSVOL/NETLOGON: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 6  -  SYSVOL/NETLOGON error: $($_.Exception.Message)")
}

# ===============================================================================
# SECTION 7  -  DNS HEALTH
# ===============================================================================
Write-Progress2 "Section 7: DNS Health..."
$Sec6Html = ''
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $domainDns = (Get-ADDomain -ErrorAction Stop).DNSRoot

    # SRV record test
    $srvRecord = "_ldap._tcp.$domainDns"
    $srvResult = ''
    try {
        $resolved = Resolve-DnsName -Name $srvRecord -Type SRV -ErrorAction Stop
        $srvResult = StatusBadge 'Resolved' 'green'
    } catch {
        $srvResult = StatusBadge 'Failed' 'red'
        $CriticalFindings.Add("Section 7  -  DNS SRV record $srvRecord could not be resolved.")
    }

    # DNS Zones
    $dnsZonesHtml = ''
    try {
        $zones = Get-DnsServerZone -ErrorAction SilentlyContinue
        if ($zones) {
            $zoneRows = $zones | ForEach-Object {
                $type  = HtmlEncode $_.ZoneType.ToString()
                $rt    = if ($_.IsReverseLookupZone) { StatusBadge 'Reverse' 'blue' } else { StatusBadge 'Forward' 'green' }
                $rep   = HtmlEncode $_.ReplicationScope.ToString()
                "<tr><td>$(HtmlEncode $_.ZoneName)</td><td>$type</td><td>$rt</td><td>$rep</td></tr>"
            }
            $dnsZonesHtml = @"
<h3>DNS Zones</h3>
<div class='table-wrap'>
<table>
  <thead><tr><th>Zone Name</th><th>Zone Type</th><th>Lookup Type</th><th>Replication</th></tr></thead>
  <tbody>$($zoneRows -join '')</tbody>
</table>
</div>
"@
        } else {
            $dnsZonesHtml = "<p class='warn'>No DNS zones returned (run script on a DNS server).</p>"
        }
    } catch {
        $dnsZonesHtml = "<p class='warn'>Could not retrieve DNS zones: $(HtmlEncode $_.Exception.Message)</p>"
    }

    # Forwarders
    $forwardersHtml = ''
    try {
        $fwds = Get-DnsServerForwarder -ErrorAction SilentlyContinue
        if ($fwds -and $fwds.IPAddress) {
            $fwdList = ($fwds.IPAddress | ForEach-Object { "<li>$(HtmlEncode $_.ToString())</li>" }) -join ''
            $forwardersHtml = "<h3>DNS Forwarders</h3><ul>$fwdList</ul>"
        } else {
            $forwardersHtml = "<h3>DNS Forwarders</h3><p class='warn'>No forwarders configured.</p>"
        }
    } catch {
        $forwardersHtml = "<h3>DNS Forwarders</h3><p class='warn'>Could not retrieve forwarders: $(HtmlEncode $_.Exception.Message)</p>"
    }

    $Sec6Html = @"
<p><strong>SRV Record:</strong> <code>$(HtmlEncode $srvRecord)</code> &nbsp; $srvResult</p>
$dnsZonesHtml
$forwardersHtml
"@
} catch {
    $Sec6Html = "<p class='error'>Error checking DNS Health: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 7  -  DNS Health error: $($_.Exception.Message)")
}

# ===============================================================================
# SECTION 8  -  FSMO ROLE HOLDERS
# ===============================================================================
Write-Progress2 "Section 8: FSMO Role Holders..."
$Sec7Html = ''
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $d = Get-ADDomain -ErrorAction Stop
    $f = Get-ADForest  -ErrorAction Stop

    $fsmoRoles = @(
        @{ Role = 'PDC Emulator';          Holder = $d.PDCEmulator }
        @{ Role = 'RID Master';            Holder = $d.RIDMaster }
        @{ Role = 'Infrastructure Master'; Holder = $d.InfrastructureMaster }
        @{ Role = 'Schema Master';         Holder = $f.SchemaMaster }
        @{ Role = 'Domain Naming Master';  Holder = $f.DomainNamingMaster }
    )

    $rows = foreach ($r in $fsmoRoles) {
        $holder = $r.Holder
        $reachable = $false
        try {
            $reachable = Test-Connection -ComputerName $holder -Count 1 -Quiet -ErrorAction SilentlyContinue
        } catch {}
        $badge = if ($reachable) { StatusBadge 'Reachable' 'green' } else {
            $CriticalFindings.Add("Section 8  -  FSMO $($r.Role) holder $holder is unreachable.")
            StatusBadge 'Unreachable' 'red'
        }
        "<tr><td>$(HtmlEncode $r.Role)</td><td>$(HtmlEncode $holder)</td><td>$badge</td></tr>"
    }

    $Sec7Html = @"
<div class='table-wrap'>
<table>
  <thead><tr><th>FSMO Role</th><th>Role Holder</th><th>Reachability</th></tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
} catch {
    $Sec7Html = "<p class='error'>Error checking FSMO Roles: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 8  -  FSMO check error: $($_.Exception.Message)")
}

# ===============================================================================
# SECTION 9  -  AD TRUST RELATIONSHIPS
# ===============================================================================
Write-Progress2 "Section 9: AD Trust Relationships..."
$Sec8Html = ''
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $trusts = Get-ADTrust -Filter * -ErrorAction SilentlyContinue

    if ($null -eq $trusts -or @($trusts).Count -eq 0) {
        $Sec8Html = "<p class='info'>No trusts configured in this domain.</p>"
    } else {
        $rows = foreach ($t in @($trusts)) {
            $dir    = HtmlEncode $t.Direction.ToString()
            $type   = HtmlEncode $t.TrustType.ToString()
            $status = if ($t.TrustAttributes) { StatusBadge 'Configured' 'green' } else { StatusBadge 'Unknown' 'grey' }
            "<tr><td>$(HtmlEncode $t.Name)</td><td>$dir</td><td>$type</td><td>$status</td></tr>"
        }
        $Sec8Html = @"
<div class='table-wrap'>
<table>
  <thead><tr><th>Trust Name</th><th>Direction</th><th>Trust Type</th><th>Status</th></tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
    }
} catch {
    $Sec8Html = "<p class='error'>Error checking AD Trusts: $(HtmlEncode $_.Exception.Message)</p>"
}

# ===============================================================================
# SECTION 10  -  AD TOMBSTONE & RECYCLE BIN
# ===============================================================================
Write-Progress2 "Section 10: Tombstone Lifetime & Recycle Bin..."
$Sec9Html = ''
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }

    # Tombstone lifetime
    $configNC  = (Get-ADRootDSE -ErrorAction Stop).configurationNamingContext
    $tsObj     = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$configNC" `
                              -Properties tombstoneLifetime -ErrorAction SilentlyContinue
    $tsLife    = if ($tsObj -and $tsObj.tombstoneLifetime) { "$($tsObj.tombstoneLifetime) days" } else { '60 days (default)' }

    # Recycle Bin
    $rbFeature = Get-ADOptionalFeature -Filter { Name -eq 'Recycle Bin Feature' } -ErrorAction SilentlyContinue
    $rbEnabled = $false
    if ($rbFeature) {
        $rbEnabled = ($rbFeature.EnabledScopes.Count -gt 0)
    }
    $rbBadge = if ($rbEnabled) { StatusBadge 'Enabled' 'green' } else { StatusBadge 'Disabled' 'yellow' }

    $Sec9Html = @"
<table class='kv-table'>
  <tbody>
    <tr><td class='td-label'>Tombstone Lifetime</td><td>$(HtmlEncode $tsLife)</td></tr>
    <tr><td class='td-label'>AD Recycle Bin</td><td>$rbBadge</td></tr>
  </tbody>
</table>
"@
} catch {
    $Sec9Html = "<p class='error'>Error checking Tombstone/Recycle Bin: $(HtmlEncode $_.Exception.Message)</p>"
}

# ===============================================================================
# SECTION 11  -  PRIVILEGED ACCOUNT AUDIT
# ===============================================================================
Write-Progress2 "Section 11: Privileged Account Audit..."
$Sec10Html = ''
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $staleDate = (Get-Date).AddDays(-$StaleThresholdDays)

    $privGroups = @('Domain Admins','Enterprise Admins','Schema Admins')
    $rows = foreach ($grpName in $privGroups) {
        try {
            $members = Get-ADGroupMember -Identity $grpName -Recursive -ErrorAction SilentlyContinue
            if ($null -eq $members) { $members = @() }
            $total = @($members).Count
            $stale = 0
            foreach ($m in @($members)) {
                try {
                    $user = Get-ADUser -Identity $m.SamAccountName -Properties LastLogonDate -ErrorAction SilentlyContinue
                    if ($user -and ($null -eq $user.LastLogonDate -or $user.LastLogonDate -lt $staleDate)) {
                        $stale++
                    }
                } catch {}
            }
            $staleBadge = if ($stale -gt 0) { StatusBadge "$stale stale" 'yellow' } else { StatusBadge '0 stale' 'green' }
            "<tr><td>$(HtmlEncode $grpName)</td><td>$total</td><td>$staleBadge</td></tr>"
        } catch {
            "<tr><td>$(HtmlEncode $grpName)</td><td colspan='2'><em class='warn'>Error: $(HtmlEncode $_.Exception.Message)</em></td></tr>"
        }
    }

    $Sec10Html = @"
<div class='table-wrap'>
<table>
  <thead><tr><th>Group</th><th>Total Members</th><th>Stale Accounts (&gt;$StaleThresholdDays days)</th></tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
} catch {
    $Sec10Html = "<p class='error'>Error in Privileged Account Audit: $(HtmlEncode $_.Exception.Message)</p>"
}

# ===============================================================================
# SECTION 12  -  PASSWORD POLICY
# ===============================================================================
Write-Progress2 "Section 12: Default Domain Password Policy..."
$Sec11Html = ''
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $pp = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop

    $rows = @(
        @('Min Password Length',        (HtmlEncode $pp.MinPasswordLength.ToString())),
        @('Complexity Enabled',         (HtmlEncode $pp.ComplexityEnabled.ToString())),
        @('Lockout Threshold',          (HtmlEncode $pp.LockoutThreshold.ToString())),
        @('Lockout Duration',           (HtmlEncode $pp.LockoutDuration.ToString())),
        @('Lockout Observation Window', (HtmlEncode $pp.LockoutObservationWindow.ToString())),
        @('Max Password Age',           (HtmlEncode $pp.MaxPasswordAge.ToString())),
        @('Min Password Age',           (HtmlEncode $pp.MinPasswordAge.ToString())),
        @('Password History Count',     (HtmlEncode $pp.PasswordHistoryCount.ToString())),
        @('Reversible Encryption',      (HtmlEncode $pp.ReversibleEncryptionEnabled.ToString()))
    )
    $rowsHtml = ($rows | ForEach-Object { "<tr><td class='td-label'>$($_[0])</td><td>$($_[1])</td></tr>" }) -join ''

    $Sec11Html = "<table class='kv-table'><tbody>$rowsHtml</tbody></table>"
} catch {
    $Sec11Html = "<p class='error'>Error retrieving Password Policy: $(HtmlEncode $_.Exception.Message)</p>"
}

# ===============================================================================
# SECTION 13  -  STALE OBJECTS
# ===============================================================================
Write-Progress2 "Section 13: Stale Objects..."
$Sec12Html = ''
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $staleDate = (Get-Date).AddDays(-$StaleThresholdDays)

    # -- Stale Users - fetch full details for the expandable table ---------
    $staleUserList = $null
    try {
        $staleUserList = @(Get-ADUser -Filter {
            Enabled -eq $true -and LastLogonDate -lt $staleDate
        } -Properties DisplayName, EmailAddress, Department,
                       LastLogonDate, PasswordLastSet, DistinguishedName `
          -ErrorAction SilentlyContinue | Sort-Object LastLogonDate)
    } catch { $staleUserList = $null }

    # -- Stale Computers - fetch full details ------------------------------
    $staleCompList = $null
    try {
        $staleCompList = @(Get-ADComputer -Filter {
            Enabled -eq $true -and LastLogonDate -lt $staleDate
        } -Properties DNSHostName, OperatingSystem, OperatingSystemVersion,
                       LastLogonDate, DistinguishedName `
          -ErrorAction SilentlyContinue | Sort-Object LastLogonDate)
    } catch { $staleCompList = $null }

    $staleUsers = if ($null -ne $staleUserList) { $staleUserList.Count } else { -1 }
    $staleComps = if ($null -ne $staleCompList) { $staleCompList.Count } else { -1 }

    # -- Helper: extract OU path from DistinguishedName --------------------
    function Get-OUFromDN {
        param([string]$dn)
        if ([string]::IsNullOrEmpty($dn)) { return '' }
        # Remove the first CN=... component to get the containing OU/container path
        $parts = $dn -split ',', 2
        if ($parts.Count -gt 1) { return $parts[1] } else { return $dn }
    }

    # -- Build stale-user detail table -------------------------------------
    $uDetailHtml = ''
    if ($staleUsers -gt 0) {
        $uRows = foreach ($u in $staleUserList) {
            $ll  = if ($u.LastLogonDate)   { HtmlEncode $u.LastLogonDate.ToString('yyyy-MM-dd')   } else { '<em>Never</em>' }
            $pls = if ($u.PasswordLastSet) { HtmlEncode $u.PasswordLastSet.ToString('yyyy-MM-dd') } else { '<em>Never</em>' }
            $ou  = HtmlEncode (Get-OUFromDN $u.DistinguishedName)
            "<tr>
               <td>$(HtmlEncode $u.SamAccountName)</td>
               <td>$(HtmlEncode $u.DisplayName)</td>
               <td>$(HtmlEncode $u.EmailAddress)</td>
               <td>$(HtmlEncode $u.Department)</td>
               <td>$ll</td>
               <td>$pls</td>
               <td class='stale-ou'>$ou</td>
             </tr>"
        }
        $uDetailHtml = @"
<div class='table-wrap stale-detail-wrap'>
<table>
  <thead><tr>
    <th>SAM Account</th><th>Display Name</th><th>Email</th><th>Department</th>
    <th>Last Logon</th><th>Password Last Set</th><th>OU / Container</th>
  </tr></thead>
  <tbody>$($uRows -join '')</tbody>
</table>
</div>
"@
    }

    # -- Build stale-computer detail table ---------------------------------
    $cDetailHtml = ''
    if ($staleComps -gt 0) {
        $cRows = foreach ($c in $staleCompList) {
            $ll  = if ($c.LastLogonDate) { HtmlEncode $c.LastLogonDate.ToString('yyyy-MM-dd') } else { '<em>Never</em>' }
            $ou  = HtmlEncode (Get-OUFromDN $c.DistinguishedName)
            "<tr>
               <td>$(HtmlEncode $c.Name)</td>
               <td>$(HtmlEncode $c.DNSHostName)</td>
               <td>$(HtmlEncode $c.OperatingSystem)</td>
               <td>$(HtmlEncode $c.OperatingSystemVersion)</td>
               <td>$ll</td>
               <td class='stale-ou'>$ou</td>
             </tr>"
        }
        $cDetailHtml = @"
<div class='table-wrap stale-detail-wrap'>
<table>
  <thead><tr>
    <th>Computer Name</th><th>DNS Host Name</th><th>Operating System</th>
    <th>OS Version</th><th>Last Logon</th><th>OU / Container</th>
  </tr></thead>
  <tbody>$($cRows -join '')</tbody>
</table>
</div>
"@
    }

    # -- Badges and expandable panels --------------------------------------
    $uBadgeText = if ($staleUsers -ge 0) { "$staleUsers stale users" } else { 'Error' }
    $uBadge     = if ($staleUsers -gt 0) { StatusBadge $uBadgeText 'yellow' }
                  elseif ($staleUsers -eq 0) { StatusBadge $uBadgeText 'green' }
                  else { StatusBadge $uBadgeText 'grey' }

    $cBadgeText = if ($staleComps -ge 0) { "$staleComps stale computers" } else { 'Error' }
    $cBadge     = if ($staleComps -gt 0) { StatusBadge $cBadgeText 'yellow' }
                  elseif ($staleComps -eq 0) { StatusBadge $cBadgeText 'green' }
                  else { StatusBadge $cBadgeText 'grey' }

    # Wrap each row in a <details> for click-to-expand when there are stale objects
    $uRow = if ($staleUsers -gt 0) {
        "<tr><td class='td-label'>Stale User Accounts</td><td>
          <details class='stale-details'>
            <summary class='stale-summary' aria-label='$($staleUsers) stale user accounts - click to expand'>$uBadge <span class='stale-hint' aria-hidden='true'>&#x25BE; click to view accounts</span></summary>
            $uDetailHtml
          </details>
        </td></tr>"
    } else {
        "<tr><td class='td-label'>Stale User Accounts</td><td>$uBadge</td></tr>"
    }

    $cRow = if ($staleComps -gt 0) {
        "<tr><td class='td-label'>Stale Computer Accounts</td><td>
          <details class='stale-details'>
            <summary class='stale-summary' aria-label='$($staleComps) stale computer accounts - click to expand'>$cBadge <span class='stale-hint' aria-hidden='true'>&#x25BE; click to view devices</span></summary>
            $cDetailHtml
          </details>
        </td></tr>"
    } else {
        "<tr><td class='td-label'>Stale Computer Accounts</td><td>$cBadge</td></tr>"
    }

    $Sec12Html = @"
<p>Threshold: <strong>$StaleThresholdDays days</strong> of inactivity (enabled objects only)</p>
<table class='kv-table'>
  <tbody>
    $uRow
    $cRow
  </tbody>
</table>
"@
} catch {
    $Sec12Html = "<p class='error'>Error checking Stale Objects: $(HtmlEncode $_.Exception.Message)</p>"
}

# ===============================================================================
# SECTION 14  -  DIRECTORY SERVICE EVENT LOG
# ===============================================================================
Write-Progress2 "Section 14: System Event Log..."
$Sec13Html = ''

# ---------------------------------------------------------------------------
# Advisory map: known Event IDs -> Impact level + resolution suggestion.
# Covers Directory Service, System (Service Control Manager, Disk, BugCheck,
# Networking, NTP) and Application (VSS, WMI) event sources.
# ---------------------------------------------------------------------------
$EventAdvisory = @{
    # -- Active Directory / Directory Service ------------------------------
    1000 = @{ Impact = 'Critical'; Suggestion = 'AD DS stopped unexpectedly. Immediate investigation required - this may indicate AD has shut down.' }
    1084 = @{ Impact = 'Critical'; Suggestion = 'AD replication failure. Check network connectivity and replication topology (repadmin /replsummary).' }
    1308 = @{ Impact = 'Warning';  Suggestion = 'AD replication inconsistency detected. Monitor replication status with repadmin /showrepl.' }
    1311 = @{ Impact = 'Critical'; Suggestion = 'AD replication topology broken. Run repadmin /replsummary and check DNS SRV records.' }
    1388 = @{ Impact = 'Critical'; Suggestion = 'Lingering objects detected. Run: repadmin /removelingeringobjects.' }
    1925 = @{ Impact = 'Critical'; Suggestion = 'Could not establish AD replication link. Verify DNS and network connectivity between DCs.' }
    2042 = @{ Impact = 'Critical'; Suggestion = 'Replication has not occurred within the tombstone lifetime. Immediate action required - consider authoritative restore.' }
    5807 = @{ Impact = 'Warning';  Suggestion = 'Netlogon: no DC found for a site. Check site links, subnets, and DC availability.' }
    5808 = @{ Impact = 'Warning';  Suggestion = 'Netlogon DC-locator warning. Review Active Directory Sites & Services subnet configuration.' }

    # -- Service Control Manager (System log) ------------------------------
    7000 = @{ Impact = 'Error';    Suggestion = 'Service failed to start. Check the service account, dependencies, and the Application log for detail.' }
    7001 = @{ Impact = 'Error';    Suggestion = 'Service dependency failed. Verify all prerequisite services are running.' }
    7009 = @{ Impact = 'Warning';  Suggestion = 'Service start timed out. Check system performance; consider increasing the service timeout in the registry.' }
    7011 = @{ Impact = 'Warning';  Suggestion = 'Service did not respond in time. Investigate service health; restart if necessary.' }
    7022 = @{ Impact = 'Warning';  Suggestion = 'Service hung during start. Review the service log and restart the service.' }
    7023 = @{ Impact = 'Error';    Suggestion = 'Service terminated with an error. Review event details and the Application log for error code context.' }
    7024 = @{ Impact = 'Error';    Suggestion = 'Service terminated with a service-specific error. Check the service documentation for the error code.' }
    7031 = @{ Impact = 'Error';    Suggestion = 'Service crashed and was restarted. Investigate root cause; repeated crashes may indicate a software defect.' }
    7032 = @{ Impact = 'Warning';  Suggestion = 'Service restart was attempted. Monitor the service; schedule maintenance if restarts are frequent.' }
    7034 = @{ Impact = 'Error';    Suggestion = 'Service terminated unexpectedly. Review Application/System logs and service event logs for details.' }
    7035 = @{ Impact = 'Warning';  Suggestion = 'Service control request (start/stop) sent. Confirm this was intentional; audit if unexpected.' }
    7036 = @{ Impact = 'Warning';  Suggestion = 'Service state changed. If a critical service stopped, start it immediately and investigate the cause.' }
    7038 = @{ Impact = 'Error';    Suggestion = 'Service could not log on. Verify the service account password and permissions.' }
    7040 = @{ Impact = 'Warning';  Suggestion = 'Service start type changed. Confirm this was intentional; revert if unauthorized.' }
    7045 = @{ Impact = 'Warning';  Suggestion = 'A new service was installed. Verify this is an authorized installation; investigate if unexpected.' }

    # -- Disk / Storage (System log) ---------------------------------------
    7  = @{ Impact = 'Critical'; Suggestion = 'Disk I/O error. Run chkdsk /f /r on the affected volume and check hardware health (SMART data).' }
    11 = @{ Impact = 'Critical'; Suggestion = 'Disk controller error. Check cabling, disk health (SMART), and consider replacing the disk if errors persist.' }
    15 = @{ Impact = 'Error';    Suggestion = 'Disk not ready. Ensure the disk is properly connected and not failing.' }
    51 = @{ Impact = 'Warning';  Suggestion = 'Paging operation error. Run chkdsk and review disk health; consider adding RAM to reduce paging.' }
    55 = @{ Impact = 'Critical'; Suggestion = 'NTFS filesystem corruption detected. Run chkdsk /f immediately and restore from backup if needed.' }
    57 = @{ Impact = 'Critical'; Suggestion = 'NTFS failed to flush data. Potential data loss risk - run chkdsk and inspect disk hardware immediately.' }

    # -- System / BugCheck -------------------------------------------------
    1001 = @{ Impact = 'Critical'; Suggestion = 'System crashed (BugCheck/BSOD). Analyze the dump file with WinDbg (!analyze -v). Check drivers and hardware.' }
    6008 = @{ Impact = 'Critical'; Suggestion = 'Unexpected shutdown. Verify power supply, check for BugCheck events and hardware errors in event logs.' }
    6009 = @{ Impact = 'Warning';  Suggestion = 'System version logged at boot. Normal if after maintenance; investigate if unexpected reboot.' }
    41   = @{ Impact = 'Critical'; Suggestion = 'System rebooted without clean shutdown. Check for power issues, BugCheck events, or hardware failures.' }

    # -- Network / DNS -----------------------------------------------------
    4015 = @{ Impact = 'Error';    Suggestion = 'DNS server critical error. Restart DNS Server service; check zone integrity with dnscmd /zoneprint.' }
    4016 = @{ Impact = 'Warning';  Suggestion = 'DNS internal processing error. Review DNS debug log and check zone configuration.' }
    5719 = @{ Impact = 'Critical'; Suggestion = 'No DC available to authenticate. Check DNS, network connectivity, and that the NetLogon service is running.' }
    5783 = @{ Impact = 'Critical'; Suggestion = 'Netlogon could not locate a DC. Verify DNS SRV records: dcdiag /test:dns /v.' }

    # -- Time Synchronization ----------------------------------------------
    36 = @{ Impact = 'Warning';  Suggestion = 'W32tm time sync error. Run: w32tm /config /syncfromflags:domhier /update; w32tm /resync /force.' }
    37 = @{ Impact = 'Warning';  Suggestion = 'Time-provider NtpClient: cannot reach time source. Check firewall rules for UDP 123 and NTP source reachability.' }
    38 = @{ Impact = 'Warning';  Suggestion = 'NTP time provider did not receive a timely response. Verify NTP server availability and UDP 123 connectivity.' }

    # -- VSS / Volume Shadow Copy -------------------------------------------
    8193 = @{ Impact = 'Error';    Suggestion = 'VSS call failure. Check VSS writers (vssadmin list writers); restart VSS and affected writer services.' }
    8194 = @{ Impact = 'Error';    Suggestion = 'VSS error accessing a provider. Run: vssadmin list providers; re-register VSS if needed.' }
    12293= @{ Impact = 'Error';    Suggestion = 'VSS volume error. Ensure sufficient free space on the shadow copy storage volume.' }
    12298= @{ Impact = 'Warning';  Suggestion = 'VSS pre-create snapshot failure. Verify disk space and that no VSS writer is in a failed state.' }

    # -- WMI / Application -------------------------------------------------
    10 = @{ Impact = 'Warning';  Suggestion = 'WMI event filter activation error. Run: winmgmt /resetrepository or rebuild the WMI repository if persisting.' }
}

# Logs to scan on each DC: System catches service stops, hardware errors, etc.
# Directory Service keeps the existing AD-specific coverage.
# Application catches app crashes and VSS/WMI failures.
$ScanLogs = @('System', 'Directory Service', 'Application')

try {
    if ($AllDCs.Count -eq 0) { throw "No Domain Controllers found." }
    $eventSince = (Get-Date).AddHours(-$EventLogHours)
    $allEventRows  = [System.Collections.Generic.List[string]]::new()
    $eventRowCount = 0   # count only actual event rows (not error placeholder rows)

    foreach ($dc in $AllDCs) {
        $dcName = $dc.Name
        try {
            # Collect events from System, Directory Service, and Application logs per DC
            $events = foreach ($log in $ScanLogs) {
                Get-WinEvent -ComputerName $dcName -FilterHashtable @{
                    LogName   = $log
                    StartTime = $eventSince
                    Level     = @(1, 2, 3)   # Critical=1, Error=2, Warning=3
                } -ErrorAction SilentlyContinue
            }

            if ($null -eq $events) { continue }

            # Sort all events from all logs by time descending
            $events = @($events) | Sort-Object TimeCreated -Descending

            foreach ($ev in $events) {
                $level = switch ($ev.Level) {
                    1 { 'Critical' }
                    2 { 'Error' }
                    3 { 'Warning' }
                    default { 'Info' }
                }

                $rawMsg   = if ($ev.Message) { $ev.Message } else { "Event ID $($ev.Id)" }
                $safeMsg  = HtmlEncode (TruncateMessage $rawMsg 500)
                $safeTime = HtmlEncode $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
                $safeId   = HtmlEncode $ev.Id.ToString()
                $safeSrc  = HtmlEncode $ev.ProviderName
                $safeLog  = HtmlEncode $ev.LogName
                $safeDC   = HtmlEncode $dcName

                # Impact & Suggestion - advisory map first, fall back to log-aware defaults
                $advisory   = $EventAdvisory[[int]$ev.Id]
                $impact     = if ($advisory) { $advisory.Impact } else { $level }
                $suggestion = if ($advisory) {
                    $advisory.Suggestion
                } else {
                    switch ($ev.LogName) {
                        'Directory Service' { 'Review AD event details; check replication and DC health.' }
                        'System'            { 'Review system event details and correlate with recent changes or hardware alerts.' }
                        'Application'       { 'Review application event details; check related service logs or application documentation.' }
                        default             { 'Review event details and correlate with recent changes.' }
                    }
                }

                if ($impact -eq 'Critical') {
                    $CriticalFindings.Add("Section 14 - CRITICAL event $($ev.Id) [$($ev.LogName)] on ${dcName}: $(TruncateMessage $rawMsg 100)")
                }

                $rowClass = switch ($impact) {
                    'Critical' { 'row-critical' }
                    'Error'    { 'row-error'    }
                    'Warning'  { 'row-warning'  }
                    default    { ''             }
                }

                $impactBadge = switch ($impact) {
                    'Critical' { StatusBadge 'CRITICAL - Immediate Action Required' 'red'    }
                    'Error'    { StatusBadge 'Error'                                 'red'    }
                    'Warning'  { StatusBadge 'Warning'                               'yellow' }
                    default    { StatusBadge 'Informational'                         'blue'   }
                }

                $safeSugg  = HtmlEncode $suggestion
                $safeLevel = HtmlEncode $level

                $allEventRows.Add("<tr class='$rowClass'><td>$safeDC</td><td>$safeTime</td><td>$safeLevel</td><td>$safeId</td><td>$safeLog</td><td>$safeSrc</td><td>$safeMsg</td><td>$impactBadge</td><td>$safeSugg</td></tr>")
                $eventRowCount++
            }
        } catch {
            $allEventRows.Add("<tr><td colspan='9'><em class='warn'>$(HtmlEncode $dcName) - Error: $(HtmlEncode $_.Exception.Message)</em></td></tr>")
        }
    }

    if ($allEventRows.Count -eq 0) {
        $Sec13Html = "<p class='info'>No Warning/Error/Critical events found in the last $EventLogHours hour(s) on any DC (System, Directory Service &amp; Application logs).</p>"
    } else {
        $countLabel = if ($eventRowCount -gt 0) { "$eventRowCount event(s)" } else { "No events found - see DC connection errors below" }
        $Sec13Html = @"
<p>Scanning last <strong>$EventLogHours hour(s)</strong> on all DCs (System, Directory Service &amp; Application logs). Found <strong>$countLabel</strong>.</p>
<div class='table-wrap'>
<table class='event-table'>
  <thead><tr>
    <th>DC</th><th>Time</th><th>Level</th><th>Event ID</th>
    <th>Log</th><th>Source</th><th>Message</th><th>Impact</th><th>Suggestion</th>
  </tr></thead>
  <tbody>$($allEventRows -join '')</tbody>
</table>
</div>
"@
    }
} catch {
    $Sec13Html = "<p class='error'>Error reading System Event Log: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 14  -  Event Log error: $($_.Exception.Message)")
}

# ===============================================================================
# SECTION 15  -  WINDOWS UPDATE STATUS
# ===============================================================================
Write-Progress2 "Section 15: Windows Update Status..."
$Sec14Html = ''
try {
    if ($AllDCs.Count -eq 0) { throw "No Domain Controllers found." }
    $updateSince = (Get-Date).AddDays(-30)
    $allDCUpdateHtml = [System.Collections.Generic.List[string]]::new()

    foreach ($dc in $AllDCs) {
        $dcName = $dc.Name
        $hotfixRows = [System.Collections.Generic.List[string]]::new()

        # Recent installed hotfixes (last 30 days)
        try {
            $hotfixes = Get-HotFix -ComputerName $dcName -ErrorAction SilentlyContinue |
                        Where-Object { $_.InstalledOn -ge $updateSince } |
                        Sort-Object InstalledOn -Descending
            if ($hotfixes) {
                foreach ($hf in $hotfixes) {
                    $installed = if ($hf.InstalledOn) { HtmlEncode $hf.InstalledOn.ToString('yyyy-MM-dd') } else { 'Unknown' }
                    $hotfixRows.Add("<tr><td>$(HtmlEncode $hf.HotFixID)</td><td>$(HtmlEncode $hf.Description)</td><td>$installed</td></tr>")
                }
            } else {
                $hotfixRows.Add("<tr><td colspan='3'><em>No hotfixes installed in the last 30 days.</em></td></tr>")
            }
        } catch {
            $hotfixRows.Add("<tr><td colspan='3'><em class='warn'>Could not retrieve hotfixes: $(HtmlEncode $_.Exception.Message)</em></td></tr>")
        }

        # Pending updates via COM (Microsoft.Update.Session)  -  remote invocation via Invoke-Command
        $pendingHtml = ''
        try {
            $pendingResult = Invoke-Command -ComputerName $dcName -ScriptBlock {
                try {
                    $session  = New-Object -ComObject 'Microsoft.Update.Session' -ErrorAction Stop
                    $searcher = $session.CreateUpdateSearcher()
                    $result   = $searcher.Search("IsInstalled=0 and IsHidden=0")
                    $updates  = $result.Updates
                    $out = [System.Collections.Generic.List[hashtable]]::new()
                    for ($i = 0; $i -lt $updates.Count; $i++) {
                        $u = $updates.Item($i)
                        $out.Add(@{
                            Title = $u.Title
                            Severity = if ($u.MsrcSeverity) { $u.MsrcSeverity } else { 'Unknown' }
                        })
                    }
                    return $out
                } catch {
                    return $null
                }
            } -ErrorAction SilentlyContinue

            if ($null -eq $pendingResult) {
                # WinRM unavailable - attempt a direct local COM fallback.
                # This succeeds when the script is running on the same DC being checked.
                $localFallbackDone = $false
                try {
                    # Compare FQDN of $dcName against local computer name to avoid
                    # false matches (e.g. "DC1" vs "DC10").
                    $localFQDN   = [System.Net.Dns]::GetHostEntry('').HostName
                    $localMachine = ($dcName -eq $env:COMPUTERNAME) -or
                                    ($dcName -eq $localFQDN)
                    if ($localMachine) {
                        $localSession  = New-Object -ComObject 'Microsoft.Update.Session' -ErrorAction Stop
                        $localSearcher = $localSession.CreateUpdateSearcher()
                        $localResult   = $localSearcher.Search("IsInstalled=0 and IsHidden=0")
                        $localFallbackDone = $true
                        if ($localResult.Updates.Count -eq 0) {
                            $pendingHtml = "<p class='info'>$(StatusBadge 'Windows is Up-to-Date' 'green')</p>"
                        } else {
                            $localRows = foreach ($u in $localResult.Updates) {
                                $sev    = if ($u.MsrcSeverity) { $u.MsrcSeverity } else { 'Unknown' }
                                $sevBadge = switch ($sev) {
                                    'Critical'  { StatusBadge 'Critical'  'red'    }
                                    'Important' { StatusBadge 'Important' 'yellow' }
                                    default     { StatusBadge $sev        'grey'   }
                                }
                                if ($sev -in @('Critical','Important')) {
                                    $CriticalFindings.Add("Section 15  -  DC $dcName has pending $sev update: $($u.Title)")
                                }
                                "<tr><td>$(HtmlEncode $u.Title)</td><td>$sevBadge</td></tr>"
                            }
                            $pendingHtml = @"
<h4>Pending Updates ($($localResult.Updates.Count))</h4>
<div class='table-wrap'>
<table>
  <thead><tr><th>Update Title</th><th>Severity</th></tr></thead>
  <tbody>$($localRows -join '')</tbody>
</table>
</div>
"@
                        }
                    }
                } catch {
                    # Local COM fallback failed (e.g. Windows Update service disabled or COM error).
                    # $localFallbackDone stays $false so the warning is shown below.
                    Write-Warning "Local Windows Update COM fallback failed for ${dcName}: $_"
                }
                if (-not $localFallbackDone) {
                    $pendingHtml = "<p class='warn'>Could not query pending updates (WinRM may be unavailable). Run this script directly on $(HtmlEncode $dcName) for accurate results.</p>"
                }
            } elseif (@($pendingResult).Count -eq 0) {
                $pendingHtml = "<p class='info'>$(StatusBadge 'No pending updates' 'green')</p>"
            } else {
                $pendingRows = foreach ($u in @($pendingResult)) {
                    $sev = HtmlEncode $u.Severity
                    $sevBadge = switch ($u.Severity) {
                        'Critical'  { StatusBadge 'Critical'  'red'    }
                        'Important' { StatusBadge 'Important' 'yellow' }
                        default     { StatusBadge $u.Severity 'grey'   }
                    }
                    if ($u.Severity -in @('Critical','Important')) {
                        $CriticalFindings.Add("Section 15  -  DC $dcName has pending $($u.Severity) update: $($u.Title)")
                    }
                    "<tr><td>$(HtmlEncode $u.Title)</td><td>$sevBadge</td></tr>"
                }
                $pendingHtml = @"
<h4>Pending Updates ($(@($pendingResult).Count))</h4>
<div class='table-wrap'>
<table>
  <thead><tr><th>Update Title</th><th>Severity</th></tr></thead>
  <tbody>$($pendingRows -join '')</tbody>
</table>
</div>
"@
            }
        } catch {
            $pendingHtml = "<p class='warn'>Pending update check error: $(HtmlEncode $_.Exception.Message)</p>"
        }

        $allDCUpdateHtml.Add(@"
<details>
  <summary><strong>$(HtmlEncode $dcName)</strong></summary>
  <h4>Recently Installed Hotfixes (Last 30 Days)</h4>
  <div class='table-wrap'>
  <table>
    <thead><tr><th>HotFix ID</th><th>Description</th><th>Installed On</th></tr></thead>
    <tbody>$($hotfixRows -join '')</tbody>
  </table>
  </div>
  $pendingHtml
</details>
"@)
    }

    $Sec14Html = $allDCUpdateHtml -join ''
} catch {
    $Sec14Html = "<p class='error'>Error checking Windows Update Status: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 15  -  Windows Update check error: $($_.Exception.Message)")
}

# ===============================================================================
# SUMMARY BAR COUNTS
# ===============================================================================
$domainName = ''
$forestLevel = ''
try {
    if ($ADModuleAvailable) {
        $domainName  = [string](Get-ADDomain -ErrorAction SilentlyContinue).DNSRoot
        $forestLevel = ConvertADMode ([string](Get-ADForest -ErrorAction SilentlyContinue).ForestMode)
    }
} catch {}

# Recalculate summary counts based on critical findings
$CritCount = [int]$CriticalFindings.Count
if ($CritCount -gt 0) { $CritDCs = [Math]::Min($CritCount, [int]$DCCount) }
$HealthyDCs = [Math]::Max(0, [int]$DCCount - $CritDCs)

Write-Progress2 "Building HTML report..."

# ===============================================================================
# COMPANY LOGO HTML
# ===============================================================================
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

# ===============================================================================
# HELPER: Build collapsible section
# ===============================================================================
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

$EndTime     = Get-Date
$Duration    = ($EndTime - $StartTime).ToString('hh\:mm\:ss')
$ReportDate  = $EndTime.ToString('dddd, dd MMMM yyyy HH:mm:ss')

$AuthorLink = if (-not [string]::IsNullOrWhiteSpace($CompanyWebsite)) {
    "<a href='$(HtmlEncode $CompanyWebsite)' target='_blank' style='color:var(--link);'>$(HtmlEncode $AuthorName)</a>"
} else {
    HtmlEncode $AuthorName
}

# ===============================================================================
# BUILD FULL HTML
# ===============================================================================
$HtmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AD Health Check Report  -  $(HtmlEncode $domainName)</title>
<style>
/* -- CSS VARIABLES -- */
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

/* -- RESET & BASE -- */
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
em { font-style: italic; }

/* -- LAYOUT -- */
.page-wrap  { max-width: 1400px; margin: 0 auto; padding: 0 16px 40px; }
.header     { background: var(--head-bg); border-bottom: 1px solid var(--border);
              padding: 16px 24px; display: flex; align-items: center;
              justify-content: space-between; flex-wrap: wrap; gap: 12px; }
.header-left  { display: flex; align-items: center; gap: 16px; }
.logo-wrap img { max-height: 52px; }
.header-title h1 { font-size: 1.4rem; color: var(--text); }
.header-title p  { font-size: .8rem; color: var(--muted); margin:0; }

/* -- THEME TOGGLE -- */
.theme-toggle { cursor: pointer; background: var(--card); border: 1px solid var(--border);
                color: var(--text); border-radius: 20px; padding: 6px 14px;
                font-size: 12px; display: flex; align-items: center; gap: 6px; }
.theme-toggle:hover { background: var(--th-bg); }

/* -- SUMMARY BAR -- */
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

/* -- SECTION CARDS -- */
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

/* -- TABLES -- */
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

/* -- EVENT ROW COLORS -- */
.row-critical { background: rgba(218,54,51,.15) !important; }
.row-error    { background: rgba(218,54,51,.08) !important; }
.row-warning  { background: rgba(210,153,34,.12) !important; }

/* -- BADGES -- */
.badge { display: inline-block; font-size: .7rem; font-weight: 600; padding: 2px 8px;
         border-radius: 20px; color: #fff; white-space: nowrap; }

/* -- CODE BLOCK -- */
.code-block { background: var(--pre-bg); border: 1px solid var(--border); border-radius: 6px;
              padding: 12px; overflow-x: auto; font-size: 12px; font-family: 'SFMono-Regular',
              Consolas, monospace; white-space: pre; color: var(--text); line-height: 1.5; }
.repl-fail  { color: #f85149; font-weight: 600; }

/* -- EVENT LOG TABLE  -  constrain wide message/suggestion columns -- */
.event-table td:nth-child(6),
.event-table td:nth-child(8) {
  max-width: 280px;
  word-break: break-word;
  white-space: normal;
}

/* -- STALE OBJECTS - expandable detail panels -- */
.stale-details { display: block; }
.stale-summary { list-style: none; cursor: pointer; display: inline-flex;
                 align-items: center; gap: 6px; user-select: none; }
.stale-summary::-webkit-details-marker { display: none; }
.stale-summary:focus { outline: 2px solid var(--blue); outline-offset: 2px;
                       border-radius: 3px; }
.stale-hint { font-size: .72rem; color: var(--muted); font-style: italic; }
.stale-detail-wrap { margin-top: 10px; }
.stale-ou { font-size: .75rem; color: var(--muted); word-break: break-all; }

/* -- MESSAGE CLASSES -- */
.error { color: #f85149; padding: 8px 12px; background: rgba(248,81,73,.1);
         border-left: 3px solid #f85149; border-radius: 4px; }
.warn  { color: #d29922; }
.info  { color: var(--muted); }

/* -- FOOTER -- */
.footer { border-top: 1px solid var(--border); padding: 20px 0;
          margin-top: 24px; color: var(--muted); font-size: .8rem;
          display: flex; justify-content: space-between; flex-wrap: wrap; gap: 8px; }

/* -- RESPONSIVE -- */
@media (max-width: 768px) {
  .summary-bar { gap: 8px; }
  .stat-card   { flex: 1 1 100px; }
  .stat-count  { font-size: 1.5rem; }
  .header      { flex-direction: column; align-items: flex-start; }
}
</style>
</head>
<body data-theme="dark">

<!-- === HEADER === -->
<div class="header">
  <div class="header-left">
    $LogoHtml
    <div class="header-title">
      <h1>&#x26A1; Active Directory Health Check</h1>
      <p>$(HtmlEncode $domainName) &nbsp;|&nbsp; Generated: $(HtmlEncode $ReportDate)</p>
    </div>
  </div>
  <button class="theme-toggle" onclick="toggleTheme()" title="Toggle Dark/Light mode">
    <span id="theme-icon">&#x2600;&#xFE0F;</span> Toggle Theme
  </button>
</div>

<!-- === SUMMARY BAR === -->
<div class="page-wrap">
<div class="summary-bar">
  <div class="stat-card">
    <div class="stat-count c-blue">$DCCount</div>
    <div class="stat-label">Total DCs</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-green">$HealthyDCs</div>
    <div class="stat-label">Healthy</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-yellow">$WarnDCs</div>
    <div class="stat-label">Warnings</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-red">$CritCount</div>
    <div class="stat-label">Critical Findings</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-muted" style="font-size:1rem; padding-top:8px;">$(HtmlEncode $domainName)</div>
    <div class="stat-label">Domain</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-muted" style="font-size:1rem; padding-top:8px;">$(HtmlEncode $forestLevel)</div>
    <div class="stat-label">Forest Level</div>
  </div>
</div>

<!-- === 15 SECTIONS === -->
$(BuildSection 1  'Server Details'                  $Sec0Html  ($Sec0Html  -match 'error')   $true)
$(BuildSection 2  'Domain & Forest Info'            $Sec1Html  ($Sec1Html  -match 'error')   $true)
$(BuildSection 3  'Domain Controller Inventory'     $Sec2Html  ($Sec2Html  -match 'error')   $true)
$(BuildSection 4  'AD Services Status per DC'       $Sec3Html  ($Sec3Html  -match 'error')   $false)
$(BuildSection 5  'Replication Health'              $Sec4Html  ($Sec4Html  -match 'error')   $false)
$(BuildSection 6  'SYSVOL & Netlogon Shares'        $Sec5Html  ($Sec5Html  -match 'error')   $false)
$(BuildSection 7  'DNS Health'                      $Sec6Html  ($Sec6Html  -match 'error')   $false)
$(BuildSection 8  'FSMO Role Holders'               $Sec7Html  ($Sec7Html  -match 'error')   $false)
$(BuildSection 9  'AD Trust Relationships'          $Sec8Html  ($Sec8Html  -match 'error')   $false)
$(BuildSection 10 'AD Tombstone & Recycle Bin'      $Sec9Html  ($Sec9Html  -match 'error')   $false)
$(BuildSection 11 'Privileged Account Audit'        $Sec10Html ($Sec10Html -match 'error')   $false)
$(BuildSection 12 'Default Domain Password Policy'  $Sec11Html ($Sec11Html -match 'error')   $false)
$(BuildSection 13 'Stale Objects'                   $Sec12Html ($Sec12Html -match 'error')   $false)
$(BuildSection 14 'System Event Log'               $Sec13Html ($Sec13Html -match 'error')   $false)
$(BuildSection 15 'Windows Update Status'           $Sec14Html ($Sec14Html -match 'error')   $false)

<!-- === FOOTER === -->
<div class="footer">
  <div>
    <strong>AD Health Check v$ScriptVersion</strong> &nbsp;|&nbsp;
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

# ===============================================================================
# WRITE REPORT FILE
# ===============================================================================
try {
    [System.IO.File]::WriteAllText($ReportFile, $HtmlReport, [System.Text.Encoding]::UTF8)
    Write-Host ""
    Write-Host "  [OK] Report saved to: $ReportFile" -ForegroundColor Green
} catch {
    Write-Warning "Failed to write report: $_"
}

# ===============================================================================
# STATUS SUMMARY FILE  (plain-text companion - _HEALTHY.txt or _CRITICAL.txt)
# ===============================================================================
$isCritical   = $CriticalFindings.Count -gt 0
$statusSuffix = if ($isCritical) { '_CRITICAL' } else { '_HEALTHY' }
$StatusFile   = Join-Path $ReportsDir ($ReportStamp + $statusSuffix + '.txt')

$separator = '=' * 70

if ($isCritical) {
    $statusContent  = "$separator`r`n"
    $statusContent += " AD HEALTH CHECK  -  *** CRITICAL ALERT ***`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " Status       : CRITICAL`r`n"
    $statusContent += " Domain       : $domainName`r`n"
    $statusContent += " Forest Level : $forestLevel`r`n"
    $statusContent += " Total DCs    : $DCCount`r`n"
    $statusContent += " Generated    : $ReportDate`r`n"
    $statusContent += " Duration     : $Duration`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += "`r`n CRITICAL FINDINGS ($($CriticalFindings.Count)):`r`n`r`n"
    $statusContent += (@($CriticalFindings) | ForEach-Object { "  [!] $_" }) -join "`r`n"
    $statusContent += "`r`n`r`n$separator`r`n"
    $statusContent += " Full HTML report : $ReportFile`r`n"
    $statusContent += " Status file      : $StatusFile`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " AD Health Check v$ScriptVersion  by $AuthorName`r`n"
    $statusContent += "$separator`r`n"
} else {
    $statusContent  = "$separator`r`n"
    $statusContent += " AD HEALTH CHECK  -  HEALTHY STATE`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " Status       : HEALTHY`r`n"
    $statusContent += " Domain       : $domainName`r`n"
    $statusContent += " Forest Level : $forestLevel`r`n"
    $statusContent += " Total DCs    : $DCCount`r`n"
    $statusContent += " Generated    : $ReportDate`r`n"
    $statusContent += " Duration     : $Duration`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += "`r`n No critical findings, errors, or warnings were detected.`r`n"
    $statusContent += " Your Active Directory environment is in a healthy state.`r`n"
    $statusContent += "`r`n$separator`r`n"
    $statusContent += " Full HTML report : $ReportFile`r`n"
    $statusContent += " Status file      : $StatusFile`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " AD Health Check v$ScriptVersion  by $AuthorName`r`n"
    $statusContent += "$separator`r`n"
}

if ($EnableStatusFile) {
    try {
        # CRLF line endings are intentional: this is a Windows Server script and .txt files
        # opened in Notepad/Explorer display correctly with CRLF. Consistent with HTML report write.
        [System.IO.File]::WriteAllText($StatusFile, $statusContent, [System.Text.Encoding]::UTF8)
        $statusColor = if ($isCritical) { 'Red' } else { 'Green' }
        Write-Host "  [OK] Status file saved to: $StatusFile" -ForegroundColor $statusColor
    } catch {
        Write-Warning "Failed to write status file: $_"
    }
}

Write-Host ""
Write-Host "===============================================================" -ForegroundColor DarkCyan
Write-Host "  AD Health Check complete.  Duration: $Duration" -ForegroundColor DarkCyan
Write-Host "  Report     : $ReportFile"  -ForegroundColor Yellow
if ($EnableStatusFile) {
    $statusLabel = if ($isCritical) { 'Status (CRITICAL)' } else { 'Status (HEALTHY)' }
    Write-Host "  $statusLabel : $StatusFile" -ForegroundColor $(if ($isCritical) { 'Red' } else { 'Green' })
    Write-Host ""
    Write-Host "  To send email alerts, run: .\AD_HealthCheck_EmailAlert.ps1  (configure SMTP settings inside first)" -ForegroundColor Cyan
}
Write-Host "===============================================================" -ForegroundColor DarkCyan
Write-Host ""
