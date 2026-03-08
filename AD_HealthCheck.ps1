successfully downloaded text file (SHA: 1ec6df7b012d09084230f888103c27860487843f)#Requires -Version 5.1
<#
.SYNOPSIS
    Comprehensive Active Directory Health Check Script

.DESCRIPTION
    Performs a thorough health check of the Active Directory environment and
    exports the results to a single self-contained HTML file with a professional
    dark/light-themed dashboard UI.

.NOTES
    Version    : 1.0.0
    Author     : Tushar Gudde
    Requires   : PowerShell 5.1+, RSAT AD DS Tools (ActiveDirectory module)
    Permissions: Domain Admin or equivalent
#>

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
# Company Branding
$CompanyLogoURL    = ''                          # URL/path to company logo (PNG/SVG). Leave blank to skip.
$CompanyWebsite    = ''                          # Company website URL for logo hyperlink.

# Author
$AuthorName        = 'Tushar Gudde'              # Author name shown in footer

# Event Log Settings
$EventLogHours     = 2                           # How many hours back to scan events

# Stale Object Threshold
$StaleThresholdDays = 90                         # Days of inactivity before flagged stale

# Email Alert Configuration
$EnableEmailAlert       = $false
$SMTPServer             = 'smtp.yourdomain.com'
$SMTPPort               = 587
$SMTPFrom               = 'ad-healthcheck@yourdomain.com'
$SMTPTo                 = @('admin@yourdomain.com')
$SMTPSubject            = 'AD Health Check - CRITICAL ALERT'
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
$ReportFile = Join-Path $ReportsDir ("AD_Health_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

# ── IMPORT ACTIVE DIRECTORY MODULE ───────────────────────────────────────────
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

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1  -  DOMAIN & FOREST INFO
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 1: Gathering Domain & Forest Info..."
$Sec1Html = ''
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $domain  = Get-ADDomain -ErrorAction Stop
    $forest  = Get-ADForest  -ErrorAction Stop

    $rows = @(
        @('Domain Name',            (HtmlEncode $domain.DNSRoot))
        @('NetBIOS Name',           (HtmlEncode $domain.NetBIOSName))
        @('Domain DN',              (HtmlEncode $domain.DistinguishedName))
        @('Forest Name',            (HtmlEncode $forest.Name))
        @('Forest Functional Level',(HtmlEncode $forest.ForestMode.ToString()))
        @('Domain Functional Level',(HtmlEncode $domain.DomainMode.ToString()))
        @('PDC Emulator',           (HtmlEncode $domain.PDCEmulator))
        @('RID Master',             (HtmlEncode $domain.RIDMaster))
        @('Infrastructure Master',  (HtmlEncode $domain.InfrastructureMaster))
        @('Schema Master',          (HtmlEncode $forest.SchemaMaster))
        @('Domain Naming Master',   (HtmlEncode $forest.DomainNamingMaster))
        @('Domains in Forest',      (HtmlEncode ($forest.Domains -join ', ')))
        @('Sites',                  (HtmlEncode ($forest.Sites -join ', ')))
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
    $CriticalFindings.Add("Section 1  -  Domain/Forest Info error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2  -  DOMAIN CONTROLLER INVENTORY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 2: Domain Controller Inventory..."
$Sec2Html   = ''
$AllDCs     = @()
$DCCount    = 0
$HealthyDCs = 0
$WarnDCs    = 0
$CritDCs    = 0
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $AllDCs = Get-ADDomainController -Filter * -ErrorAction Stop | Sort-Object Name
    $DCCount = [int]$AllDCs.Count

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
    $CriticalFindings.Add("Section 2  -  DC Inventory error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3  -  AD SERVICES STATUS PER DC
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 3: AD Services Status per DC..."
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
                    StatusBadge 'N/A' 'grey'
                } elseif ($s.Status -eq 'Running') {
                    StatusBadge 'Running' 'green'
                } else {
                    $CriticalFindings.Add("Section 3  -  DC $dcName service $svc is $($s.Status)")
                    StatusBadge $s.Status.ToString() 'red'
                }
            } catch {
                StatusBadge 'Error' 'red'
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
    $CriticalFindings.Add("Section 3  -  AD Services error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4  -  REPLICATION HEALTH
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 4: Replication Health..."
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
    if ($hasFailures) { $CriticalFindings.Add("Section 4  -  Replication failures detected by repadmin.") }

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
    $CriticalFindings.Add("Section 4  -  Replication check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5  -  SYSVOL & NETLOGON SHARE
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 5: SYSVOL & Netlogon Share..."
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
                    $CriticalFindings.Add("Section 5  -  $path is not accessible.")
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
    $CriticalFindings.Add("Section 5  -  SYSVOL/NETLOGON error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6  -  DNS HEALTH
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 6: DNS Health..."
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
        $CriticalFindings.Add("Section 6  -  DNS SRV record $srvRecord could not be resolved.")
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
    $CriticalFindings.Add("Section 6  -  DNS Health error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7  -  FSMO ROLE HOLDERS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 7: FSMO Role Holders..."
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
            $CriticalFindings.Add("Section 7  -  FSMO $($r.Role) holder $holder is unreachable.")
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
    $CriticalFindings.Add("Section 7  -  FSMO check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 8  -  AD TRUST RELATIONSHIPS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 8: AD Trust Relationships..."
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

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 9  -  AD TOMBSTONE & RECYCLE BIN
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 9: Tombstone Lifetime & Recycle Bin..."
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

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 10  -  PRIVILEGED ACCOUNT AUDIT
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 10: Privileged Account Audit..."
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

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 11  -  PASSWORD POLICY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 11: Default Domain Password Policy..."
$Sec11Html = ''
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $pp = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop

    $rows = @(
        @('Min Password Length',    (HtmlEncode $pp.MinPasswordLength.ToString()))
        @('Complexity Enabled',     (HtmlEncode $pp.ComplexityEnabled.ToString()))
        @('Lockout Threshold',      (HtmlEncode $pp.LockoutThreshold.ToString()))
        @('Lockout Duration',       (HtmlEncode $pp.LockoutDuration.ToString()))
        @('Lockout Observation Window', (HtmlEncode $pp.LockoutObservationWindow.ToString()))
        @('Max Password Age',       (HtmlEncode $pp.MaxPasswordAge.ToString()))
        @('Min Password Age',       (HtmlEncode $pp.MinPasswordAge.ToString()))
        @('Password History Count', (HtmlEncode $pp.PasswordHistoryCount.ToString()))
        @('Reversible Encryption',  (HtmlEncode $pp.ReversibleEncryptionEnabled.ToString()))
    )
    $rowsHtml = ($rows | ForEach-Object { "<tr><td class='td-label'>$($_[0])</td><td>$($_[1])</td></tr>" }) -join ''

    $Sec11Html = "<table class='kv-table'><tbody>$rowsHtml</tbody></table>"
} catch {
    $Sec11Html = "<p class='error'>Error retrieving Password Policy: $(HtmlEncode $_.Exception.Message)</p>"
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 12  -  STALE OBJECTS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 12: Stale Objects..."
$Sec12Html = ''
try {
    if (-not $ADModuleAvailable) { throw "ActiveDirectory module not available." }
    $staleDate = (Get-Date).AddDays(-$StaleThresholdDays)

    $staleUsers = 0
    $staleComps = 0

    try {
        $staleUsers = @(Get-ADUser -Filter {
            Enabled -eq $true -and LastLogonDate -lt $staleDate
        } -Properties LastLogonDate -ErrorAction SilentlyContinue).Count
    } catch { $staleUsers = -1 }

    try {
        $staleComps = @(Get-ADComputer -Filter {
            Enabled -eq $true -and LastLogonDate -lt $staleDate
        } -Properties LastLogonDate -ErrorAction SilentlyContinue).Count
    } catch { $staleComps = -1 }

    $uBadge = if ($staleUsers -gt 0) { StatusBadge "$staleUsers stale users" 'yellow' }
              elseif ($staleUsers -eq 0) { StatusBadge '0 stale users' 'green' }
              else { StatusBadge 'Error' 'grey' }

    $cBadge = if ($staleComps -gt 0) { StatusBadge "$staleComps stale computers" 'yellow' }
              elseif ($staleComps -eq 0) { StatusBadge '0 stale computers' 'green' }
              else { StatusBadge 'Error' 'grey' }

    $Sec12Html = @"
<p>Threshold: <strong>$StaleThresholdDays days</strong> of inactivity (enabled objects only)</p>
<table class='kv-table'>
  <tbody>
    <tr><td class='td-label'>Stale User Accounts</td><td>$uBadge</td></tr>
    <tr><td class='td-label'>Stale Computer Accounts</td><td>$cBadge</td></tr>
  </tbody>
</table>
"@
} catch {
    $Sec12Html = "<p class='error'>Error checking Stale Objects: $(HtmlEncode $_.Exception.Message)</p>"
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 13  -  DIRECTORY SERVICE EVENT LOG
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 13: Directory Service Event Log..."
$Sec13Html = ''

# Map of known AD Event IDs to suggestions/impact
$EventAdvisory = @{
    1000 = @{ Impact = 'Critical'; Suggestion = 'AD DS stopped unexpectedly. Immediate investigation required  -  this may cause AD shutdown.' }
    1084 = @{ Impact = 'Critical'; Suggestion = 'Replication failure. Check network connectivity and replication topology.' }
    1308 = @{ Impact = 'Warning';  Suggestion = 'Replication warning  -  inconsistency detected. Monitor replication status.' }
    1311 = @{ Impact = 'Critical'; Suggestion = 'Replication topology broken. Run repadmin /replsummary to investigate.' }
    1388 = @{ Impact = 'Critical'; Suggestion = 'Lingering objects detected. Run repadmin /removelingeringobjects.' }
    1925 = @{ Impact = 'Critical'; Suggestion = 'Could not establish replication link. Check DNS and network connectivity.' }
    2042 = @{ Impact = 'Critical'; Suggestion = 'Replication not occurred in tombstone lifetime. Immediate action required.' }
    5807 = @{ Impact = 'Warning';  Suggestion = 'Netlogon detected no DC for a site. Check site links and DC availability.' }
    5808 = @{ Impact = 'Warning';  Suggestion = 'Netlogon warning regarding DC locator. Review site and subnet configuration.' }
}

try {
    if ($AllDCs.Count -eq 0) { throw "No Domain Controllers found." }
    $eventSince = (Get-Date).AddHours(-$EventLogHours)
    $allEventRows = [System.Collections.Generic.List[string]]::new()

    foreach ($dc in $AllDCs) {
        $dcName = $dc.Name
        try {
            $events = Get-WinEvent -ComputerName $dcName -FilterHashtable @{
                LogName   = 'Directory Service'
                StartTime = $eventSince
                Level     = @(1, 2, 3)   # Critical=1, Error=2, Warning=3
            } -ErrorAction SilentlyContinue

            if ($null -eq $events) { continue }

            foreach ($ev in $events) {
                $level = switch ($ev.Level) {
                    1 { 'Critical' }
                    2 { 'Error' }
                    3 { 'Warning' }
                    default { 'Info' }
                }

                # Use $ev.Message for human-readable text; encode and truncate
                $rawMsg   = if ($ev.Message) { $ev.Message } else { "Event ID $($ev.Id)" }
                $safeMsg  = HtmlEncode (TruncateMessage $rawMsg 500)
                $safeTime = HtmlEncode $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
                $safeId   = HtmlEncode $ev.Id.ToString()
                $safeSrc  = HtmlEncode $ev.ProviderName
                $safeDC   = HtmlEncode $dcName

                # Impact & Suggestion
                $advisory = $EventAdvisory[$ev.Id]
                $impact     = if ($advisory) { $advisory.Impact }     else { $level }
                $suggestion = if ($advisory) { $advisory.Suggestion } else { 'Review event details and correlate with recent changes.' }

                if ($impact -eq 'Critical') {
                    $CriticalFindings.Add("Section 13 - CRITICAL event $($ev.Id) on ${dcName}: $(TruncateMessage $rawMsg 100)")
                }

                $rowClass = switch ($impact) {
                    'Critical'      { 'row-critical' }
                    'Error'         { 'row-error'    }
                    'Warning'       { 'row-warning'  }
                    default         { ''              }
                }

                $impactBadge = switch ($impact) {
                    'Critical' { StatusBadge 'CRITICAL - Immediate Action Required' 'red'    }
                    'Error'    { StatusBadge 'Error'                                 'red'    }
                    'Warning'  { StatusBadge 'Warning'                               'yellow' }
                    default    { StatusBadge 'Informational'                         'blue'   }
                }

                $allEventRows.Add(@"
<tr class='$rowClass'>
  <td>$safeDC</td>
  <td>$safeTime</td>
  <td>$(HtmlEncode $level)</td>
  <td>$safeId</td>
  <td>$safeSrc</td>
  <td>$safeMsg</td>
  <td>$impactBadge</td>
  <td>$(HtmlEncode $suggestion)</td>
</tr>
"@)
            }
        } catch {
            $allEventRows.Add("<tr><td colspan='8'><em class='warn'>$(HtmlEncode $dcName) - Error: $(HtmlEncode $_.Exception.Message)</em></td></tr>")
        }
    }

    if ($allEventRows.Count -eq 0) {
        $Sec13Html = "<p class='info'>No Warning/Error/Critical events found in the last $EventLogHours hour(s) on any DC.</p>"
    } else {
        $Sec13Html = @"
<p>Scanning last <strong>$EventLogHours hour(s)</strong> on all DCs. Found <strong>$($allEventRows.Count)</strong> event(s).</p>
<div class='table-wrap'>
<table>
  <thead><tr>
    <th>DC</th><th>Time</th><th>Level</th><th>Event ID</th>
    <th>Source</th><th>Message</th><th>Impact</th><th>Suggestion</th>
  </tr></thead>
  <tbody>$($allEventRows -join '')</tbody>
</table>
</div>
"@
    }
} catch {
    $Sec13Html = "<p class='error'>Error reading Directory Service Event Log: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 13  -  Event Log error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 14  -  WINDOWS UPDATE STATUS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 14: Windows Update Status..."
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
                $pendingHtml = "<p class='warn'>Could not query pending updates (WinRM may be unavailable).</p>"
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
                        $CriticalFindings.Add("Section 14  -  DC $dcName has pending $($u.Severity) update: $($u.Title)")
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
    $CriticalFindings.Add("Section 14  -  Windows Update check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SUMMARY BAR COUNTS
# ═══════════════════════════════════════════════════════════════════════════════
$domainName = ''
$forestLevel = ''
try {
    if ($ADModuleAvailable) {
        $domainName  = (Get-ADDomain -ErrorAction SilentlyContinue).DNSRoot
        $forestLevel = (Get-ADForest -ErrorAction SilentlyContinue).ForestMode.ToString()
    }
} catch {}

# Recalculate summary counts based on critical findings
$CritCount = [int]$CriticalFindings.Count
if ($CritCount -gt 0) { $CritDCs = [Math]::Min($CritCount, [int]$DCCount) }
$HealthyDCs = [Math]::Max(0, [int]$DCCount - $CritDCs)

Write-Progress2 "Building HTML report..."

# ═══════════════════════════════════════════════════════════════════════════════
# COMPANY LOGO HTML
# ═══════════════════════════════════════════════════════════════════════════════
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

# ═══════════════════════════════════════════════════════════════════════════════
# HELPER: Build collapsible section
# ═══════════════════════════════════════════════════════════════════════════════
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

$EndTime     = Get-Date
$Duration    = ($EndTime - $StartTime).ToString('hh\:mm\:ss')
$ReportDate  = $EndTime.ToString('dddd, dd MMMM yyyy HH:mm:ss')

# ═══════════════════════════════════════════════════════════════════════════════
# BUILD FULL HTML
# ═══════════════════════════════════════════════════════════════════════════════
$HtmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AD Health Check Report  -  $(HtmlEncode $domainName)</title>
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
em { font-style: italic; }

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
.section-summary::before { content: '▶'; font-size: 10px; color: var(--muted);
                            transition: transform .2s; }
details[open] > .section-summary::before { transform: rotate(90deg); }
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

/* ── CODE BLOCK ── */
.code-block { background: var(--pre-bg); border: 1px solid var(--border); border-radius: 6px;
              padding: 12px; overflow-x: auto; font-size: 12px; font-family: 'SFMono-Regular',
              Consolas, monospace; white-space: pre; color: var(--text); line-height: 1.5; }
.repl-fail  { color: #f85149; font-weight: 600; }

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

<!-- ═══ HEADER ═══ -->
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

<!-- ═══ SUMMARY BAR ═══ -->
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

<!-- ═══ 14 SECTIONS ═══ -->
$(BuildSection 1 'Domain & Forest Info'            $Sec1Html  ($Sec1Html  -match 'error')   $true)
$(BuildSection 2 'Domain Controller Inventory'     $Sec2Html  ($Sec2Html  -match 'error')   $true)
$(BuildSection 3 'AD Services Status per DC'       $Sec3Html  ($Sec3Html  -match 'error')   $false)
$(BuildSection 4 'Replication Health'              $Sec4Html  ($Sec4Html  -match 'error')   $false)
$(BuildSection 5 'SYSVOL & Netlogon Shares'        $Sec5Html  ($Sec5Html  -match 'error')   $false)
$(BuildSection 6 'DNS Health'                      $Sec6Html  ($Sec6Html  -match 'error')   $false)
$(BuildSection 7 'FSMO Role Holders'               $Sec7Html  ($Sec7Html  -match 'error')   $false)
$(BuildSection 8 'AD Trust Relationships'          $Sec8Html  ($Sec8Html  -match 'error')   $false)
$(BuildSection 9 'AD Tombstone & Recycle Bin'      $Sec9Html  ($Sec9Html  -match 'error')   $false)
$(BuildSection 10 'Privileged Account Audit'       $Sec10Html ($Sec10Html -match 'error')   $false)
$(BuildSection 11 'Default Domain Password Policy' $Sec11Html ($Sec11Html -match 'error')   $false)
$(BuildSection 12 'Stale Objects'                  $Sec12Html ($Sec12Html -match 'error')   $false)
$(BuildSection 13 'Directory Service Event Log'    $Sec13Html ($Sec13Html -match 'error')   $false)
$(BuildSection 14 'Windows Update Status'          $Sec14Html ($Sec14Html -match 'error')   $false)

<!-- ═══ FOOTER ═══ -->
<div class="footer">
  <div>
    <strong>AD Health Check v$ScriptVersion</strong> &nbsp;|&nbsp;
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

# ═══════════════════════════════════════════════════════════════════════════════
# EMAIL ALERT (OPTIONAL)
# ═══════════════════════════════════════════════════════════════════════════════
if ($EnableEmailAlert -and $CriticalFindings.Count -gt 0) {
    Write-Progress2 "Sending critical alert email..."
    try {
        $emailBody = "Active Directory Health Check detected $($CriticalFindings.Count) critical finding(s):`r`n`r`n"
        $emailBody += ($CriticalFindings | ForEach-Object { "• $_" }) -join "`r`n"
        $emailBody += "`r`n`r`nPlease review the full report: $ReportFile"
        $emailBody += "`r`n`r`n-- AD Health Check v$ScriptVersion by $AuthorName"

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
Write-Host "  AD Health Check complete.  Duration: $Duration" -ForegroundColor DarkCyan
Write-Host "  Report: $ReportFile"                             -ForegroundColor Yellow
Write-Host "===============================================================" -ForegroundColor DarkCyan
Write-Host ""
