#Requires -Version 5.1
<#
.SYNOPSIS
    Comprehensive Active Directory Certificate Services (ADCS) Health Check Script

.DESCRIPTION
    Performs a thorough health check of the ADCS environment and exports the
    results to a single self-contained HTML file with a professional dark/light-themed
    dashboard UI.

.NOTES
    Version    : 1.0.0
    Author     : Tushar Gudde
    Requires   : PowerShell 5.1+, certutil.exe, RSAT ADCS Tools (optional)
    Permissions: Local Admin on CA server, or Domain Admin
#>

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
# Company Branding
$CompanyLogoURL    = ''                          # URL/path to company logo (PNG/SVG). Leave blank to skip.
$CompanyWebsite    = ''                          # Company website URL for logo hyperlink.

# Author
$AuthorName        = 'Tushar Gudde'              # Author name shown in footer

# Event Log Settings
$EventLogHours     = 24                          # How many hours back to scan ADCS events

# Certificate expiry warning thresholds
$CertWarnDays      = 60                          # Warn if CA cert expires within this many days
$CertCritDays      = 30                          # Critical if CA cert expires within this many days

# Email Alert Configuration
$EnableEmailAlert       = $false
$SMTPServer             = 'smtp.yourdomain.com'
$SMTPPort               = 587
$SMTPFrom               = 'adcs-healthcheck@yourdomain.com'
$SMTPTo                 = @('admin@yourdomain.com')
$SMTPSubject            = 'ADCS Health Check - CRITICAL ALERT'
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
$ReportFile = Join-Path $ReportsDir ("ADCS_Health_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

# ── DETECT ADCS INSTALLATION ─────────────────────────────────────────────────
$ADCSInstalled = $false
$CAName        = ''
$CAType        = ''

try {
    $certSvc = Get-Service -Name 'CertSvc' -ErrorAction Stop
    $ADCSInstalled = $true
    # Retrieve CA name from registry
    $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\CertSvc\Configuration'
    if (Test-Path $regPath) {
        $activeCa = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).Active
        if ($activeCa) {
            $CAName = $activeCa
            $caTypeVal = (Get-ItemProperty -Path "$regPath\$activeCa" -ErrorAction SilentlyContinue).CAType
            $CAType = switch ($caTypeVal) {
                0 { 'Enterprise Root CA' }
                1 { 'Enterprise Subordinate CA' }
                3 { 'Standalone Root CA' }
                4 { 'Standalone Subordinate CA' }
                default { "Unknown ($caTypeVal)" }
            }
        }
    }
} catch {
    $ADCSInstalled = $false
}

Write-Host ""
Write-Host "+==============================================================+" -ForegroundColor DarkCyan
Write-Host "|   Active Directory Certificate Services Health Check  v$ScriptVersion  |" -ForegroundColor DarkCyan
Write-Host "+==============================================================+" -ForegroundColor DarkCyan
Write-Host ""

if (-not $ADCSInstalled) {
    Write-Warning "Active Directory Certificate Services (CertSvc) is not installed or not accessible on this machine."
    Write-Warning "Some sections will be skipped. Run this script on the CA server for full results."
}

# =============================================================================
# SECTION 1 - CA SERVICE STATUS
# =============================================================================
Write-Progress2 "Section 1: CA Service Status..."
$Sec1Html  = ''
$Sec1Error = $false
try {
    $svc = Get-Service -Name 'CertSvc' -ErrorAction Stop
    $statusClr = switch ($svc.Status) {
        'Running' { 'green' }
        'Stopped' { 'red' }
        default    { 'yellow' }
    }
    if ($svc.Status -ne 'Running') {
        $CriticalFindings.Add("Section 1 - CertSvc service is $($svc.Status) - certificate authority may be unavailable")
        $Sec1Error = $true
    }

    $startType = $svc.StartType.ToString()
    $startClr  = if ($startType -eq 'Automatic') { 'green' } else { 'yellow' }

    $rows = @(
        @('Service Name',   (HtmlEncode $svc.Name))
        @('Display Name',   (HtmlEncode $svc.DisplayName))
        @('Status',         (StatusBadge $svc.Status.ToString() $statusClr))
        @('Start Type',     (StatusBadge $startType $startClr))
        @('CA Name',        (HtmlEncode $CAName))
        @('CA Type',        (HtmlEncode $CAType))
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
    $Sec1Html  = "<p class='error'>CertSvc not found - certificate authority is not installed or has been removed: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec1Error = $true
    $CriticalFindings.Add("Section 1 - CA Service not found: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 2 - CA CONFIGURATION
# =============================================================================
Write-Progress2 "Section 2: CA Configuration..."
$Sec2Html  = ''
$Sec2Error = $false
try {
    if (-not $ADCSInstalled) { throw "ADCS is not installed on this server." }

    $regBase = "HKLM:\SYSTEM\CurrentControlSet\Services\CertSvc\Configuration\$CAName"
    if (-not (Test-Path $regBase)) { throw "CA registry path not found: $regBase" }

    $caReg = Get-ItemProperty -Path $regBase -ErrorAction Stop

    # Audit enabled?
    $auditFilter = $caReg.AuditFilter
    $auditBadge  = if ($auditFilter -gt 0) { StatusBadge 'Enabled' 'green' } else { StatusBadge 'Disabled' 'red' }
    if ($auditFilter -eq 0) {
        $CriticalFindings.Add("Section 2 - CA auditing is disabled (AuditFilter=0). Enable CA auditing for compliance.")
        $Sec2Error = $true
    }

    # Database paths
    $dbPath  = HtmlEncode ($caReg.DBDirectory)
    $logPath = HtmlEncode ($caReg.LogDirectory)

    # CRL distribution points
    $cdpRaw  = try { certutil -getreg CA\CRLPublicationURLs 2>&1 | Out-String } catch { 'Unable to retrieve CDP' }
    $cdpEnc  = HtmlEncode (TruncateMessage $cdpRaw 1000)

    $rows = @(
        @('CA Name',        (HtmlEncode $CAName))
        @('CA Type',        (HtmlEncode $CAType))
        @('CA Auditing',    $auditBadge)
        @('DB Directory',   $dbPath)
        @('Log Directory',  $logPath)
    )

    $rowsHtml = ($rows | ForEach-Object {
        "<tr><td class='td-label'>$($_[0])</td><td>$($_[1])</td></tr>"
    }) -join ''

    $Sec2Html = @"
<table class='kv-table'>
  <tbody>$rowsHtml</tbody>
</table>
<h3>CRL Distribution Points (certutil -getreg)</h3>
<pre class='code-block'>$cdpEnc</pre>
"@
} catch {
    $Sec2Html  = "<p class='error'>Error retrieving CA configuration: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec2Error = $true
    $CriticalFindings.Add("Section 2 - CA configuration error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 3 - CA CERTIFICATE VALIDITY
# =============================================================================
Write-Progress2 "Section 3: CA Certificate Validity..."
$Sec3Html      = ''
$Sec3Error     = $false
$CACertDaysLeft = $null
try {
    if (-not $ADCSInstalled) {
        $Sec3Html = "<p class='warn'>ADCS not installed - certificate validity check skipped on this server.</p>"
    } else {
        # Get CA cert from LocalMachine store
        $caCerts = Get-ChildItem -Path 'Cert:\LocalMachine\CA' -ErrorAction SilentlyContinue |
                   Where-Object { $_.Subject -like "*$CAName*" }
        if (-not $caCerts -or $caCerts.Count -eq 0) {
            # Fallback: also check Root store
            $caCerts = Get-ChildItem -Path 'Cert:\LocalMachine\Root' -ErrorAction SilentlyContinue |
                       Where-Object { $_.Subject -like "*$CAName*" }
        }

        if ($caCerts -and $caCerts.Count -gt 0) {
            $caCert = $caCerts | Sort-Object NotAfter -Descending | Select-Object -First 1
            $CACertDaysLeft = [int](($caCert.NotAfter - (Get-Date)).TotalDays)

            $statusStr = if ($CACertDaysLeft -le $CertCritDays) {
                'Critical'
            } elseif ($CACertDaysLeft -le $CertWarnDays) {
                'Warning'
            } else {
                'Healthy'
            }
            $statusClr = switch ($statusStr) {
                'Critical' { 'red' }
                'Warning'  { 'yellow' }
                default     { 'green' }
            }

            if ($statusStr -eq 'Critical') {
                $CriticalFindings.Add("Section 3 - CA certificate expires in $CACertDaysLeft day(s) - CRITICAL renewal required")
                $Sec3Error = $true
            } elseif ($statusStr -eq 'Warning') {
                $CriticalFindings.Add("Section 3 - CA certificate expires in $CACertDaysLeft day(s) - renewal recommended")
            }

            $rows = @(
                @('Subject',         (HtmlEncode $caCert.Subject))
                @('Issuer',          (HtmlEncode $caCert.Issuer))
                @('Thumbprint',      (HtmlEncode $caCert.Thumbprint))
                @('Valid From',      (HtmlEncode $caCert.NotBefore.ToString('yyyy-MM-dd HH:mm:ss')))
                @('Valid To',        (HtmlEncode $caCert.NotAfter.ToString('yyyy-MM-dd HH:mm:ss')))
                @('Days Remaining',  "<span style='font-size:1.3rem;font-weight:700;'>$CACertDaysLeft days remaining &nbsp;$(StatusBadge $statusStr $statusClr)</span>")
                @('Key Algorithm',   (HtmlEncode $caCert.PublicKey.Oid.FriendlyName))
                @('Serial Number',   (HtmlEncode $caCert.SerialNumber))
            )

            $rowsHtml = ($rows | ForEach-Object {
                "<tr><td class='td-label'>$($_[0])</td><td>$($_[1])</td></tr>"
            }) -join ''

            $Sec3Html = @"
<table class='kv-table'>
  <tbody>$rowsHtml</tbody>
</table>
"@
        } else {
            $Sec3Html  = "<p class='warn'>Could not determine CA certificate details. No certificate found matching CA name '$(HtmlEncode $CAName)' in LocalMachine\CA or LocalMachine\Root stores.</p>"
            $Sec3Error = $true
            $CriticalFindings.Add("Section 3 - CA certificate not found in certificate stores")
        }
    }
} catch {
    $Sec3Html  = "<p class='error'>Error checking CA certificate validity: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec3Error = $true
    $CriticalFindings.Add("Section 3 - CA certificate validity error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 4 - CRL HEALTH
# =============================================================================
Write-Progress2 "Section 4: CRL Health..."
$Sec4Html  = ''
$Sec4Error = $false
try {
    if (-not $ADCSInstalled) {
        $Sec4Html = "<p class='warn'>ADCS not installed - CRL health check skipped on this server.</p>"
    } else {
        # Get CRL info via certutil
        $crlRaw  = certutil -CRL 2>&1 | Out-String
        $crlEnc  = HtmlEncode (TruncateMessage $crlRaw 2000)

        # Check for CRL expiry using certutil -store CA
        $crlStore = certutil -store CA 2>&1 | Out-String
        $crlExpired = $crlStore -match 'expire|expired|invalid'
        if ($crlExpired) {
            $CriticalFindings.Add("Section 4 - CRL may be expired or invalid. Review certutil -store CA output.")
            $Sec4Error = $true
        }

        # Check CDP accessibility
        $cdpUrls = @()
        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\CertSvc\Configuration\$CAName"
        if (Test-Path $regPath) {
            $cdpVal = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).CRLPublicationURLs
            if ($cdpVal) {
                # Extract http(s) and ldap URLs
                $cdpUrls = $cdpVal -split "`n" | Where-Object { $_ -match 'https?://|ldap://' } | ForEach-Object { $_.Trim().TrimStart('0123456789:') }
            }
        }

        $cdpRows = ''
        foreach ($url in $cdpUrls) {
            if ([string]::IsNullOrWhiteSpace($url)) { continue }
            $urlEnc = HtmlEncode $url
            try {
                $req = [System.Net.HttpWebRequest]::Create($url)
                $req.Timeout = 5000
                $resp = $req.GetResponse()
                $resp.Close()
                $cdpRows += "<tr><td>$urlEnc</td><td>$(StatusBadge 'Accessible' 'green')</td></tr>"
            } catch {
                $cdpRows += "<tr><td>$urlEnc</td><td>$(StatusBadge 'Unreachable' 'red')</td></tr>"
                $CriticalFindings.Add("Section 4 - CDP URL unreachable: $url")
                $Sec4Error = $true
            }
        }

        $cdpTableHtml = ''
        if ($cdpRows) {
            $cdpTableHtml = @"
<h3>CDP URL Accessibility</h3>
<div class='table-wrap'>
<table>
  <thead><tr><th>URL</th><th>Status</th></tr></thead>
  <tbody>$cdpRows</tbody>
</table>
</div>
"@
        }

        $Sec4Html = @"
<h3>certutil -CRL Output</h3>
<pre class='code-block'>$crlEnc</pre>
$cdpTableHtml
"@
    }
} catch {
    $Sec4Html  = "<p class='error'>Error checking CRL health: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec4Error = $true
    $CriticalFindings.Add("Section 4 - CRL health error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 5 - MACHINE STORE CERTIFICATES (with private key)
# =============================================================================
Write-Progress2 "Section 5: Machine Store Certificates (with private key)..."
$Sec5Html  = ''
$Sec5Error = $false
try {
    $allCerts = Get-ChildItem -Path 'Cert:\LocalMachine\My' -ErrorAction Stop |
                Where-Object { $_.HasPrivateKey }

    if ($allCerts.Count -eq 0) {
        $Sec5Html = "<p class='info'>No certificates with private keys found in LocalMachine\My store.</p>"
    } else {
        $rows = foreach ($cert in $allCerts | Sort-Object NotAfter) {
            $daysLeft = [int](($cert.NotAfter - (Get-Date)).TotalDays)
            $expClr   = if ($daysLeft -le $CertCritDays) { 'red' } elseif ($daysLeft -le $CertWarnDays) { 'yellow' } else { 'green' }

            if ($daysLeft -le $CertCritDays) {
                $CriticalFindings.Add("Section 5 - Machine cert expiring in $daysLeft day(s): $($cert.Subject)")
                $Sec5Error = $true
            }

            $subject    = HtmlEncode (TruncateMessage $cert.Subject 80)
            $issuer     = HtmlEncode (TruncateMessage $cert.Issuer 60)
            $thumb      = HtmlEncode $cert.Thumbprint
            $notAfter   = HtmlEncode $cert.NotAfter.ToString('yyyy-MM-dd')
            $daysHtml   = StatusBadge "$daysLeft d" $expClr
            $pkProvider = try {
                if ($cert.PrivateKey -and $cert.PrivateKey.CspKeyContainerInfo) {
                    HtmlEncode $cert.PrivateKey.CspKeyContainerInfo.ProviderName
                } else {
                    # CNG key provider (KSP) does not expose CspKeyContainerInfo
                    HtmlEncode 'CNG/KSP'
                }
            } catch { HtmlEncode 'Unknown' }

            "<tr><td>$subject</td><td>$issuer</td><td>$thumb</td><td>$notAfter</td><td>$daysHtml</td><td>$pkProvider</td></tr>"
        }

        $Sec5Html = @"
<p style='margin:14px 0 8px;'>Machine Store Certificates (with private key) &nbsp;$(StatusBadge "$($allCerts.Count) certs" 'blue')</p>
<div class='table-wrap'>
<table>
  <thead><tr>
    <th>Subject</th><th>Issuer</th><th>Thumbprint</th>
    <th>Expires</th><th>Days Left</th><th>Key Provider</th>
  </tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
    }
} catch {
    $Sec5Html  = "<p class='error'>Error checking machine store certificates: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec5Error = $true
    $CriticalFindings.Add("Section 5 - Machine store certificate error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 6 - CERTIFICATE TEMPLATES
# =============================================================================
Write-Progress2 "Section 6: Certificate Templates..."
$Sec6Html  = ''
$Sec6Error = $false
try {
    if (-not $ADCSInstalled) {
        $Sec6Html = "<p class='warn'>ADCS not installed - certificate template check skipped.</p>"
    } else {
        $templatesRaw = certutil -CATemplates 2>&1 | Out-String
        $templatesEnc = HtmlEncode (TruncateMessage $templatesRaw 3000)

        if ($templatesRaw -match 'error|failed') {
            $CriticalFindings.Add("Section 6 - Error retrieving certificate templates")
            $Sec6Error = $true
        }

        $Sec6Html = @"
<h3>Available Certificate Templates</h3>
<pre class='code-block'>$templatesEnc</pre>
"@
    }
} catch {
    $Sec6Html  = "<p class='error'>Error retrieving certificate templates: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec6Error = $true
    $CriticalFindings.Add("Section 6 - Certificate templates error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 7 - PENDING AND FAILED REQUESTS
# =============================================================================
Write-Progress2 "Section 7: Pending and Failed Requests..."
$Sec7Html  = ''
$Sec7Error = $false
try {
    if (-not $ADCSInstalled) {
        $Sec7Html = "<p class='warn'>ADCS not installed - pending/failed request check skipped.</p>"
    } else {
        # Pending requests (Disposition=9 = pending administrator approval)
        $pendingRaw = certutil -view -restrict "Disposition=9" -out "CommonName,NotAfter,RequesterName,StatusCode" 2>&1 | Out-String
        $pendingCount = ([regex]::Matches($pendingRaw, 'Row\s+\d+')).Count

        # Denied/failed requests (Disposition=30 = denied; last 24 h)
        $failedRaw  = certutil -view -restrict "Disposition=30,Request.SubmittedWhen>=$((Get-Date).AddDays(-1).ToString('M/d/yyyy'))" -out "CommonName,NotAfter,RequesterName,StatusCode" 2>&1 | Out-String
        $failedCount = ([regex]::Matches($failedRaw, 'Row\s+\d+')).Count

        if ($pendingCount -gt 0) {
            $CriticalFindings.Add("Section 7 - $pendingCount pending certificate request(s) require administrator action")
        }
        if ($failedCount -gt 0) {
            $CriticalFindings.Add("Section 7 - $failedCount failed/denied certificate request(s) in the last 24 hours")
            $Sec7Error = $true
        }

        $pendingEnc = HtmlEncode (TruncateMessage $pendingRaw 1500)
        $failedEnc  = HtmlEncode (TruncateMessage $failedRaw  1500)

        $Sec7Html = @"
<h3>Pending Requests &nbsp;$(StatusBadge "$pendingCount pending" $(if ($pendingCount -gt 0) { 'yellow' } else { 'green' }))</h3>
<pre class='code-block'>$pendingEnc</pre>
<h3>Failed / Denied Requests (last 24 h) &nbsp;$(StatusBadge "$failedCount failed" $(if ($failedCount -gt 0) { 'red' } else { 'green' }))</h3>
<pre class='code-block'>$failedEnc</pre>
"@
    }
} catch {
    $Sec7Html  = "<p class='error'>Error retrieving pending/failed requests: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec7Error = $true
    $CriticalFindings.Add("Section 7 - Pending/failed requests error: $($_.Exception.Message)")
}

# =============================================================================
# SECTION 8 - ADCS EVENT LOG
# =============================================================================
Write-Progress2 "Section 8: ADCS Event Log..."
$Sec8Html  = ''
$Sec8Error = $false

$adcsEventMap = @{
    4870 = @{ Impact = 'Warning';  Suggestion = 'CA certificate revoked.' }
    4879 = @{ Impact = 'Warning';  Suggestion = 'CA certificate renewed.' }
    4880 = @{ Impact = 'Info';     Suggestion = 'Certificate Services started.' }
    4881 = @{ Impact = 'Critical'; Suggestion = 'Certificate Services stopped - investigate immediately.' }
    4882 = @{ Impact = 'Warning';  Suggestion = 'Security permissions changed on Certificate Services.' }
    4883 = @{ Impact = 'Warning';  Suggestion = 'Archived key retrieved.' }
    4884 = @{ Impact = 'Warning';  Suggestion = 'Certificate approved.' }
    4885 = @{ Impact = 'Warning';  Suggestion = 'Certificate audit filter changed.' }
    4886 = @{ Impact = 'Info';     Suggestion = 'Certificate requested.' }
    4887 = @{ Impact = 'Info';     Suggestion = 'Certificate issued.' }
    4888 = @{ Impact = 'Warning';  Suggestion = 'Certificate request denied.' }
    4889 = @{ Impact = 'Warning';  Suggestion = 'Certificate request set to pending.' }
    4890 = @{ Impact = 'Warning';  Suggestion = 'Certificate manager settings changed.' }
    4896 = @{ Impact = 'Critical'; Suggestion = 'CA database row deleted - investigate for tampering.' }
    4897 = @{ Impact = 'Critical'; Suggestion = 'Role separation enabled - review for compliance.' }
    4898 = @{ Impact = 'Warning';  Suggestion = 'Certificate template loaded.' }
}

try {
    $since   = (Get-Date).AddHours(-$EventLogHours)
    $events  = Get-WinEvent -LogName 'Security' -ErrorAction SilentlyContinue |
               Where-Object {
                   $_.TimeCreated -ge $since -and
                   $adcsEventMap.ContainsKey($_.Id)
               } | Select-Object -First 200

    if ($null -eq $events -or @($events).Count -eq 0) {
        $Sec8Html = "<p class='info'>No ADCS-related security events found in the last $EventLogHours hour(s).</p>"
    } else {
        $rows = foreach ($ev in $events | Sort-Object TimeCreated -Descending) {
            $meta    = $adcsEventMap[$ev.Id]
            $impact  = $meta.Impact
            $impClr  = switch ($impact) {
                'Critical' { 'red' }
                'Warning'  { 'yellow' }
                default     { 'blue' }
            }
            if ($impact -eq 'Critical') {
                $CriticalFindings.Add("Section 8 - CRITICAL event $($ev.Id): $($meta.Suggestion)")
                $Sec8Error = $true
            }
            $rowClass = switch ($impact) {
                'Critical' { 'row-critical' }
                'Warning'  { 'row-warning' }
                default     { '' }
            }
            $timeStr  = HtmlEncode $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            $msgShort = HtmlEncode (TruncateMessage $ev.Message 120)
            $suggest  = HtmlEncode $meta.Suggestion
            "<tr class='$rowClass'><td>$timeStr</td><td>$(HtmlEncode $ev.Id.ToString())</td><td>$(StatusBadge $impact $impClr)</td><td>$msgShort</td><td>$suggest</td></tr>"
        }

        $Sec8Html = @"
<div class='table-wrap'>
<table>
  <thead><tr>
    <th>Time</th><th>Event ID</th><th>Impact</th><th>Message</th><th>Suggestion</th>
  </tr></thead>
  <tbody>$($rows -join '')</tbody>
</table>
</div>
"@
    }
} catch {
    $Sec8Html  = "<p class='error'>Error reading ADCS event log: $(HtmlEncode $_.Exception.Message)</p>"
    $Sec8Error = $true
    $CriticalFindings.Add("Section 8 - Event log error: $($_.Exception.Message)")
}

# =============================================================================
# SUMMARY COUNTS
# =============================================================================
$CritCount   = $CriticalFindings.Count
$TotalCerts  = 0
try {
    $TotalCerts = @(Get-ChildItem -Path 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue).Count
} catch {}

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
<title>ADCS Health Check Report - $(HtmlEncode $CAName)</title>
<style>
/* CSS VARIABLES */
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
.page-wrap  { max-width: 1400px; margin: 0 auto; padding: 0 16px 40px; }
.header     { background: var(--head-bg); border-bottom: 1px solid var(--border);
              padding: 16px 24px; display: flex; align-items: center;
              justify-content: space-between; flex-wrap: wrap; gap: 12px; }
.header-left  { display: flex; align-items: center; gap: 16px; }
.logo-wrap img { max-height: 52px; }
.header-title h1 { font-size: 1.4rem; color: var(--text); }
.header-title p  { font-size: .8rem; color: var(--muted); margin:0; }
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
.kv-table   { width: 100%; border-collapse: collapse; font-size: 13px; }
.kv-table td { padding: 7px 12px; border-bottom: 1px solid var(--border); }
.td-label   { font-weight: 600; white-space: nowrap; width: 220px; color: var(--muted); }
.row-critical { background: rgba(218,54,51,.15) !important; }
.row-error    { background: rgba(218,54,51,.08) !important; }
.row-warning  { background: rgba(210,153,34,.12) !important; }
.badge { display: inline-block; font-size: .7rem; font-weight: 600; padding: 2px 8px;
         border-radius: 20px; color: #fff; white-space: nowrap; }
.code-block { background: var(--pre-bg); border: 1px solid var(--border); border-radius: 6px;
              padding: 12px; overflow-x: auto; font-size: 12px; font-family: 'SFMono-Regular',
              Consolas, monospace; white-space: pre; color: var(--text); line-height: 1.5; }
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
      <h1>&#x1F512; ADCS Health Check</h1>
      <p>$(HtmlEncode $CAName) &nbsp;|&nbsp; Generated: $(HtmlEncode $ReportDate)</p>
    </div>
  </div>
  <button class="theme-toggle" onclick="toggleTheme()" title="Toggle Dark/Light mode">
    <span id="theme-icon">&#x2600;&#xFE0F;</span> Toggle Theme
  </button>
</div>

<div class="page-wrap">
<div class="summary-bar">
  <div class="stat-card">
    <div class="stat-count c-blue">$TotalCerts</div>
    <div class="stat-label">Machine Certs</div>
  </div>
  <div class="stat-card">
    <div class="stat-count $(if ($CACertDaysLeft -ne $null -and $CACertDaysLeft -le $CertCritDays) { 'c-red' } elseif ($CACertDaysLeft -ne $null -and $CACertDaysLeft -le $CertWarnDays) { 'c-yellow' } else { 'c-green' })">$(if ($CACertDaysLeft -ne $null) { $CACertDaysLeft } else { 'N/A' })</div>
    <div class="stat-label">CA Cert Days Left</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-red">$CritCount</div>
    <div class="stat-label">Critical Findings</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-muted" style="font-size:1rem; padding-top:8px;">$(HtmlEncode $CAName)</div>
    <div class="stat-label">CA Name</div>
  </div>
  <div class="stat-card">
    <div class="stat-count c-muted" style="font-size:1rem; padding-top:8px;">$(HtmlEncode $CAType)</div>
    <div class="stat-label">CA Type</div>
  </div>
</div>

$(BuildSection 1 'CA Service Status'                     $Sec1Html $Sec1Error $true)
$(BuildSection 2 'CA Configuration'                      $Sec2Html $Sec2Error $false)
$(BuildSection 3 'CA Certificate Validity'               $Sec3Html $Sec3Error $true)
$(BuildSection 4 'CRL Health'                            $Sec4Html $Sec4Error $false)
$(BuildSection 5 'Machine Store Certificates'            $Sec5Html $Sec5Error $false)
$(BuildSection 6 'Certificate Templates'                 $Sec6Html $Sec6Error $false)
$(BuildSection 7 'Pending and Failed Requests'           $Sec7Html $Sec7Error $false)
$(BuildSection 8 'ADCS Event Log'                        $Sec8Html $Sec8Error $false)

<div class="footer">
  <div>
    <strong>ADCS Health Check v$ScriptVersion</strong> &nbsp;|&nbsp;
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
        $emailBody  = "ADCS Health Check detected $($CriticalFindings.Count) critical finding(s):`r`n`r`n"
        $emailBody += ($CriticalFindings | ForEach-Object { "* $_" }) -join "`r`n"
        $emailBody += "`r`n`r`nPlease review the full report: $ReportFile"
        $emailBody += "`r`n`r`n-- ADCS Health Check v$ScriptVersion by $AuthorName"

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
Write-Host "  ADCS Health Check complete.  Duration: $Duration" -ForegroundColor DarkCyan
Write-Host "  Report: $ReportFile"                               -ForegroundColor Yellow
Write-Host "===============================================================" -ForegroundColor DarkCyan
Write-Host ""
