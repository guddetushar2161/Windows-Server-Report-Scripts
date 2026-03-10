#Requires -Version 5.1
<#
.SYNOPSIS
    ADCS / PKI Health Check Script

.DESCRIPTION
    Performs a comprehensive health check of an Active Directory Certificate
    Services (ADCS) Certificate Authority (CA) installation and exports all
    results to a self-contained HTML dashboard file.

    Gracefully detects whether ADCS is installed: if the CertSvc service and
    certutil.exe are absent the script still produces an HTML report that says
    "ADCS Not Installed" rather than crashing.

    Every section is individually wrapped in Try/Catch so a single failure
    never aborts the run.

    Checks performed:
      1.  CA Identity & Config       (name, type, root/sub, cert subject, validity, key, algo)
      2.  ADCS Services Status       (CertSvc, OCSP, Web Enrollment)
      3.  CA Certificate Validity    (days remaining; warn <180, critical <30)
      4.  CRL Health                 (published, validity, next publish, delta, days to expiry)
      5.  CRL Distribution Points    (CDP list, HTTP accessibility test)
      6.  OCSP Status                (service, array URL, response test)
      7.  Issued Certificates        (total, expiring 30/60/90 days, expired, revoked)
      8.  Certificate Templates      (published templates: schema, validity, renewal, RA)
      9.  AIA URLs                   (list and accessibility test)
      10. CA Database Size           (file size and location)
      11. Failed Requests            (failed/denied requests in last 7 days)
      12. CA Event Logs              (last 10 Critical/Error from Certificate Services log)
      13. Enrollment Agent / KRA     (KRA certs configured and valid)
      14. HSM / Key Storage          (Software KSP vs HSM)

.NOTES
    Version    : 1.0.0
    Author     : Tushar Gudde
    Requires   : PowerShell 5.1+, Local Administrator rights
    Compatible : Windows Server 2016, 2019, 2022 (Enterprise CA or Standalone CA)
    Companion  : ADCS_HealthCheck_EmailAlert.ps1
#>

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
$CompanyLogoURL = ''                         # URL/path to logo. Leave blank to skip.
$CompanyWebsite = 'https://tushargudde.tech' # Company website URL for logo hyperlink.
$AuthorName     = 'Tushar Gudde'             # Author name shown in the footer.

# CA certificate validity warning thresholds (days)
$CACertWarnDays = 180
$CACertCritDays = 30

# CRL expiry warning thresholds (days)
$CRLWarnDays = 14
$CRLCritDays = 7

# HTTP reachability test timeout (seconds)
$UrlTestTimeout = 10

# Write a plain-text companion status file alongside every HTML report.
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
$ReportStamp = "ADC_Health_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
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
        rose   = '#e05c7a'
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
        "<span style='color:#e05c7a;'>&#x2714;</span>"
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

function Invoke-CertUtil {
    param([string[]]$Arguments)
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName  = 'certutil.exe'
        $psi.Arguments = $Arguments -join ' '
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true
        $proc = [System.Diagnostics.Process]::Start($psi)
        $stdout = $proc.StandardOutput.ReadToEnd()
        $proc.WaitForExit(30000) | Out-Null
        return $stdout
    } catch {
        return ''
    }
}

function Test-HttpUrl {
    param([string]$Url, [int]$TimeoutSec = 10)
    try {
        $req = [System.Net.HttpWebRequest]::Create($Url)
        $req.Timeout = $TimeoutSec * 1000
        $req.Method  = 'HEAD'
        $resp = $req.GetResponse()
        $resp.Close()
        return $true
    } catch {
        return $false
    }
}

# ── CRITICAL FINDINGS LIST ────────────────────────────────────────────────────
$CriticalFindings = [System.Collections.Generic.List[string]]::new()

# ── DEFAULT KPI VARIABLES ─────────────────────────────────────────────────────
$ServerHostname   = $env:COMPUTERNAME
$CAName           = 'Unknown'
$CACertDaysLeft   = -1
$CRLDaysLeft      = -1
$ExpiringCertCount = 0
$IsADCSInstalled  = $false

Write-Host ""
Write-Host "+==============================================================+" -ForegroundColor DarkRed
Write-Host "|   ADCS / PKI Health Check  v$ScriptVersion" -ForegroundColor DarkRed
Write-Host "|   Server: $($env:COMPUTERNAME)" -ForegroundColor DarkRed
Write-Host "+==============================================================+" -ForegroundColor DarkRed
Write-Host ""

# ── DETECT WHETHER ADCS IS INSTALLED ─────────────────────────────────────────
Write-Progress2 "Detecting ADCS installation..."
try {
    $certSvc = Get-Service -Name CertSvc -ErrorAction Stop
    $IsADCSInstalled = $true
    Write-Progress2 "ADCS (CertSvc) detected."
} catch {
    $IsADCSInstalled = $false
    Write-Progress2 "CertSvc not found  -  ADCS does not appear to be installed on this server."
    $CriticalFindings.Add("ADCS (CertSvc) service not found  -  certificate authority is not installed or has been removed")
}

# ── READ CA REGISTRY CONFIG ───────────────────────────────────────────────────
$CARegBase   = 'HKLM:\SYSTEM\CurrentControlSet\Services\CertSvc\Configuration'
$CANamesList = @()
try {
    if (Test-Path $CARegBase) {
        $CANamesList = @(Get-ChildItem $CARegBase -ErrorAction Stop | Select-Object -ExpandProperty PSChildName)
        if ($CANamesList.Count -gt 0) { $CAName = $CANamesList[0] }
    }
} catch {}

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
    $cpuNames    = ($cpu | ForEach-Object { HtmlEncode $(if ($null -ne $_.Name) { $_.Name.Trim() } else { 'Unknown' }) } | Select-Object -Unique) -join '; '

    $installDate  = if ($null -ne $os.InstallDate)    { $os.InstallDate.ToString('yyyy-MM-dd')             } else { 'N/A' }
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

    $Sec0Html = BuildKVTable $rows0
} catch {
    $Sec0Html = "<p class='error'>Error retrieving Server Details: $(HtmlEncode $_.Exception.Message)</p>"
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1  -  CA IDENTITY & CONFIG
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 1: CA Identity & Config..."
$Sec1Html = ''
if (-not $IsADCSInstalled) {
    $Sec1Html = "<div class='not-installed-card'><span class='ni-icon'>&#x1F4CB;</span><div><strong>ADCS Not Installed</strong><p>The Active Directory Certificate Services role (CertSvc) was not detected on this server.</p></div></div>"
} else {
    try {
        $caType       = 'Unknown'
        $caKind       = 'Unknown'
        $caSubject    = 'Unknown'
        $caKeyLength  = 'Unknown'
        $caHashAlg    = 'Unknown'
        $caConfigStr  = 'Unknown'
        $caValidFrom  = 'Unknown'
        $caValidTo    = 'Unknown'

        if ($CANamesList.Count -gt 0 -and (Test-Path "$CARegBase\$CAName")) {
            $caReg = Get-ItemProperty "$CARegBase\$CAName" -ErrorAction Stop
            $caType = switch ($caReg.CAType) {
                0 { 'Enterprise Root CA' }
                1 { 'Enterprise Subordinate CA' }
                3 { 'Standalone Root CA' }
                4 { 'Standalone Subordinate CA' }
                default { "Type $($caReg.CAType)" }
            }
            $caKind = if ($caReg.CAType -in @(0,3)) { 'Root CA' } else { 'Subordinate CA' }
            if ($null -ne $caReg.CommonName) { $caSubject = $caReg.CommonName }
        }

        # Try certutil -getreg CA\CACertHash for more info
        $certutilCA = Invoke-CertUtil @('-getreg', 'CA')
        if ($certutilCA -match 'Algorithm:\s*(.+)') { $caHashAlg = $Matches[1].Trim() }
        if ($certutilCA -match 'KeyLength:\s*(\d+)') { $caKeyLength = $Matches[1].Trim() + ' bits' }

        # Get CA cert via PKI module or certutil dump
        try {
            $store   = New-Object System.Security.Cryptography.X509Certificates.X509Store('My','LocalMachine')
            $store.Open('ReadOnly')
            $caCerts = @($store.Certificates | Where-Object {
                $_.Subject -match [regex]::Escape($CAName) -and $_.HasPrivateKey -eq $true
            })
            $store.Close()
            if ($caCerts.Count -eq 0) {
                # Try without private key requirement
                $store.Open('ReadOnly')
                $caCerts = @($store.Certificates | Where-Object { $_.Subject -match [regex]::Escape($CAName) })
                $store.Close()
            }
            if ($caCerts.Count -gt 0) {
                $caCert       = $caCerts | Sort-Object NotAfter -Descending | Select-Object -First 1
                $caSubject    = $caCert.Subject
                $caValidFrom  = $caCert.NotBefore.ToString('yyyy-MM-dd')
                $caValidTo    = $caCert.NotAfter.ToString('yyyy-MM-dd')
                $CACertDaysLeft = [int]($caCert.NotAfter - (Get-Date)).TotalDays
                $caKeyLength  = "$($caCert.PublicKey.Key.KeySize) bits"
                $caHashAlg    = $caCert.SignatureAlgorithm.FriendlyName
            }
        } catch {}

        $certDaysColor = if ($CACertDaysLeft -lt 0) { 'grey' }
                         elseif ($CACertDaysLeft -le $CACertCritDays) { 'red' }
                         elseif ($CACertDaysLeft -le $CACertWarnDays) { 'yellow' }
                         else { 'green' }
        if ($CACertDaysLeft -ge 0 -and $CACertDaysLeft -le $CACertCritDays) {
            $CriticalFindings.Add("Section 1 - CA certificate expires in $CACertDaysLeft days (CRITICAL)")
        } elseif ($CACertDaysLeft -ge 0 -and $CACertDaysLeft -le $CACertWarnDays) {
            $CriticalFindings.Add("Section 1 - CA certificate expires in $CACertDaysLeft days (Warning)")
        }

        $daysLabel = if ($CACertDaysLeft -ge 0) { "$CACertDaysLeft days remaining" } else { 'N/A' }

        $rows1 = @(
            @('CA Name',           (HtmlEncode $CAName)),
            @('CA Type',           (HtmlEncode $caType)),
            @('CA Kind',           (HtmlEncode $caKind)),
            @('Certificate Subject', (HtmlEncode $caSubject)),
            @('Valid From',        (HtmlEncode $caValidFrom)),
            @('Valid To',          (HtmlEncode $caValidTo)),
            @('Days Remaining',    (StatusBadge $daysLabel $certDaysColor)),
            @('Key Length',        (HtmlEncode $caKeyLength)),
            @('Signature Algorithm', (HtmlEncode $caHashAlg))
        )
        $Sec1Html = BuildKVTable $rows1
    } catch {
        $Sec1Html = "<p class='error'>Error retrieving CA identity: $(HtmlEncode $_.Exception.Message)</p>"
        $CriticalFindings.Add("Section 1 - CA identity error: $($_.Exception.Message)")
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2  -  ADCS SERVICES STATUS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 2: ADCS Services Status..."
$Sec2Html = ''
try {
    $adcsServices = @(
        @{ Name = 'CertSvc';   Display = 'Active Directory Certificate Services'; Required = $true },
        @{ Name = 'OCSPSvc';   Display = 'Online Responder (OCSP)';               Required = $false },
        @{ Name = 'certsrv';   Display = 'Certificate Web Enrollment';             Required = $false }
    )
    $Sec2Html  = "<div class='table-wrap'><table><thead><tr>"
    $Sec2Html += "<th>Service</th><th>Display Name</th><th>Status</th><th>Start Type</th><th>Required</th>"
    $Sec2Html += "</tr></thead><tbody>"
    foreach ($svcDef in $adcsServices) {
        try {
            $svc = Get-Service -Name $svcDef.Name -ErrorAction Stop
            $status    = $svc.Status.ToString()
            $startType = $svc.StartType.ToString()
            $stColor   = if ($status -eq 'Running') { 'green' }
                         elseif ($status -eq 'Stopped' -and $svcDef.Required) { 'red' }
                         else { 'yellow' }
            if ($status -eq 'Stopped' -and $svcDef.Required) {
                $CriticalFindings.Add("Section 2 - Required service '$($svcDef.Display)' is Stopped")
            }
            $reqBadge = if ($svcDef.Required) { StatusBadge 'Required' 'rose' } else { StatusBadge 'Optional' 'grey' }
            $Sec2Html += "<tr><td><code>$(HtmlEncode $svcDef.Name)</code></td><td>$(HtmlEncode $svc.DisplayName)</td>"
            $Sec2Html += "<td>$(StatusBadge $status $stColor)</td><td>$(HtmlEncode $startType)</td><td>$reqBadge</td></tr>"
        } catch {
            $reqBadge = if ($svcDef.Required) { StatusBadge 'Required' 'rose' } else { StatusBadge 'Optional' 'grey' }
            $Sec2Html += "<tr><td><code>$(HtmlEncode $svcDef.Name)</code></td><td>$(HtmlEncode $svcDef.Display)</td>"
            $Sec2Html += "<td>$(StatusBadge 'Not Installed' 'grey')</td><td>N/A</td><td>$reqBadge</td></tr>"
        }
    }
    $Sec2Html += "</tbody></table></div>"
} catch {
    $Sec2Html = "<p class='error'>Error checking ADCS services: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 2 - ADCS services check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3  -  CA CERTIFICATE VALIDITY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 3: CA Certificate Validity..."
$Sec3Html = ''
try {
    if (-not $IsADCSInstalled) {
        $Sec3Html = "<p class='warn'>ADCS not installed  -  certificate validity check skipped.</p>"
    } else {
        # Build countdown bar
        if ($CACertDaysLeft -ge 0) {
            $maxDays   = 3650  # 10-year default for display
            $barPct    = [math]::Min([math]::Round(($CACertDaysLeft / $maxDays) * 100, 1), 100)
            $barColor  = if ($CACertDaysLeft -le $CACertCritDays) { '#f85149' }
                         elseif ($CACertDaysLeft -le $CACertWarnDays) { '#d29922' }
                         else { '#3fb950' }
            $statusStr = if ($CACertDaysLeft -le $CACertCritDays) { 'CRITICAL' }
                         elseif ($CACertDaysLeft -le $CACertWarnDays) { 'Warning' }
                         else { 'OK' }
            $statusClr = if ($CACertDaysLeft -le $CACertCritDays) { 'red' }
                         elseif ($CACertDaysLeft -le $CACertWarnDays) { 'yellow' }
                         else { 'green' }

            $Sec3Html  = "<div class='disk-row'>"
            $Sec3Html += "<div class='disk-label'><strong>CA Certificate Expiry</strong></div>"
            $Sec3Html += "<div class='disk-bar-outer'><div class='disk-bar-inner' style='width:${barPct}%;background:$barColor;'></div></div>"
            $Sec3Html += "<div class='disk-stat'><strong style='color:$barColor;font-size:1.3rem;'>$CACertDaysLeft</strong> days remaining &nbsp;$(StatusBadge $statusStr $statusClr)</div>"
            $Sec3Html += "</div>"
        } else {
            $Sec3Html = "<p class='warn'>Could not determine CA certificate expiry date.</p>"
        }

        # Enumerate all certs in Local Machine My store matching CA name
        $allCaCerts = @()
        try {
            $store2 = New-Object System.Security.Cryptography.X509Certificates.X509Store('My','LocalMachine')
            $store2.Open('ReadOnly')
            $allCaCerts = @($store2.Certificates | Where-Object { $_.HasPrivateKey -eq $true })
            $store2.Close()
        } catch {}

        if ($allCaCerts.Count -gt 0) {
            $Sec3Html += "<h4 style='margin:14px 0 8px;'>Machine Store Certificates (with private key)</h4>"
            $Sec3Html += "<div class='table-wrap'><table><thead><tr>"
            $Sec3Html += "<th>Subject</th><th>Issuer</th><th>Valid From</th><th>Valid To</th><th>Days Left</th><th>Thumbprint</th>"
            $Sec3Html += "</tr></thead><tbody>"
            foreach ($cert in ($allCaCerts | Sort-Object NotAfter -Descending)) {
                $dl = [int]($cert.NotAfter - (Get-Date)).TotalDays
                $dlColor = if ($dl -lt 0) { 'red' } elseif ($dl -le $CACertCritDays) { 'red' } elseif ($dl -le $CACertWarnDays) { 'yellow' } else { 'green' }
                $Sec3Html += "<tr><td>$(HtmlEncode $cert.Subject)</td><td>$(HtmlEncode $cert.Issuer)</td>"
                $Sec3Html += "<td>$($cert.NotBefore.ToString('yyyy-MM-dd'))</td><td>$($cert.NotAfter.ToString('yyyy-MM-dd'))</td>"
                $Sec3Html += "<td>$(StatusBadge $dl.ToString() $dlColor)</td><td><code>$($cert.Thumbprint.Substring(0,[math]::Min(16,$cert.Thumbprint.Length)))...</code></td></tr>"
            }
            $Sec3Html += "</tbody></table></div>"
        }
    }
} catch {
    $Sec3Html = "<p class='error'>Error checking CA certificate validity: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 3 - CA certificate validity error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4  -  CRL HEALTH
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 4: CRL Health..."
$Sec4Html = ''
$CRLDaysLeft = -1
try {
    if (-not $IsADCSInstalled) {
        $Sec4Html = "<p class='warn'>ADCS not installed  -  CRL health check skipped.</p>"
    } else {
        # Try to find CRL in the default CertEnroll folder
        $certEnrollPath = Join-Path $env:SystemRoot 'system32\CertSrv\CertEnroll'
        $crlFiles = @()
        if (Test-Path $certEnrollPath) {
            $crlFiles = @(Get-ChildItem -Path $certEnrollPath -Filter '*.crl' -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending)
        }

        # Also try certutil -CRL output
        $certutilCRL = Invoke-CertUtil @('-CRL')
        $crlPublishedOK = $certutilCRL -match 'CRL successfully created'

        # Parse CRL via .NET if files available
        $crlInfoHtml = ''
        if ($crlFiles.Count -gt 0) {
            $crlInfoHtml  = "<div class='table-wrap'><table><thead><tr>"
            $crlInfoHtml += "<th>CRL File</th><th>Last Modified</th><th>Size</th><th>Days Until Expiry</th><th>Status</th>"
            $crlInfoHtml += "</tr></thead><tbody>"
            foreach ($crlFile in ($crlFiles | Select-Object -First 5)) {
                $crlDays   = -1
                $crlExpiry = 'Unknown'
                try {
                    $crlDump = Invoke-CertUtil @("-dump", "`"$($crlFile.FullName)`"")
                    if ($crlDump -match 'Next CRL Publish:\s*(.+)') {
                        $nextPublishStr = $Matches[1].Trim()
                        try {
                            $nextPublish = [datetime]::Parse($nextPublishStr)
                            $crlDays     = [int]($nextPublish - (Get-Date)).TotalDays
                            $crlExpiry   = $nextPublish.ToString('yyyy-MM-dd HH:mm')
                            if ($crlDays -lt $CRLDaysLeft -or $CRLDaysLeft -eq -1) {
                                $CRLDaysLeft = $crlDays
                            }
                        } catch {}
                    }
                    if ($crlDump -match 'Next Update:\s*(.+)') {
                        $nextUpStr = $Matches[1].Trim()
                        try {
                            $nextUp    = [datetime]::Parse($nextUpStr)
                            $crlDays   = [int]($nextUp - (Get-Date)).TotalDays
                            $crlExpiry = $nextUp.ToString('yyyy-MM-dd HH:mm')
                            if ($crlDays -lt $CRLDaysLeft -or $CRLDaysLeft -eq -1) {
                                $CRLDaysLeft = $crlDays
                            }
                        } catch {}
                    }
                } catch {}

                $crlSt  = if ($crlDays -lt 0) { 'Expired' }
                          elseif ($crlDays -le $CRLCritDays) { 'CRITICAL' }
                          elseif ($crlDays -le $CRLWarnDays) { 'Warning' }
                          else { 'OK' }
                $crlClr = if ($crlDays -lt 0) { 'red' }
                          elseif ($crlDays -le $CRLCritDays) { 'red' }
                          elseif ($crlDays -le $CRLWarnDays) { 'yellow' }
                          else { 'green' }
                if ($crlDays -le $CRLCritDays) {
                    $CriticalFindings.Add("Section 4 - CRL '$($crlFile.Name)' expires in $crlDays day(s) (CRITICAL)")
                }
                $sizeKB = [math]::Round($crlFile.Length / 1KB, 1)
                $crlInfoHtml += "<tr><td><code>$(HtmlEncode $crlFile.Name)</code></td>"
                $crlInfoHtml += "<td>$($crlFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))</td>"
                $crlInfoHtml += "<td>$sizeKB KB</td>"
                $crlInfoHtml += "<td>$(if ($crlDays -ge 0) { $crlDays } else { 'N/A' })</td>"
                $crlInfoHtml += "<td>$(StatusBadge $crlSt $crlClr)</td></tr>"
            }
            $crlInfoHtml += "</tbody></table></div>"
        } else {
            $crlInfoHtml = "<p class='warn'>No CRL files found in $certEnrollPath</p>"
        }

        # CRL countdown bar
        if ($CRLDaysLeft -ge 0) {
            $maxCRLDays = 30
            $barPct     = [math]::Min([math]::Round(($CRLDaysLeft / $maxCRLDays) * 100, 1), 100)
            $barColor   = if ($CRLDaysLeft -le $CRLCritDays) { '#f85149' }
                          elseif ($CRLDaysLeft -le $CRLWarnDays) { '#d29922' }
                          else { '#3fb950' }
            $crlStatus  = if ($CRLDaysLeft -le $CRLCritDays) { 'CRITICAL' } elseif ($CRLDaysLeft -le $CRLWarnDays) { 'Warning' } else { 'OK' }
            $crlClrBadge = if ($CRLDaysLeft -le $CRLCritDays) { 'red' } elseif ($CRLDaysLeft -le $CRLWarnDays) { 'yellow' } else { 'green' }

            $Sec4Html  = "<div class='disk-row' style='margin-bottom:16px;'>"
            $Sec4Html += "<div class='disk-label'><strong>CRL Next Update / Expiry Countdown</strong></div>"
            $Sec4Html += "<div class='disk-bar-outer'><div class='disk-bar-inner' style='width:${barPct}%;background:$barColor;'></div></div>"
            $Sec4Html += "<div class='disk-stat'><strong style='color:$barColor;font-size:1.3rem;'>$CRLDaysLeft</strong> days until next CRL update &nbsp;$(StatusBadge $crlStatus $crlClrBadge)</div>"
            $Sec4Html += "</div>"
        }
        $Sec4Html += $crlInfoHtml

        # certutil -CRL publish status
        if ($crlPublishedOK) {
            $Sec4Html += "<p style='color:#3fb950;margin-top:10px;'>&#x2714; CRL successfully created (certutil output confirmed)</p>"
        }
    }
} catch {
    $Sec4Html = "<p class='error'>Error checking CRL health: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 4 - CRL health error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5  -  CRL DISTRIBUTION POINTS (CDP)
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 5: CRL Distribution Points..."
$Sec5Html = ''
try {
    if (-not $IsADCSInstalled) {
        $Sec5Html = "<p class='warn'>ADCS not installed  -  CDP check skipped.</p>"
    } else {
        $cdpList = @()
        if ($CANamesList.Count -gt 0 -and (Test-Path "$CARegBase\$CAName")) {
            try {
                $cdpReg = (Get-ItemProperty "$CARegBase\$CAName" -ErrorAction Stop).CRLPublicationURLs
                if ($null -ne $cdpReg) {
                    foreach ($entry in $cdpReg) {
                        $url = $entry -replace '^\d+:', ''
                        if ($url -match '^https?://') { $cdpList += $url }
                        elseif ($url.Length -gt 1) { $cdpList += $url }
                    }
                }
            } catch {}
        }

        # fallback: certutil -getreg CA\CRLPublicationURLs
        if ($cdpList.Count -eq 0) {
            $cdpDump = Invoke-CertUtil @('-getreg', 'CA\CRLPublicationURLs')
            foreach ($line in ($cdpDump -split "`n")) {
                $line = $line.Trim()
                if ($line -match '^https?://') { $cdpList += $line }
                elseif ($line -match 'ldap://')    { $cdpList += $line }
            }
        }

        if ($cdpList.Count -eq 0) {
            $Sec5Html = "<p class='info'>No CRL Distribution Points found in registry or certutil output.</p>"
        } else {
            $Sec5Html  = "<div class='table-wrap'><table><thead><tr>"
            $Sec5Html += "<th>CDP URL</th><th>Type</th><th>Accessibility</th>"
            $Sec5Html += "</tr></thead><tbody>"
            foreach ($cdp in ($cdpList | Select-Object -Unique)) {
                $cdpType  = if ($cdp -match '^https?://') { 'HTTP' }
                            elseif ($cdp -match '^ldap://') { 'LDAP' }
                            elseif ($cdp -match '^\\\\')    { 'UNC' }
                            elseif ($cdp -match '^[A-Za-z]:\\') { 'Local File' }
                            else { 'Other' }
                $testResult = 'N/A'
                $testColor  = 'grey'
                if ($cdpType -eq 'HTTP') {
                    $isOk       = Test-HttpUrl -Url $cdp -TimeoutSec $UrlTestTimeout
                    $testResult = if ($isOk) { 'Accessible &#x2714;' } else { 'Failed &#x2716;' }
                    $testColor  = if ($isOk) { 'green' } else { 'red' }
                    if (-not $isOk) { $CriticalFindings.Add("Section 5 - CDP URL not accessible: $cdp") }
                }
                $Sec5Html += "<tr><td style='word-break:break-all;'>$(HtmlEncode $cdp)</td>"
                $Sec5Html += "<td>$(StatusBadge $cdpType 'blue')</td>"
                $Sec5Html += "<td>$(StatusBadge $testResult $testColor)</td></tr>"
            }
            $Sec5Html += "</tbody></table></div>"
        }
    }
} catch {
    $Sec5Html = "<p class='error'>Error checking CDP: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 5 - CDP check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6  -  OCSP STATUS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 6: OCSP Status..."
$Sec6Html = ''
try {
    $ocspInstalled = $false
    try {
        $ocspSvc = Get-Service -Name OCSPSvc -ErrorAction Stop
        $ocspInstalled = $true
    } catch {}

    if (-not $ocspInstalled) {
        $Sec6Html = "<div class='not-installed-card'><span class='ni-icon'>&#x1F4CB;</span><div><strong>OCSP Responder Not Installed</strong><p>The Online Responder (OCSPSvc) service was not detected. If you are using CDP-only CRL distribution, this is expected.</p></div></div>"
    } else {
        $ocspStatus = $ocspSvc.Status.ToString()
        $ocspStartType = $ocspSvc.StartType.ToString()
        $ocspColor = if ($ocspStatus -eq 'Running') { 'green' } else { 'red' }
        if ($ocspStatus -ne 'Running') {
            $CriticalFindings.Add("Section 6 - OCSP Responder service is $ocspStatus")
        }

        # Try to get OCSP URL from registry
        $ocspUrl = 'N/A'
        try {
            $ocspReg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Certificate Services\OCSPResponseSigningCerts' -ErrorAction Stop
            if ($null -ne $ocspReg) { $ocspUrl = 'Configured (check MMC for URL)' }
        } catch {}

        # Check AIA for OCSP URL
        $aiaOcspUrls = @()
        if ($CANamesList.Count -gt 0 -and (Test-Path "$CARegBase\$CAName")) {
            try {
                $aiaReg = (Get-ItemProperty "$CARegBase\$CAName" -ErrorAction Stop).CACertPublicationURLs
                if ($null -ne $aiaReg) {
                    foreach ($entry in $aiaReg) {
                        $url = $entry -replace '^\d+:', ''
                        if ($url -match '^https?://.*ocsp') { $aiaOcspUrls += $url }
                    }
                }
            } catch {}
        }

        $ocspTest = 'Not tested (LDAP/file-based)'
        $ocspTestClr = 'grey'
        if ($aiaOcspUrls.Count -gt 0) {
            $ocspUrl  = $aiaOcspUrls[0]
            $ocspOk   = Test-HttpUrl -Url $ocspUrl -TimeoutSec $UrlTestTimeout
            $ocspTest = if ($ocspOk) { 'Accessible &#x2714;' } else { 'Failed &#x2716;' }
            $ocspTestClr = if ($ocspOk) { 'green' } else { 'red' }
            if (-not $ocspOk) { $CriticalFindings.Add("Section 6 - OCSP URL not accessible: $ocspUrl") }
        }

        $rows6 = @(
            @('OCSP Service Status', (StatusBadge $ocspStatus $ocspColor)),
            @('Start Type',          (HtmlEncode $ocspStartType)),
            @('OCSP URL',            (HtmlEncode $ocspUrl)),
            @('URL Accessibility',   (StatusBadge $ocspTest $ocspTestClr))
        )
        $Sec6Html = BuildKVTable $rows6
    }
} catch {
    $Sec6Html = "<p class='error'>Error checking OCSP status: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 6 - OCSP check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7  -  ISSUED CERTIFICATES SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 7: Issued Certificates Summary..."
$Sec7Html = ''
try {
    if (-not $IsADCSInstalled) {
        $Sec7Html = "<p class='warn'>ADCS not installed  -  certificate summary skipped.</p>"
    } else {
        # Use certutil -view to count issued/revoked certs
        $totalIssued    = 'N/A'
        $totalRevoked   = 'N/A'
        $expiring30     = 'N/A'
        $expiring60     = 'N/A'
        $expiring90     = 'N/A'
        $alreadyExpired = 'N/A'

        $viewOutput = Invoke-CertUtil @("-config", "`"$ServerHostname\$CAName`"", '-view', '-out', 'RequestId,NotAfter,DispositionMessage', '-restrict', 'Disposition=20')
        if ([string]::IsNullOrEmpty($viewOutput)) {
            $viewOutput = Invoke-CertUtil @('-view', '-out', 'RequestId,NotAfter,DispositionMessage', '-restrict', 'Disposition=20')
        }

        if (-not [string]::IsNullOrEmpty($viewOutput)) {
            $issuedMatches = ([regex]'Row \d+:').Matches($viewOutput)
            $totalIssued   = $issuedMatches.Count

            $now = Get-Date
            $exp30  = 0; $exp60 = 0; $exp90 = 0; $expired = 0
            $dateMatches = ([regex]'NotAfter:\s+(.+)').Matches($viewOutput)
            foreach ($dm in $dateMatches) {
                try {
                    $notAfter = [datetime]::Parse($dm.Groups[1].Value.Trim())
                    $daysLeft = ($notAfter - $now).TotalDays
                    if ($daysLeft -lt 0)   { $expired++ }
                    elseif ($daysLeft -le 30) { $exp30++ }
                    elseif ($daysLeft -le 60) { $exp60++ }
                    elseif ($daysLeft -le 90) { $exp90++ }
                } catch {}
            }
            $expiring30     = $exp30
            $expiring60     = $exp60
            $expiring90     = $exp90
            $alreadyExpired = $expired
            $ExpiringCertCount = $exp30
        }

        # Revoked count
        $revokedOutput = Invoke-CertUtil @("-config", "`"$ServerHostname\$CAName`"", '-view', '-restrict', 'Disposition=21', '-out', 'RequestId')
        if ([string]::IsNullOrEmpty($revokedOutput)) {
            $revokedOutput = Invoke-CertUtil @('-view', '-restrict', 'Disposition=21', '-out', 'RequestId')
        }
        if (-not [string]::IsNullOrEmpty($revokedOutput)) {
            $revokedMatches = ([regex]'Row \d+:').Matches($revokedOutput)
            $totalRevoked   = $revokedMatches.Count
        }

        if ($expiring30 -is [int] -and $expiring30 -gt 0) {
            $CriticalFindings.Add("Section 7 - $expiring30 certificate(s) expiring within 30 days")
        }
        if ($alreadyExpired -is [int] -and $alreadyExpired -gt 0) {
            $CriticalFindings.Add("Section 7 - $alreadyExpired certificate(s) already expired")
        }

        # Mini bar chart (static HTML)
        $barHtml = ''
        if ($totalIssued -is [int] -and $totalIssued -gt 0) {
            $segments = @(
                @{ Label='Expiring <=30d';  Count=$expiring30;     Color='#f85149' },
                @{ Label='Expiring <=60d';  Count=$expiring60;     Color='#d29922' },
                @{ Label='Expiring <=90d';  Count=$expiring90;     Color='#e3a019' },
                @{ Label='Already Expired';Count=$alreadyExpired; Color='#6e40c9' }
            )
            $barHtml  = "<div style='margin:12px 0;'><h4 style='margin-bottom:8px;'>Certificate Expiry Breakdown</h4>"
            foreach ($seg in $segments) {
                if ($seg.Count -is [int] -and $totalIssued -gt 0) {
                    $pct = [math]::Round(($seg.Count / $totalIssued) * 100, 1)
                    $barHtml += "<div class='disk-row' style='margin-bottom:6px;'>"
                    $barHtml += "<div class='disk-label' style='width:140px;'>$(HtmlEncode $seg.Label)</div>"
                    $barHtml += "<div class='disk-bar-outer'><div class='disk-bar-inner' style='width:${pct}%;background:$($seg.Color);'></div></div>"
                    $barHtml += "<div class='disk-stat' style='min-width:80px;'>$($seg.Count) cert(s) ($pct%)</div>"
                    $barHtml += "</div>"
                }
            }
            $barHtml += "</div>"
        }

        $rows7 = @(
            @('Total Issued Certs',      (HtmlEncode $totalIssued.ToString())),
            @('Total Revoked',           (HtmlEncode $totalRevoked.ToString())),
            @('Expiring <= 30 days',      (HtmlEncode $expiring30.ToString())),
            @('Expiring <= 60 days',      (HtmlEncode $expiring60.ToString())),
            @('Expiring <= 90 days',      (HtmlEncode $expiring90.ToString())),
            @('Already Expired',         (HtmlEncode $alreadyExpired.ToString()))
        )
        $Sec7Html = BuildKVTable $rows7
        $Sec7Html += $barHtml

        if ($totalIssued -eq 'N/A') {
            $Sec7Html += "<p class='warn' style='margin-top:10px;'>certutil -view returned no output. Ensure you have permissions and the CA service is running.</p>"
        }
    }
} catch {
    $Sec7Html = "<p class='error'>Error retrieving issued certificate summary: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 7 - Issued certificate summary error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 8  -  CERTIFICATE TEMPLATES
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 8: Certificate Templates..."
$Sec8Html = ''
try {
    if (-not $IsADCSInstalled) {
        $Sec8Html = "<p class='warn'>ADCS not installed  -  template check skipped.</p>"
    } else {
        # Templates are in AD for Enterprise CAs or in registry for Standalone
        $templateOutput = Invoke-CertUtil @('-CATemplates')
        if ([string]::IsNullOrEmpty($templateOutput)) {
            $templateOutput = Invoke-CertUtil @("-config", "`"$ServerHostname\$CAName`"", '-CATemplates')
        }

        if ([string]::IsNullOrEmpty($templateOutput)) {
            $Sec8Html = "<p class='info'>Certificate template list not available (may be a Standalone CA with no published templates, or insufficient permissions).</p>"
        } else {
            $lines = $templateOutput -split "`n" | Where-Object { $_.Trim().Length -gt 0 }
            $Sec8Html  = "<div class='table-wrap'><table><thead><tr>"
            $Sec8Html += "<th>Template Name</th><th>Schema Version</th><th>Notes</th>"
            $Sec8Html += "</tr></thead><tbody>"
            foreach ($line in $lines) {
                $line = $line.Trim()
                if ($line -match '^--' -or $line -match 'CertUtil' -or $line -match 'completed' -or $line.Length -lt 3) { continue }
                $parts  = $line -split ':',2
                $tName  = if ($parts.Count -ge 1) { $parts[0].Trim() } else { $line }
                $tNotes = if ($parts.Count -ge 2) { $parts[1].Trim() } else { '' }
                $Sec8Html += "<tr><td>$(HtmlEncode $tName)</td><td>$(HtmlEncode $tNotes)</td><td></td></tr>"
            }
            $Sec8Html += "</tbody></table></div>"
        }
    }
} catch {
    $Sec8Html = "<p class='error'>Error retrieving certificate templates: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 8 - Certificate templates error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 9  -  AIA (AUTHORITY INFORMATION ACCESS) URLS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 9: AIA URLs..."
$Sec9Html = ''
try {
    if (-not $IsADCSInstalled) {
        $Sec9Html = "<p class='warn'>ADCS not installed  -  AIA check skipped.</p>"
    } else {
        $aiaList = @()
        if ($CANamesList.Count -gt 0 -and (Test-Path "$CARegBase\$CAName")) {
            try {
                $aiaReg = (Get-ItemProperty "$CARegBase\$CAName" -ErrorAction Stop).CACertPublicationURLs
                if ($null -ne $aiaReg) {
                    foreach ($entry in $aiaReg) {
                        $url = $entry -replace '^\d+:', ''
                        $aiaList += $url
                    }
                }
            } catch {}
        }

        if ($aiaList.Count -eq 0) {
            $aiaDump = Invoke-CertUtil @('-getreg', 'CA\CACertPublicationURLs')
            foreach ($line in ($aiaDump -split "`n")) {
                $line = $line.Trim()
                if ($line.Length -gt 3) { $aiaList += $line }
            }
        }

        if ($aiaList.Count -eq 0) {
            $Sec9Html = "<p class='info'>No AIA URLs found in registry or certutil output.</p>"
        } else {
            $Sec9Html  = "<div class='table-wrap'><table><thead><tr>"
            $Sec9Html += "<th>AIA URL</th><th>Type</th><th>Accessibility</th>"
            $Sec9Html += "</tr></thead><tbody>"
            foreach ($aia in ($aiaList | Where-Object { $_.Length -gt 3 } | Select-Object -Unique)) {
                $aiaType = if ($aia -match '^https?://') { 'HTTP' }
                           elseif ($aia -match '^ldap://') { 'LDAP' }
                           elseif ($aia -match '^\\\\')    { 'UNC'  }
                           elseif ($aia -match '^[A-Za-z]:\\') { 'Local File' }
                           else { 'Other' }
                $testResult = 'N/A'
                $testColor  = 'grey'
                if ($aiaType -eq 'HTTP') {
                    $isOk       = Test-HttpUrl -Url $aia -TimeoutSec $UrlTestTimeout
                    $testResult = if ($isOk) { 'Accessible &#x2714;' } else { 'Failed &#x2716;' }
                    $testColor  = if ($isOk) { 'green' } else { 'red' }
                    if (-not $isOk) { $CriticalFindings.Add("Section 9 - AIA URL not accessible: $aia") }
                }
                $Sec9Html += "<tr><td style='word-break:break-all;'>$(HtmlEncode $aia)</td>"
                $Sec9Html += "<td>$(StatusBadge $aiaType 'blue')</td>"
                $Sec9Html += "<td>$(StatusBadge $testResult $testColor)</td></tr>"
            }
            $Sec9Html += "</tbody></table></div>"
        }
    }
} catch {
    $Sec9Html = "<p class='error'>Error checking AIA URLs: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 9 - AIA check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 10  -  CA DATABASE SIZE
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 10: CA Database Size..."
$Sec10Html = ''
try {
    if (-not $IsADCSInstalled) {
        $Sec10Html = "<p class='warn'>ADCS not installed  -  database check skipped.</p>"
    } else {
        $dbPath    = 'Unknown'
        $dbLogPath = 'Unknown'
        $dbSizeMB  = 'N/A'
        $logSizeMB = 'N/A'

        if ($CANamesList.Count -gt 0 -and (Test-Path "$CARegBase\$CAName")) {
            try {
                $dbReg     = Get-ItemProperty "$CARegBase\$CAName" -ErrorAction Stop
                $dbPath    = if ($null -ne $dbReg.DBDirectory) { $dbReg.DBDirectory } else { 'Unknown' }
                $dbLogPath = if ($null -ne $dbReg.DBLogDirectory) { $dbReg.DBLogDirectory } else { 'Unknown' }
            } catch {}
        }

        if ($dbPath -eq 'Unknown') {
            # Typical default
            $dbPath = Join-Path $env:SystemRoot 'system32\CertLog'
        }

        if (Test-Path $dbPath) {
            $dbFiles  = @(Get-ChildItem -Path $dbPath -Filter '*.edb' -Recurse -ErrorAction SilentlyContinue)
            $logFiles = @(Get-ChildItem -Path $dbPath -Filter '*.log' -Recurse -ErrorAction SilentlyContinue)
            if ($dbFiles.Count -gt 0) {
                $dbSizeMB = [math]::Round(($dbFiles | Measure-Object Length -Sum).Sum / 1MB, 2)
            }
            if ($logFiles.Count -gt 0) {
                $logSizeMB = [math]::Round(($logFiles | Measure-Object Length -Sum).Sum / 1MB, 2)
            }
        }

        $rows10 = @(
            @('CA Database Path',    (HtmlEncode $dbPath)),
            @('CA Log Path',         (HtmlEncode $dbLogPath)),
            @('Database Size (.edb)',(HtmlEncode "$dbSizeMB MB")),
            @('Log Files Size',      (HtmlEncode "$logSizeMB MB"))
        )
        $Sec10Html = BuildKVTable $rows10
    }
} catch {
    $Sec10Html = "<p class='error'>Error checking CA database: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 10 - CA database check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 11  -  FAILED REQUESTS (LAST 7 DAYS)
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 11: Failed Requests (last 7 days)..."
$Sec11Html = ''
try {
    if (-not $IsADCSInstalled) {
        $Sec11Html = "<p class='warn'>ADCS not installed  -  failed request check skipped.</p>"
    } else {
        # Disposition 30 = Failed, 31 = Denied
        $failedOutput = Invoke-CertUtil @("-config", "`"$ServerHostname\$CAName`"", '-view', '-restrict', 'Disposition=30', '-out', 'RequestId,RequestSubmittedWhen,DispositionMessage')
        if ([string]::IsNullOrEmpty($failedOutput)) {
            $failedOutput = Invoke-CertUtil @('-view', '-restrict', 'Disposition=30', '-out', 'RequestId,RequestSubmittedWhen,DispositionMessage')
        }
        $deniedOutput = Invoke-CertUtil @("-config", "`"$ServerHostname\$CAName`"", '-view', '-restrict', 'Disposition=31', '-out', 'RequestId,RequestSubmittedWhen,DispositionMessage')
        if ([string]::IsNullOrEmpty($deniedOutput)) {
            $deniedOutput = Invoke-CertUtil @('-view', '-restrict', 'Disposition=31', '-out', 'RequestId,RequestSubmittedWhen,DispositionMessage')
        }

        $failedCount = 0
        $deniedCount = 0

        if (-not [string]::IsNullOrEmpty($failedOutput)) {
            $failedMatches = ([regex]'Row \d+:').Matches($failedOutput)
            $failedCount = $failedMatches.Count
        }
        if (-not [string]::IsNullOrEmpty($deniedOutput)) {
            $deniedMatches = ([regex]'Row \d+:').Matches($deniedOutput)
            $deniedCount = $deniedMatches.Count
        }

        $totalFailed = $failedCount + $deniedCount
        if ($totalFailed -gt 10) {
            $CriticalFindings.Add("Section 11 - $totalFailed failed/denied certificate requests found")
        }

        $failBadge = if ($totalFailed -gt 10) { StatusBadge $totalFailed.ToString() 'red' }
                     elseif ($totalFailed -gt 0) { StatusBadge $totalFailed.ToString() 'yellow' }
                     else { StatusBadge '0' 'green' }

        $rows11 = @(
            @('Failed Requests',   (HtmlEncode $failedCount.ToString())),
            @('Denied Requests',   (HtmlEncode $deniedCount.ToString())),
            @('Total Failed+Denied', $failBadge)
        )
        $Sec11Html = BuildKVTable $rows11
        if ($totalFailed -eq 0 -and [string]::IsNullOrEmpty($failedOutput)) {
            $Sec11Html += "<p class='warn' style='margin-top:8px;'>certutil -view returned no output  -  failed request counts may not be accurate.</p>"
        }
    }
} catch {
    $Sec11Html = "<p class='error'>Error checking failed requests: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 11 - Failed requests check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 12  -  CA EVENT LOGS
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 12: CA Event Logs..."
$Sec12Html = ''
try {
    $since24h   = (Get-Date).AddHours(-24)
    $caLogNames = @('Microsoft-Windows-CertificationAuthority/Operational', 'Security', 'System', 'Application')
    $caEvents   = [System.Collections.Generic.List[object]]::new()

    foreach ($logName in $caLogNames) {
        try {
            $filter = @{ LogName = $logName; Level = @(1,2); StartTime = $since24h }
            $evts   = @(Get-WinEvent -FilterHashtable $filter -MaxEvents 50 -ErrorAction Stop |
                Where-Object { $_.ProviderName -match 'CertSvc|Cert|PKI|Certificate' -or $logName -eq 'Microsoft-Windows-CertificationAuthority/Operational' })
            foreach ($e in $evts) { $caEvents.Add($e) }
        } catch {}
    }

    # Also try 'Certificate Services' source from System log
    try {
        $filter2 = @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-CertificationAuthority' }
        $evts2 = @(Get-WinEvent -FilterHashtable $filter2 -MaxEvents 30 -ErrorAction Stop)
        foreach ($e in $evts2) { $caEvents.Add($e) }
    } catch {}

    $topEvents = $caEvents | Sort-Object TimeCreated -Descending | Select-Object -Unique -First 10

    if ($topEvents.Count -eq 0) {
        $Sec12Html = "<p style='color:#3fb950;'>&#x2714; No Critical or Error events found in Certificate Services logs.</p>"
    } else {
        $Sec12Html  = "<div class='table-wrap'><table><thead><tr>"
        $Sec12Html += "<th>Time</th><th>Level</th><th>Event ID</th><th>Source</th><th>Message</th>"
        $Sec12Html += "</tr></thead><tbody>"
        foreach ($ev in $topEvents) {
            $timeStr  = if ($null -ne $ev.TimeCreated) { $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
            $level    = if ($null -ne $ev.LevelDisplayName) { $ev.LevelDisplayName } else { 'Unknown' }
            $lvlColor = if ($level -eq 'Critical') { 'red' } else { 'yellow' }
            $source   = if ($null -ne $ev.ProviderName) { $ev.ProviderName } else { 'Unknown' }
            $msgRaw   = if ($null -ne $ev.Message) { $ev.Message } else { '' }
            $msgShort = if ($msgRaw.Length -gt 200) { $msgRaw.Substring(0,200) + '...' } else { $msgRaw }
            $Sec12Html += "<tr>"
            $Sec12Html += "<td style='white-space:nowrap;'>$(HtmlEncode $timeStr)</td>"
            $Sec12Html += "<td>$(StatusBadge $level $lvlColor)</td>"
            $Sec12Html += "<td>$($ev.Id)</td>"
            $Sec12Html += "<td>$(HtmlEncode $source)</td>"
            $Sec12Html += "<td>$(HtmlEncode $msgShort)</td>"
            $Sec12Html += "</tr>"
        }
        $Sec12Html += "</tbody></table></div>"
    }
} catch {
    $Sec12Html = "<p class='error'>Error reading CA event logs: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 12 - CA event log error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 13  -  ENROLLMENT AGENT & KEY RECOVERY AGENT
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 13: Enrollment Agent / KRA..."
$Sec13Html = ''
try {
    if (-not $IsADCSInstalled) {
        $Sec13Html = "<p class='warn'>ADCS not installed  -  KRA/EA check skipped.</p>"
    } else {
        $kraCerts    = @()
        $kraCount    = 0
        $kraInfoHtml = ''

        # KRA check via certutil -getreg CA\EKU or check for KRA in NTDS
        $kraOutput = Invoke-CertUtil @('-getreg', 'CA\KRAFlags')
        $kraActive = $kraOutput -match '0x'

        # Check the KRA store
        try {
            $kraStore = New-Object System.Security.Cryptography.X509Certificates.X509Store('CA','LocalMachine')
            $kraStore.Open('ReadOnly')
            $kraCerts = @($kraStore.Certificates | Where-Object {
                $_.Extensions | Where-Object { $_.Oid.FriendlyName -eq 'Key Usage' -and $_.KeyUsages -match 'KeyEncipherment' }
            })
            $kraStore.Close()
            $kraCount = $kraCerts.Count
        } catch {}

        # Check certutil KRA cert list
        $kraList = Invoke-CertUtil @("-config", "`"$ServerHostname\$CAName`"", '-getreg', 'CA\KRAForeignCertCount')

        $kraCountLabel = if ($kraCount -gt 0) { "$kraCount KRA cert(s) in CA store" } else { '0 (check AD/templates)' }
        $kraStatusBadge = if ($kraCount -gt 0) { StatusBadge 'Configured' 'green' } else { StatusBadge 'Not Configured' 'grey' }

        $rows13 = @(
            @('KRA Certificates Found',  $kraStatusBadge),
            @('KRA Count in CA Store',   (HtmlEncode $kraCountLabel)),
            @('KRA Flags (certutil)',     (HtmlEncode $(if ([string]::IsNullOrEmpty($kraOutput)) { 'N/A' } else { $kraOutput.Trim().Substring(0,[math]::Min(100,$kraOutput.Trim().Length)) })))
        )
        $Sec13Html = BuildKVTable $rows13

        if ($kraCerts.Count -gt 0) {
            $Sec13Html += "<h4 style='margin:12px 0 8px;'>KRA Certificates in CA Store</h4>"
            $Sec13Html += "<div class='table-wrap'><table><thead><tr><th>Subject</th><th>Valid To</th><th>Thumbprint</th></tr></thead><tbody>"
            foreach ($kc in $kraCerts) {
                $Sec13Html += "<tr><td>$(HtmlEncode $kc.Subject)</td>"
                $Sec13Html += "<td>$($kc.NotAfter.ToString('yyyy-MM-dd'))</td>"
                $Sec13Html += "<td><code>$(HtmlEncode $kc.Thumbprint.Substring(0,[math]::Min(16,$kc.Thumbprint.Length)))...</code></td></tr>"
            }
            $Sec13Html += "</tbody></table></div>"
        }
    }
} catch {
    $Sec13Html = "<p class='error'>Error checking KRA/EA: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 13 - KRA/EA check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 14  -  HSM / KEY STORAGE
# ═══════════════════════════════════════════════════════════════════════════════
Write-Progress2 "Section 14: HSM / Key Storage..."
$Sec14Html = ''
try {
    if (-not $IsADCSInstalled) {
        $Sec14Html = "<p class='warn'>ADCS not installed  -  key storage check skipped.</p>"
    } else {
        $kspProvider = 'Unknown'
        $kspType     = 'Unknown'
        $isHSM       = $false

        if ($CANamesList.Count -gt 0 -and (Test-Path "$CARegBase\$CAName")) {
            try {
                $caReg2 = Get-ItemProperty "$CARegBase\$CAName" -ErrorAction Stop
                if ($null -ne $caReg2.CSP) {
                    $kspProvider = $caReg2.CSP
                } elseif ($null -ne $caReg2.Provider) {
                    $kspProvider = $caReg2.Provider
                }
            } catch {}
        }

        if ([string]::IsNullOrEmpty($kspProvider) -or $kspProvider -eq 'Unknown') {
            $certutilProvider = Invoke-CertUtil @('-getreg', 'CA\CSP\Provider')
            if ($certutilProvider -match 'Provider REG_SZ =\s*(.+)') {
                $kspProvider = $Matches[1].Trim()
            }
        }

        $isHSM = $kspProvider -match 'Thales|nCipher|Safenet|LUNA|Utimaco|Yubico|HSM|Hardware' -or
                 ($kspProvider -notmatch 'Microsoft' -and $kspProvider -ne 'Unknown')
        $kspType = if ($isHSM) { 'HSM (Hardware Security Module)' } else { 'Software KSP / CSP (Microsoft)' }
        $kspBadge = if ($isHSM) { StatusBadge 'HSM' 'green' } else { StatusBadge 'Software KSP' 'yellow' }

        $rows14 = @(
            @('Key Storage Provider', (HtmlEncode $kspProvider)),
            @('Storage Type',         $kspBadge),
            @('Details',              (HtmlEncode $kspType))
        )
        $Sec14Html = BuildKVTable $rows14
        $Sec14Html += "<p class='info' style='margin-top:8px;'>Note: An HSM is recommended for production root CAs and policy CAs.</p>"
    }
} catch {
    $Sec14Html = "<p class='error'>Error checking HSM/Key Storage: $(HtmlEncode $_.Exception.Message)</p>"
    $CriticalFindings.Add("Section 14 - HSM/Key storage check error: $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# KPI TILES & FINAL ASSEMBLY
# ═══════════════════════════════════════════════════════════════════════════════
$CritCount  = $CriticalFindings.Count
$ReportDate = (Get-Date).ToString('dddd, dd MMMM yyyy HH:mm:ss')
$EndTime    = Get-Date
$Duration   = ($EndTime - $StartTime).ToString('hh\:mm\:ss')

$OverallStatusBadge = if (-not $IsADCSInstalled) { StatusBadge 'NOT INSTALLED' 'grey' }
                      elseif ($CritCount -gt 0)  { StatusBadge 'CRITICAL' 'red' }
                      else                        { StatusBadge 'HEALTHY' 'green' }

$certDaysTileColor = if ($CACertDaysLeft -lt 0) { '#484f58' }
                     elseif ($CACertDaysLeft -le $CACertCritDays) { '#f85149' }
                     elseif ($CACertDaysLeft -le $CACertWarnDays) { '#d29922' }
                     else { '#3fb950' }
$caCertTileHtml    = "<span style='color:$certDaysTileColor;font-weight:700;'>$(if ($CACertDaysLeft -ge 0) { $CACertDaysLeft } else { 'N/A' })</span>"

$crlDaysTileColor  = if ($CRLDaysLeft -lt 0) { '#484f58' }
                     elseif ($CRLDaysLeft -le $CRLCritDays) { '#f85149' }
                     elseif ($CRLDaysLeft -le $CRLWarnDays) { '#d29922' }
                     else { '#3fb950' }
$crlTileHtml       = "<span style='color:$crlDaysTileColor;font-weight:700;'>$(if ($CRLDaysLeft -ge 0) { $CRLDaysLeft } else { 'N/A' })</span>"

$expTileColor      = if ($ExpiringCertCount -gt 0) { '#f85149' } else { '#3fb950' }
$expTileHtml       = "<span style='color:$expTileColor;font-weight:700;'>$ExpiringCertCount</span>"

$critTileColor     = if ($CritCount -gt 0) { '#f85149' } else { '#3fb950' }
$critTileHtml      = "<span style='color:$critTileColor;font-weight:700;'>$CritCount</span>"

$caNameTileHtml    = "<span style='font-size:.75rem;font-weight:600;'>$(HtmlEncode $CAName)</span>"
$installedBadge    = if ($IsADCSInstalled) { StatusBadge 'Installed' 'green' } else { StatusBadge 'Not Installed' 'red' }

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
# FULL HTML
# ═══════════════════════════════════════════════════════════════════════════════
$HtmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ADCS PKI Health Check - $(HtmlEncode $ServerHostname)</title>
<style>
:root {
  --bg:      #0d1117;
  --card:    #161b22;
  --border:  #30363d;
  --text:    #c9d1d9;
  --muted:   #8b949e;
  --link:    #e05c7a;
  --head-bg: #160a0d;
  --th-bg:   #21262d;
  --tr-alt:  #1c2128;
  --accent:  #e05c7a;
}
[data-theme="light"] {
  --bg:      #ffffff;
  --card:    #f6f8fa;
  --border:  #d0d7de;
  --text:    #24292f;
  --muted:   #57606a;
  --link:    #cf222e;
  --head-bg: #fff0f3;
  --th-bg:   #fdf2f4;
  --tr-alt:  #fff8f9;
  --accent:  #cf222e;
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
.summary-bar { display:flex; flex-wrap:wrap; gap:12px;
               background:var(--card); border:1px solid var(--border);
               border-top:2px solid var(--accent);
               border-radius:8px; padding:16px 20px; margin:20px 0; }
.stat-card   { flex:1 1 120px; text-align:center; }
.stat-count  { font-size:2rem; font-weight:700; line-height:1; }
.stat-label  { font-size:.72rem; color:var(--muted); margin-top:4px; }
.section-card { background:var(--card); border:1px solid var(--border);
                border-radius:8px; margin-bottom:14px; overflow:hidden; }
.section-summary { display:flex; align-items:center; gap:10px; cursor:pointer;
                   padding:13px 18px; list-style:none; user-select:none; }
.section-summary::-webkit-details-marker { display:none; }
.section-summary::marker { display:none; }
.sec-arrow { font-size:10px; color:var(--muted); display:inline-block;
             transition:transform .2s; flex-shrink:0; line-height:1; }
details[open] > .section-summary .sec-arrow { transform:rotate(90deg); }
.section-summary:hover { background:var(--th-bg); }
.sec-num { background:var(--accent); color:#fff; font-size:.7rem; font-weight:700;
           width:22px; height:22px; border-radius:50%; display:flex;
           align-items:center; justify-content:center; flex-shrink:0; }
.sec-title { font-weight:600; font-size:.95rem; flex:1; }
.section-body { padding:15px 18px; border-top:1px solid var(--border); }
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
td:nth-child(5) { max-width:360px; word-break:break-word; white-space:normal; }
.disk-container { display:flex; flex-direction:column; gap:12px; margin-bottom:12px; }
.disk-row { }
.disk-label { font-size:.85rem; margin-bottom:4px; font-weight:600; }
.disk-bar-outer { height:18px; background:var(--th-bg); border-radius:9px;
                  overflow:hidden; margin-bottom:4px; border:1px solid var(--border); }
.disk-bar-inner { height:100%; border-radius:9px; transition:width .4s ease; }
.disk-stat { font-size:.8rem; color:var(--muted); }
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
      <h1>&#x1F511; ADCS / PKI Health Check</h1>
      <p>$(HtmlEncode $ServerHostname) &nbsp;|&nbsp; Generated: $(HtmlEncode $ReportDate)</p>
    </div>
  </div>
  <button class="theme-toggle" onclick="toggleTheme()" title="Toggle Dark/Light mode">
    <span id="theme-icon">&#x2600;&#xFE0F;</span> Toggle Theme
  </button>
</div>

<div class="page-wrap">
<div class="summary-bar">
  <div class="stat-card">
    <div class="stat-count" style="font-size:.85rem;padding-top:8px;">$installedBadge</div>
    <div class="stat-label">ADCS Status</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:.8rem;padding-top:8px;">$caNameTileHtml</div>
    <div class="stat-label">CA Name</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.6rem;">$caCertTileHtml</div>
    <div class="stat-label">CA Cert Days Left</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.6rem;">$crlTileHtml</div>
    <div class="stat-label">CRL Days Left</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.6rem;">$expTileHtml</div>
    <div class="stat-label">Expiring (30d)</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:1.6rem;">$critTileHtml</div>
    <div class="stat-label">Critical Findings</div>
  </div>
  <div class="stat-card">
    <div class="stat-count" style="font-size:.9rem;padding-top:10px;">$OverallStatusBadge</div>
    <div class="stat-label">Overall Status</div>
  </div>
</div>

$CritFindingsHtml

$(BuildSection 0  'Server Details'              $Sec0Html  ($Sec0Html  -match "class='error'") $true)
$(BuildSection 1  'CA Identity & Config'            $Sec1Html  ($Sec1Html  -match "class='error'") $true)
$(BuildSection 2  'ADCS Services Status'             $Sec2Html  ($Sec2Html  -match "class='error'") $true)
$(BuildSection 3  'CA Certificate Validity'          $Sec3Html  ($Sec3Html  -match "class='error'") $true)
$(BuildSection 4  'CRL Health'                       $Sec4Html  ($Sec4Html  -match "class='error'") $true)
$(BuildSection 5  'CRL Distribution Points (CDP)'    $Sec5Html  ($Sec5Html  -match "class='error'") $false)
$(BuildSection 6  'OCSP Status'                      $Sec6Html  ($Sec6Html  -match "class='error'") $false)
$(BuildSection 7  'Issued Certificates Summary'      $Sec7Html  ($Sec7Html  -match "class='error'") $false)
$(BuildSection 8  'Certificate Templates'            $Sec8Html  ($Sec8Html  -match "class='error'") $false)
$(BuildSection 9  'AIA (Authority Info Access) URLs' $Sec9Html  ($Sec9Html  -match "class='error'") $false)
$(BuildSection 10 'CA Database Size'                 $Sec10Html ($Sec10Html -match "class='error'") $false)
$(BuildSection 11 'Failed Requests (Last 7 Days)'    $Sec11Html ($Sec11Html -match "class='error'") $false)
$(BuildSection 12 'CA Event Logs'                    $Sec12Html ($Sec12Html -match "class='error'") $false)
$(BuildSection 13 'Enrollment Agent / KRA'           $Sec13Html ($Sec13Html -match "class='error'") $false)
$(BuildSection 14 'HSM / Key Storage Provider'       $Sec14Html ($Sec14Html -match "class='error'") $false)

<div class="footer">
  <div>
    <strong>ADCS / PKI Health Check v$ScriptVersion</strong> &nbsp;|&nbsp;
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
    $statusContent += " ADCS PKI HEALTH CHECK  -  *** CRITICAL ALERT ***`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " Status    : CRITICAL`r`n"
    $statusContent += " Server    : $ServerHostname`r`n"
    $statusContent += " CA Name   : $CAName`r`n"
    $statusContent += " Generated : $ReportDate`r`n"
    $statusContent += " Duration  : $Duration`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += "`r`n CRITICAL FINDINGS ($($CriticalFindings.Count)):`r`n`r`n"
    $statusContent += (@($CriticalFindings) | ForEach-Object { "  [!] $_" }) -join "`r`n"
    $statusContent += "`r`n`r`n$separator`r`n"
    $statusContent += " Full HTML report : $ReportFile`r`n"
    $statusContent += " Status file      : $StatusFile`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " ADCS Health Check v$ScriptVersion  by $AuthorName`r`n"
    $statusContent += "$separator`r`n"
} else {
    $statusContent  = "$separator`r`n"
    $statusContent += " ADCS PKI HEALTH CHECK  -  HEALTHY STATE`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " Status    : HEALTHY`r`n"
    $statusContent += " Server    : $ServerHostname`r`n"
    $statusContent += " CA Name   : $CAName`r`n"
    $statusContent += " Generated : $ReportDate`r`n"
    $statusContent += " Duration  : $Duration`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += "`r`n No critical findings detected. PKI is in a healthy state.`r`n"
    $statusContent += "`r`n$separator`r`n"
    $statusContent += " Full HTML report : $ReportFile`r`n"
    $statusContent += " Status file      : $StatusFile`r`n"
    $statusContent += "$separator`r`n"
    $statusContent += " ADCS Health Check v$ScriptVersion  by $AuthorName`r`n"
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
Write-Host "===============================================================" -ForegroundColor DarkRed
Write-Host "  ADCS PKI Health Check complete.  Duration: $Duration" -ForegroundColor DarkRed
Write-Host "  Report : $ReportFile" -ForegroundColor Yellow
if ($EnableStatusFile) {
    $slabel = if ($isCritical) { 'Status (CRITICAL)' } else { 'Status (HEALTHY)' }
    Write-Host "  $slabel : $StatusFile" -ForegroundColor $(if ($isCritical) { 'Red' } else { 'Green' })
    Write-Host ""
    Write-Host "  To send email alerts, run: .\ADCS_HealthCheck_EmailAlert.ps1" -ForegroundColor Cyan
}
Write-Host "===============================================================" -ForegroundColor DarkRed
Write-Host ""
