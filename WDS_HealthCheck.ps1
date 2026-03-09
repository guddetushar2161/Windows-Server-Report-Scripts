#Requires -Version 5.1
<#
.SYNOPSIS
    Windows Deployment Services (WDS) Server Health Check Script

.DESCRIPTION
    Performs a comprehensive health check of the Windows Deployment Services
    role on a Windows Server and exports all results to a single self-contained
    HTML dashboard file.

    If the WDS role is not installed the script exits gracefully with a
    "Role Not Installed" HTML section rather than crashing.  Every section
    is individually wrapped in Try/Catch.

    Checks performed:
      1.  WDS Service Status          (service state, start type, PID, last start)
      2.  WDS Configuration           (initialised, remote-install path, AD vs standalone,
                                       DHCP auth, answer policy)
      3.  PXE & TFTP Health           (TFTP port-69 netstat, PXE policy, DHCP option notes)
      4.  Boot Images                 (name, arch, version, size, creation date, WinPE version;
                                       flag images older than 180 days)
      5.  Install Images              (image groups and images: OS, arch, size, creation date)
      6.  Multicast Sessions          (active transmissions, client count, progress, rate)
      7.  Recent Deployment Activity  (last 20 entries from WDS log files in
                                       %WINDIR%\System32\LogFiles\WDS\)
      8.  Pending Devices             (computers awaiting admin approval)
      9.  WDS DHCP Integration        (DHCP option 60 conflict / co-host note)
     10.  Disk Space                  (RemoteInstall volume; warn <20 GB, crit <10 GB)
     11.  WDS-related Services        (WDSSERVER, WDSTFTP, BINLSVC, WDSTransportServer)
     12.  Active Directory            (WDS computer account exists, DHCP authorisation)
     13.  Network Check               (DC DNS resolution + LDAP port 389 test)
     14.  WDS Event Log               (last 20 events from Windows Deployment Services log)
     15.  System Event Log            (last 10 Critical/Error events mentioning WDS/TFTP/BINL)

.NOTES
    Version    : 1.0.0
    Author     : Tushar Gudde
    Requires   : PowerShell 5.1+, Local Administrator rights
    Compatible : Windows Server 2012 R2, 2016, 2019, 2022
    Output     : WDS_Health_<hostname>_<date>.html  (same directory as script)
#>

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
$CompanyLogoURL      = ''
$CompanyWebsite      = 'https://tushargudde.tech'
$AuthorName          = 'Tushar Gudde'

$DiskWarnGB          = 20          # GB: warn threshold for RemoteInstall volume
$DiskCritGB          = 10          # GB: critical threshold

$ImageAgeDaysWarn    = 180         # flag boot images older than this many days

$PendingDeviceAlert  = $true       # show red tile when pending devices > 0

$EnableStatusFile    = $true       # write a plain-text .status companion file
# ──────────────────────────────────────────────────────────────────────────────

$ScriptVersion  = '1.0.0'
$ScriptStart    = Get-Date
$ServerHostname = $env:COMPUTERNAME
$ReportDate     = $ScriptStart.ToString('yyyy-MM-dd HH:mm:ss')

$ScriptDir  = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
$Stamp      = $ScriptStart.ToString('yyyyMMdd_HHmmss')
$ReportFile = Join-Path $ScriptDir "WDS_Health_${ServerHostname}_${Stamp}.html"
$StatusFile = Join-Path $ScriptDir "WDS_Health_${ServerHostname}_${Stamp}.status"

$CriticalFindings = [System.Collections.Generic.List[string]]::new()

function Add-Critical { param([string]$msg) $script:CriticalFindings.Add($msg) }

Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "================================================================" -ForegroundColor DarkCyan
Write-Host "  WDS Health Check v$ScriptVersion   Server: $ServerHostname" -ForegroundColor Cyan
Write-Host "  Started : $ReportDate" -ForegroundColor Gray
Write-Host "================================================================" -ForegroundColor DarkCyan
Write-Host ""

# ── COLOUR HELPERS ────────────────────────────────────────────────────────────
function Get-StatusBadge {
    param([string]$text, [string]$colour)
    $bg = switch ($colour) {
        'green'  { '#22c55e' }
        'red'    { '#ef4444' }
        'yellow' { '#eab308' }
        'orange' { '#f97316' }
        'blue'   { '#6366f1' }
        default  { '#6b7280' }
    }
    return "<span style='background:$bg;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.75rem;font-weight:700;'>$text</span>"
}

function Get-ArchBadge {
    param([string]$arch)
    $colour = switch ($arch) {
        'x64'   { '#4f46e5' }
        'x86'   { '#0891b2' }
        'arm64' { '#7c3aed' }
        default { '#6b7280' }
    }
    return "<span style='background:$colour;color:#fff;padding:2px 7px;border-radius:4px;font-size:0.72rem;font-weight:700;'>$arch</span>"
}

# ── HTML HEAD & THEME ─────────────────────────────────────────────────────────
$HtmlHead = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>WDS Health Check - $ServerHostname - $ReportDate</title>
<style>
  :root{
    --bg:#0f0f1a;--surface:#1a1a2e;--surface2:#16213e;--border:#2a2a4a;
    --accent:#6366f1;--accent2:#8b5cf6;--accent3:#a78bfa;
    --text:#e2e8f0;--muted:#94a3b8;--green:#22c55e;--red:#ef4444;
    --yellow:#eab308;--orange:#f97316;--blue:#38bdf8;
  }
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:var(--bg);color:var(--text);font-family:'Segoe UI',Arial,sans-serif;font-size:14px;line-height:1.5;}
  a{color:var(--accent3);text-decoration:none;}
  a:hover{text-decoration:underline;}
  .wrap{max-width:1400px;margin:0 auto;padding:24px 16px;}
  /* header */
  .header{background:linear-gradient(135deg,#1e1b4b 0%,#312e81 50%,#1e1b4b 100%);
    border-bottom:2px solid var(--accent);padding:28px 32px;border-radius:12px 12px 0 0;
    display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;}
  .header h1{font-size:1.7rem;font-weight:700;color:#fff;letter-spacing:.5px;}
  .header h1 span{color:var(--accent3);}
  .header .meta{font-size:.82rem;color:var(--muted);margin-top:4px;}
  .logo{height:48px;}
  /* summary tiles */
  .tiles{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:14px;margin:20px 0;}
  .tile{background:var(--surface);border:1px solid var(--border);border-radius:10px;
    padding:16px;text-align:center;}
  .tile .tv{font-size:2rem;font-weight:800;color:var(--accent3);}
  .tile .tl{font-size:.78rem;color:var(--muted);margin-top:4px;text-transform:uppercase;letter-spacing:.5px;}
  .tile.alert .tv{color:var(--red);}
  .tile.ok .tv{color:var(--green);}
  /* section */
  .section{background:var(--surface);border:1px solid var(--border);border-radius:10px;
    margin-bottom:20px;overflow:hidden;}
  .sec-head{background:linear-gradient(90deg,#1e1b4b,#2e1065);
    padding:12px 18px;display:flex;align-items:center;gap:10px;border-bottom:1px solid var(--border);}
  .sec-head h2{font-size:1rem;font-weight:700;color:var(--accent3);letter-spacing:.3px;}
  .sec-body{padding:18px;}
  /* table */
  table{width:100%;border-collapse:collapse;font-size:.84rem;}
  th{background:#1e1b4b;color:var(--accent3);padding:8px 10px;text-align:left;
    border-bottom:2px solid var(--border);font-weight:600;white-space:nowrap;}
  td{padding:7px 10px;border-bottom:1px solid var(--border);vertical-align:top;}
  tr:last-child td{border-bottom:none;}
  tr:hover td{background:rgba(99,102,241,.07);}
  /* kv */
  .kv{display:grid;grid-template-columns:220px 1fr;gap:4px 12px;}
  .kv .k{color:var(--muted);font-size:.82rem;}
  .kv .v{font-weight:600;}
  /* bar */
  .bar-wrap{background:#0d0d1f;border-radius:6px;height:16px;overflow:hidden;min-width:120px;}
  .bar-inner{height:100%;border-radius:6px;transition:width .3s;}
  /* badges */
  .badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:.72rem;
    font-weight:700;white-space:nowrap;}
  .badge-green{background:#14532d;color:#86efac;}
  .badge-red{background:#450a0a;color:#fca5a5;}
  .badge-yellow{background:#422006;color:#fcd34d;}
  .badge-blue{background:#1e3a5f;color:#7dd3fc;}
  .badge-gray{background:#1e293b;color:#94a3b8;}
  .badge-purple{background:#2e1065;color:#c4b5fd;}
  /* timeline row colours */
  .dep-ok td{background:rgba(34,197,94,.07);}
  .dep-fail td{background:rgba(239,68,68,.07);}
  .dep-warn td{background:rgba(234,179,8,.07);}
  /* info box */
  .info-box{background:#0d0d1f;border-left:4px solid var(--accent);
    padding:10px 14px;border-radius:0 6px 6px 0;font-size:.84rem;color:var(--muted);}
  /* disk bar */
  .disk-bar-wrap{background:#0d0d1f;border-radius:8px;height:22px;overflow:hidden;width:100%;min-width:180px;}
  .disk-bar-inner{height:100%;border-radius:8px;display:flex;align-items:center;
    padding-left:8px;font-size:.72rem;font-weight:700;color:#fff;}
  /* footer */
  .footer{text-align:center;color:var(--muted);font-size:.78rem;padding:18px;border-top:1px solid var(--border);}
  .not-installed{text-align:center;padding:40px;color:var(--muted);font-size:1rem;}
  .not-installed .ni-icon{font-size:3rem;margin-bottom:12px;}
  /* service matrix */
  .svc-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:12px;}
  .svc-card{background:#0d0d1f;border:1px solid var(--border);border-radius:8px;padding:14px;}
  .svc-card .sname{font-weight:700;font-size:.9rem;margin-bottom:6px;}
  .svc-card .sstate{font-size:.82rem;}
  pre{background:#0a0a14;border:1px solid var(--border);border-radius:6px;padding:12px;
    overflow-x:auto;font-size:.78rem;color:#a5f3fc;white-space:pre-wrap;word-break:break-all;}
</style>
</head>
<body>
<div class="wrap">
"@

# ── HTML HEADER ───────────────────────────────────────────────────────────────
$LogoHtml = ''
if ($CompanyLogoURL) {
    $LogoHtml = "<a href='$CompanyWebsite' target='_blank'><img class='logo' src='$CompanyLogoURL' alt='Logo'/></a>"
}

$HtmlHeader = @"
<div class="header">
  <div>
    <h1>WDS <span>Health Check</span></h1>
    <div class="meta">Server: <b>$ServerHostname</b> &nbsp;|&nbsp; Generated: $ReportDate &nbsp;|&nbsp; v$ScriptVersion</div>
  </div>
  $LogoHtml
</div>
"@

# ══════════════════════════════════════════════════════════════════════════════
#  PRE-CHECK: Is WDS Role installed?
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [Pre ] Checking WDS role installation..." -ForegroundColor Gray

$WdsRoleInstalled = $false
$WdsRoleHtml      = ''

try {
    $wdsFeature = Get-WindowsFeature -Name 'WDS' -ErrorAction Stop
    if ($wdsFeature.InstallState -eq 'Installed') {
        $WdsRoleInstalled = $true
        Write-Host "         WDS role is installed." -ForegroundColor Green
    } else {
        Write-Host "         WDS role is NOT installed." -ForegroundColor Yellow
        Add-Critical "WDS role is not installed on $ServerHostname"
        $WdsRoleHtml = @"
<div class="section">
  <div class="sec-head"><h2>Windows Deployment Services</h2></div>
  <div class="sec-body">
    <div class="not-installed">
      <div class="ni-icon">&#128268;</div>
      <p>The <b>Windows Deployment Services</b> role is <b>not installed</b> on this server.</p>
      <p style="margin-top:8px;font-size:.82rem;">Install it with: <code>Install-WindowsFeature WDS -IncludeManagementTools</code></p>
    </div>
  </div>
</div>
"@
    }
} catch {
    Write-Warning "Could not query WDS role: $_"
    $WdsRoleHtml = @"
<div class="section">
  <div class="sec-head"><h2>Windows Deployment Services</h2></div>
  <div class="sec-body">
    <div class="not-installed">
      <p>Could not determine WDS role installation status: $($_.Exception.Message)</p>
    </div>
  </div>
</div>
"@
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 1 - WDS Service Status
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [1/15] WDS Service Status..." -ForegroundColor Gray
$Sec1Html = ''
try {
    $wdsSvc = Get-Service -Name 'WDSServer' -ErrorAction SilentlyContinue
    if (-not $wdsSvc) {
        $Sec1Html = "<div class='info-box'>WDSServer service not found on this system.</div>"
        Add-Critical "WDSServer service not found"
    } else {
        $svcStatus = $wdsSvc.Status
        $svcStart  = $wdsSvc.StartType
        $stateColour = if ($svcStatus -eq 'Running') { 'green' } else { 'red' }
        if ($svcStatus -ne 'Running') { Add-Critical "WDSServer service is $svcStatus" }

        # Try WMI for PID and last start time
        $wmiSvc = $null
        try { $wmiSvc = Get-WmiObject -Class Win32_Service -Filter "Name='WDSServer'" -ErrorAction Stop } catch {}

        $pid     = if ($wmiSvc) { $wmiSvc.ProcessId } else { 'N/A' }
        $started = 'N/A'
        if ($pid -and $pid -ne 'N/A' -and $pid -gt 0) {
            try {
                $proc = Get-WmiObject -Class Win32_Process -Filter "ProcessId=$pid" -ErrorAction Stop
                if ($proc) {
                    $started = [System.Management.ManagementDateTimeConverter]::ToDateTime($proc.CreationDate).ToString('yyyy-MM-dd HH:mm:ss')
                }
            } catch {}
        }

        $Sec1Html = @"
<div class="kv">
  <span class="k">Service Name</span><span class="v">WDSServer</span>
  <span class="k">Status</span><span class="v">$(Get-StatusBadge $svcStatus $stateColour)</span>
  <span class="k">Start Type</span><span class="v">$svcStart</span>
  <span class="k">Process ID</span><span class="v">$pid</span>
  <span class="k">Last Start Time</span><span class="v">$started</span>
</div>
"@
    }
} catch {
    $Sec1Html = "<div class='info-box'>Error retrieving WDS service status: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 2 - WDS Configuration
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [2/15] WDS Configuration..." -ForegroundColor Gray
$Sec2Html = ''
try {
    $wdsInit      = 'Unknown'
    $wdsPath      = 'N/A'
    $wdsMode      = 'N/A'
    $wdsDhcpAuth  = 'N/A'
    $wdsAnswer    = 'N/A'

    # Try WMI namespace
    try {
        $wdsConf = Get-WmiObject -Namespace 'root\CIMV2\WDSServer' -Class 'WDSConfig' -ErrorAction Stop
        if ($wdsConf) {
            $wdsInit     = if ($wdsConf.Initialized) { 'Yes' } else { 'No' }
            $wdsPath     = $wdsConf.RemoteInstallPath
            $wdsMode     = if ($wdsConf.Standalone) { 'Standalone' } else { 'AD-Integrated' }
            $wdsDhcpAuth = if ($wdsConf.DhcpAuthorized) { 'Yes' } else { 'No' }
            $wdsAnswer   = switch ($wdsConf.AnswerPolicy) {
                0 { 'Not Responding' }
                1 { 'Known Clients Only' }
                2 { 'All Clients' }
                default { "Unknown ($($wdsConf.AnswerPolicy))" }
            }
        }
    } catch {
        # Fall back to WDSUTIL.exe
        try {
            $wdsutilOut = ''
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName               = 'wdsutil.exe'
            $psi.Arguments              = '/Get-Server /Show:Config'
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError  = $true
            $psi.UseShellExecute        = $false
            $psi.CreateNoWindow         = $true
            $proc = [System.Diagnostics.Process]::Start($psi)
            $finished = $proc.WaitForExit(15000)
            if ($finished) { $wdsutilOut = $proc.StandardOutput.ReadToEnd() }
            else { $proc.Kill() }

            if ($wdsutilOut) {
                foreach ($line in ($wdsutilOut -split "`n")) {
                    if ($line -match 'Initialized\s*:\s*(.+)')        { $wdsInit    = $Matches[1].Trim() }
                    if ($line -match 'Remote Install Folder\s*:\s*(.+)') { $wdsPath = $Matches[1].Trim() }
                    if ($line -match 'Answer Requests\s*:\s*(.+)')    { $wdsAnswer  = $Matches[1].Trim() }
                    if ($line -match 'Authorized\s*:\s*(.+)')         { $wdsDhcpAuth = $Matches[1].Trim() }
                }
            }
        } catch {
            # WDSUTIL not available
        }
    }

    $initColour   = if ($wdsInit -eq 'Yes') { 'green' } else { 'red' }
    $authColour   = if ($wdsDhcpAuth -eq 'Yes') { 'green' } else { 'yellow' }

    $Sec2Html = @"
<div class="kv">
  <span class="k">Initialized</span><span class="v">$(Get-StatusBadge $wdsInit $initColour)</span>
  <span class="k">Remote Install Path</span><span class="v">$wdsPath</span>
  <span class="k">Mode</span><span class="v">$wdsMode</span>
  <span class="k">DHCP Authorized</span><span class="v">$(Get-StatusBadge $wdsDhcpAuth $authColour)</span>
  <span class="k">Answer Policy</span><span class="v">$wdsAnswer</span>
</div>
"@
    if ($wdsInit -ne 'Yes') { Add-Critical "WDS is not initialized" }
} catch {
    $Sec2Html = "<div class='info-box'>Error retrieving WDS configuration: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 3 - PXE & TFTP Health
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [3/15] PXE and TFTP Health..." -ForegroundColor Gray
$Sec3Html = ''
try {
    # Check if TFTP port 69 is listed via netstat (UDP limitation workaround)
    $tftpActive = $false
    try {
        $netstatOut = & netstat -an 2>&1
        foreach ($line in $netstatOut) {
            if ($line -match '0\.0\.0\.0:69(\s|$)' -or $line -match '\*:69(\s|$)') {
                $tftpActive = $true; break
            }
        }
    } catch {}

    $tftpColour  = if ($tftpActive) { 'green' } else { 'yellow' }
    $tftpLabel   = if ($tftpActive) { 'Port 69 UDP Detected' } else { 'Port 69 Not Detected (UDP limitation)' }

    $Sec3Html = @"
<div class="kv">
  <span class="k">TFTP Port 69 (UDP)</span><span class="v">$(Get-StatusBadge $tftpLabel $tftpColour)</span>
  <span class="k">PXE Boot Policy</span><span class="v">Configured via WDS Answer Policy (see Section 2)</span>
  <span class="k">DHCP Option 60</span><span class="v">Should be set to PXEClient on DHCP server if not co-hosted</span>
  <span class="k">DHCP Option 66</span><span class="v">Boot server hostname - should point to this WDS server</span>
  <span class="k">DHCP Option 67</span><span class="v">Boot filename - e.g. boot\x64\wdsnbp.com</span>
</div>
<div class="info-box" style="margin-top:12px;">
  <b>Note:</b> UDP port 69 cannot be tested with a TCP connection test.
  The check above inspects active netstat listeners.
  If WDS and DHCP are co-hosted, DHCP option 60 should NOT be set (WDS handles PXE responses internally).
</div>
"@
} catch {
    $Sec3Html = "<div class='info-box'>Error checking PXE/TFTP: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 4 - Boot Images
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [4/15] Boot Images..." -ForegroundColor Gray
$Sec4Html = ''
try {
    $bootImages = $null
    $biRows     = ''

    try {
        Import-Module WDS -ErrorAction Stop
        $bootImages = Get-WdsBootImage -ErrorAction Stop
    } catch {
        # WDS module unavailable - try WDSUTIL
        try {
            $psi2 = New-Object System.Diagnostics.ProcessStartInfo
            $psi2.FileName               = 'wdsutil.exe'
            $psi2.Arguments              = '/Get-AllImages /Show:Boot'
            $psi2.RedirectStandardOutput = $true
            $psi2.RedirectStandardError  = $true
            $psi2.UseShellExecute        = $false
            $psi2.CreateNoWindow         = $true
            $proc2 = [System.Diagnostics.Process]::Start($psi2)
            $fin2  = $proc2.WaitForExit(20000)
            if (-not $fin2) { $proc2.Kill() }
            # Output parsing is complex; fall through to empty result
        } catch {}
    }

    $ageLimit = (Get-Date).AddDays(-$ImageAgeDaysWarn)

    if ($bootImages) {
        foreach ($img in $bootImages) {
            $arch        = if ($img.Architecture) { $img.Architecture.ToString().ToLower() } else { 'unknown' }
            $archBadge   = Get-ArchBadge $arch
            $name        = [System.Web.HttpUtility]::HtmlEncode($img.ImageName)
            $ver         = if ($img.Version)  { $img.Version }  else { 'N/A' }
            $winpe       = if ($img.WinPEVersion) { $img.WinPEVersion } else { 'N/A' }
            $sizeMB      = if ($img.FileSize) { [math]::Round($img.FileSize / 1MB, 1) } else { 'N/A' }
            $created     = 'N/A'
            $ageWarning  = ''
            if ($img.CreationTime) {
                $created = $img.CreationTime.ToString('yyyy-MM-dd')
                if ($img.CreationTime -lt $ageLimit) {
                    $ageWarning = " <span class='badge badge-yellow'>Old (&gt;${ImageAgeDaysWarn}d)</span>"
                    Add-Critical "Boot image '$($img.ImageName)' is older than $ImageAgeDaysWarn days"
                }
            }
            $biRows += "<tr><td>$name $ageWarning</td><td>$archBadge</td><td>$ver</td><td>$winpe</td><td>$sizeMB MB</td><td>$created</td></tr>"
        }
    }

    if ($biRows) {
        $Sec4Html = @"
<table>
<thead><tr>
  <th>Image Name</th><th>Architecture</th><th>Version</th><th>WinPE Version</th><th>Size</th><th>Created</th>
</tr></thead>
<tbody>$biRows</tbody>
</table>
"@
    } else {
        $Sec4Html = "<div class='info-box'>No boot images found or WDS module unavailable. Ensure the WDS module is loaded and images are imported.</div>"
    }
} catch {
    $Sec4Html = "<div class='info-box'>Error retrieving boot images: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 5 - Install Images
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [5/15] Install Images..." -ForegroundColor Gray
$Sec5Html = ''
try {
    $installImages = $null
    $iiRows        = ''

    try {
        Import-Module WDS -ErrorAction Stop
        $installImages = Get-WdsInstallImage -ErrorAction Stop
    } catch {}

    if ($installImages) {
        foreach ($img in $installImages) {
            $arch      = if ($img.Architecture) { $img.Architecture.ToString().ToLower() } else { 'unknown' }
            $archBadge = Get-ArchBadge $arch
            $name      = [System.Web.HttpUtility]::HtmlEncode($img.ImageName)
            $grp       = if ($img.ImageGroup) { $img.ImageGroup } else { 'N/A' }
            $sizeMB    = if ($img.FileSize)   { [math]::Round($img.FileSize / 1MB, 1) } else { 'N/A' }
            $created   = if ($img.CreationTime) { $img.CreationTime.ToString('yyyy-MM-dd') } else { 'N/A' }
            $iiRows   += "<tr><td>$name</td><td>$archBadge</td><td>$grp</td><td>$sizeMB MB</td><td>$created</td></tr>"
        }
    }

    if ($iiRows) {
        $Sec5Html = @"
<table>
<thead><tr>
  <th>OS Name</th><th>Architecture</th><th>Image Group</th><th>Size</th><th>Created</th>
</tr></thead>
<tbody>$iiRows</tbody>
</table>
"@
    } else {
        $Sec5Html = "<div class='info-box'>No install images found or WDS module unavailable.</div>"
    }
} catch {
    $Sec5Html = "<div class='info-box'>Error retrieving install images: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 6 - Multicast Sessions
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [6/15] Multicast Sessions..." -ForegroundColor Gray
$Sec6Html = ''
try {
    $mcSessions = $null
    $mcRows     = ''

    try {
        Import-Module WDS -ErrorAction Stop
        $mcSessions = Get-WdsMulticastTransmission -ErrorAction Stop
    } catch {}

    if ($mcSessions) {
        foreach ($s in $mcSessions) {
            $mcName    = [System.Web.HttpUtility]::HtmlEncode($s.Name)
            $clients   = if ($null -ne $s.ClientCount)    { $s.ClientCount }    else { 'N/A' }
            $progress  = if ($null -ne $s.PercentComplete) { "$($s.PercentComplete)%" } else { 'N/A' }
            $rate      = if ($null -ne $s.TransferRate)    { "$([math]::Round($s.TransferRate/1KB,1)) KB/s" } else { 'N/A' }
            $mcRows   += "<tr><td>$mcName</td><td>$clients</td><td>$progress</td><td>$rate</td></tr>"
        }
    }

    if ($mcRows) {
        $Sec6Html = @"
<table>
<thead><tr><th>Session Name</th><th>Clients</th><th>Progress</th><th>Transfer Rate</th></tr></thead>
<tbody>$mcRows</tbody>
</table>
"@
    } else {
        $Sec6Html = "<div class='info-box'>No active multicast transmissions found.</div>"
    }
} catch {
    $Sec6Html = "<div class='info-box'>Error retrieving multicast sessions: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 7 - Recent Deployment Activity (Log Files)
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [7/15] Recent Deployment Activity..." -ForegroundColor Gray
$Sec7Html = ''
try {
    $wdsLogDir = Join-Path $env:WINDIR 'System32\LogFiles\WDS'
    $depRows   = ''
    $logCount  = 0

    if (Test-Path $wdsLogDir) {
        $logFiles = Get-ChildItem -Path $wdsLogDir -Filter '*.log' -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 3

        $entries = [System.Collections.Generic.List[PSObject]]::new()

        foreach ($lf in $logFiles) {
            try {
                $lines = [System.IO.File]::ReadAllLines($lf.FullName)
                foreach ($line in $lines) {
                    if ($line -match '^\d{4}') {
                        $entries.Add([PSCustomObject]@{ Line = $line; File = $lf.Name })
                    }
                }
            } catch {}
        }

        # Take last 20
        $last20 = if ($entries.Count -gt 20) { $entries | Select-Object -Last 20 } else { $entries }

        foreach ($e in $last20) {
            $logCount++
            $l    = $e.Line
            $ts   = if ($l.Length -gt 20) { $l.Substring(0,20) } else { $l }
            $rest = if ($l.Length -gt 20) { $l.Substring(20) }   else { '' }

            $status   = 'Info'
            $rowClass = ''
            if ($rest -match '(?i)(fail|error|abort|denied)') {
                $status = 'Failed'; $rowClass = 'dep-fail'
                Add-Critical "WDS deployment failure in log: $ts"
            } elseif ($rest -match '(?i)(success|complete|done)') {
                $status = 'Success'; $rowClass = 'dep-ok'
            } elseif ($rest -match '(?i)(abort|cancel)') {
                $status = 'Aborted'; $rowClass = 'dep-warn'
            }

            $mac   = if ($rest -match '([0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2}') { $Matches[0] } else { 'N/A' }
            $image = if ($rest -match '(?i)image[:\s]+([^\s,]+)') { $Matches[1] } else { 'N/A' }
            $safeRest = [System.Web.HttpUtility]::HtmlEncode($rest.Substring(0, [math]::Min(80, $rest.Length)))

            $depRows += "<tr class='$rowClass'><td>$ts</td><td>$mac</td><td>$status</td><td>$image</td><td>$safeRest</td></tr>"
        }
    }

    if ($depRows) {
        $Sec7Html = @"
<table>
<thead><tr><th>Timestamp</th><th>Client MAC</th><th>Status</th><th>Image</th><th>Detail</th></tr></thead>
<tbody>$depRows</tbody>
</table>
"@
    } else {
        $Sec7Html = "<div class='info-box'>No WDS deployment log entries found in: $wdsLogDir</div>"
    }
} catch {
    $Sec7Html = "<div class='info-box'>Error reading deployment logs: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 8 - Pending Devices
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [8/15] Pending Devices..." -ForegroundColor Gray
$Sec8Html         = ''
$PendingCount     = 0
try {
    $pendingDevices = $null
    $pdRows         = ''

    try {
        Import-Module WDS -ErrorAction Stop
        $pendingDevices = Get-WdsPendingDevice -ErrorAction Stop
        $PendingCount   = @($pendingDevices).Count
    } catch {}

    if ($pendingDevices -and $PendingCount -gt 0) {
        Add-Critical "$PendingCount device(s) are in Pending state awaiting WDS approval"
        foreach ($pd in $pendingDevices) {
            $pdName = [System.Web.HttpUtility]::HtmlEncode($pd.Name)
            $pdMac  = $pd.MacAddress
            $pdTime = if ($pd.RequestTime) { $pd.RequestTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
            $pdRows += "<tr class='dep-fail'><td>$pdName</td><td>$pdMac</td><td>$pdTime</td></tr>"
        }
        $Sec8Html = @"
<table>
<thead><tr><th>Device Name</th><th>MAC Address</th><th>Request Time</th></tr></thead>
<tbody>$pdRows</tbody>
</table>
"@
    } else {
        $Sec8Html = "<div class='info-box' style='border-left-color:#22c55e;'>No pending devices awaiting approval.</div>"
    }
} catch {
    $Sec8Html = "<div class='info-box'>Error retrieving pending devices: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 9 - WDS DHCP Integration Check
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [9/15] WDS DHCP Integration..." -ForegroundColor Gray
$Sec9Html = ''
try {
    $dhcpOpt60     = 'Not Found / Not Checked'
    $dhcpOpt60Warn = $false
    $dhcpStatus    = ''

    # Check if DHCP module is present
    if (Get-Module -ListAvailable -Name 'DhcpServer' -ErrorAction SilentlyContinue) {
        try {
            Import-Module DhcpServer -ErrorAction Stop
            $srv = Get-DhcpServerSetting -ErrorAction Stop
            $opt = Get-DhcpServerv4OptionValue -OptionId 60 -ErrorAction SilentlyContinue
            if ($opt) {
                $dhcpOpt60     = $opt.Value -join ', '
                $dhcpOpt60Warn = ($dhcpOpt60 -match 'PXEClient')
                if ($dhcpOpt60Warn) {
                    $dhcpStatus = "WARNING: DHCP option 60 is set to PXEClient - this can conflict with WDS when co-hosted"
                    Add-Critical "DHCP option 60 = PXEClient conflicts with WDS co-hosting"
                }
            } else {
                $dhcpOpt60 = 'Not Set (correct for WDS co-hosted on same server)'
            }
        } catch {
            $dhcpOpt60 = "Could not read ($($_.Exception.Message))"
        }
    } else {
        $dhcpOpt60 = 'DHCP Server module not available'
    }

    $opt60Colour = if ($dhcpOpt60Warn) { 'red' } else { 'green' }
    $opt60Label  = if ($dhcpOpt60Warn) { 'Conflict Detected' } else { 'OK' }

    $Sec9Html = @"
<div class="kv">
  <span class="k">DHCP Option 60 Value</span><span class="v">$dhcpOpt60</span>
  <span class="k">Conflict Status</span><span class="v">$(Get-StatusBadge $opt60Label $opt60Colour)</span>
</div>
$(if ($dhcpStatus) { "<div class='info-box' style='margin-top:10px;border-left-color:#ef4444;'>$dhcpStatus</div>" })
<div class="info-box" style="margin-top:10px;">
  <b>Co-hosting guidance:</b> When WDS and DHCP run on the same server, do NOT set DHCP option 60.
  WDS listens directly on port 67 and intercepts PXE requests.
  When on separate servers: set option 66 (WDS server IP) and option 67 (boot filename) on the DHCP server.
</div>
"@
} catch {
    $Sec9Html = "<div class='info-box'>Error checking DHCP integration: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 10 - Disk Space (RemoteInstall Volume)
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [10/15] Disk Space..." -ForegroundColor Gray
$Sec10Html = ''
try {
    $diskRows = ''

    # Determine RemoteInstall path drive
    $riPath = $null
    try {
        $wdsConf2 = Get-WmiObject -Namespace 'root\CIMV2\WDSServer' -Class 'WDSConfig' -ErrorAction Stop
        if ($wdsConf2 -and $wdsConf2.RemoteInstallPath) { $riPath = $wdsConf2.RemoteInstallPath }
    } catch {}

    $volumes = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop

    foreach ($vol in $volumes) {
        $letter   = $vol.DeviceID
        $totalGB  = [math]::Round($vol.Size / 1GB, 1)
        $freeGB   = [math]::Round($vol.FreeSpace / 1GB, 1)
        $usedGB   = $totalGB - $freeGB
        $usedPct  = if ($totalGB -gt 0) { [math]::Round(($usedGB / $totalGB) * 100, 1) } else { 0 }

        $isRiDrive = ($riPath -and $riPath.StartsWith($letter))
        $label     = if ($isRiDrive) { "$letter <span class='badge badge-purple'>RemoteInstall</span>" } else { $letter }

        $barColour = '#6366f1'
        $warn      = ''
        if ($isRiDrive) {
            if ($freeGB -lt $DiskCritGB) {
                $barColour = '#ef4444'
                $warn = "<span class='badge badge-red'>CRITICAL - Free: ${freeGB} GB</span>"
                Add-Critical "RemoteInstall drive $letter has only ${freeGB} GB free (critical threshold: $DiskCritGB GB)"
            } elseif ($freeGB -lt $DiskWarnGB) {
                $barColour = '#eab308'
                $warn = "<span class='badge badge-yellow'>WARNING - Free: ${freeGB} GB</span>"
                Add-Critical "RemoteInstall drive $letter has only ${freeGB} GB free (warn threshold: $DiskWarnGB GB)"
            }
        }

        $barHtml = "<div class='disk-bar-wrap'><div class='disk-bar-inner' style='width:${usedPct}%;background:$barColour;'>${usedPct}%</div></div>"
        $diskRows += "<tr><td>$label</td><td>${totalGB} GB</td><td>${freeGB} GB</td><td>${usedPct}%</td><td>$barHtml</td><td>$warn</td></tr>"
    }

    if ($diskRows) {
        $Sec10Html = @"
<table>
<thead><tr><th>Drive</th><th>Total</th><th>Free</th><th>Used %</th><th>Usage Bar</th><th>Alert</th></tr></thead>
<tbody>$diskRows</tbody>
</table>
"@
    } else {
        $Sec10Html = "<div class='info-box'>No fixed disk volumes found.</div>"
    }
} catch {
    $Sec10Html = "<div class='info-box'>Error retrieving disk space: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 11 - WDS-related Services
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [11/15] WDS-related Services..." -ForegroundColor Gray
$Sec11Html = ''
try {
    $svcNames = @('WDSServer','WDSTFTP','BINLSVC','WDSTransportServer')
    $svcCards = ''

    foreach ($sn in $svcNames) {
        $svc = Get-Service -Name $sn -ErrorAction SilentlyContinue
        if ($svc) {
            $st       = $svc.Status
            $stType   = $svc.StartType
            $stColour = if ($st -eq 'Running') { '#22c55e' } else { '#ef4444' }
            if ($st -ne 'Running' -and $stType -ne 'Disabled') {
                Add-Critical "WDS-related service '$sn' is $st"
            }
            $svcCards += @"
<div class="svc-card">
  <div class="sname">$sn</div>
  <div class="sstate" style="color:$stColour;font-weight:700;">$st</div>
  <div class="sstate" style="margin-top:4px;">Start type: $stType</div>
</div>
"@
        } else {
            $svcCards += @"
<div class="svc-card">
  <div class="sname">$sn</div>
  <div class="sstate" style="color:#6b7280;">Not Found</div>
</div>
"@
        }
    }

    $Sec11Html = "<div class='svc-grid'>$svcCards</div>"
} catch {
    $Sec11Html = "<div class='info-box'>Error checking WDS services: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 12 - Active Directory Integration
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [12/15] Active Directory Integration..." -ForegroundColor Gray
$Sec12Html = ''
try {
    $adAccountExists = 'Unknown'
    $dhcpAuthorized  = 'Unknown'
    $adColour        = 'blue'
    $authColour2     = 'blue'

    # Check computer account in AD
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $adObj = Get-ADComputer -Identity $ServerHostname -ErrorAction Stop
        if ($adObj) {
            $adAccountExists = 'Found'
            $adColour        = 'green'
        }
    } catch {
        $adAccountExists = "Not Found / Error ($($_.Exception.Message))"
        $adColour        = 'yellow'
    }

    # Check DHCP authorization
    try {
        Import-Module DhcpServer -ErrorAction Stop
        $authServers = Get-DhcpServerInDC -ErrorAction Stop
        $matched     = $authServers | Where-Object {
            $_.DnsName -eq $ServerHostname -or
            $_.DnsName -like "$ServerHostname.*" -or
            $_.IPAddress -eq $ServerHostname
        }
        if ($matched) {
            $dhcpAuthorized = 'Authorized'
            $authColour2    = 'green'
        } else {
            $dhcpAuthorized = 'Not Authorized in DHCP'
            $authColour2    = 'yellow'
            Add-Critical "WDS server is not authorized in DHCP Active Directory"
        }
    } catch {
        $dhcpAuthorized = "Could not check ($($_.Exception.Message))"
    }

    $Sec12Html = @"
<div class="kv">
  <span class="k">AD Computer Account</span><span class="v">$(Get-StatusBadge $adAccountExists $adColour)</span>
  <span class="k">DHCP Authorization (AD)</span><span class="v">$(Get-StatusBadge $dhcpAuthorized $authColour2)</span>
</div>
"@
} catch {
    $Sec12Html = "<div class='info-box'>Error checking AD integration: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 13 - Network Check (DC reachability)
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [13/15] Network Check..." -ForegroundColor Gray
$Sec13Html = ''
try {
    $dcFqdn       = $null
    $dnsOk        = $false
    $ldapOk       = $false
    $dnsResult    = 'Not Tested'
    $ldapResult   = 'Not Tested'

    # Resolve a DC from domain
    try {
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $dc     = $domain.FindDomainController()
        $dcFqdn = $dc.Name
    } catch {
        try {
            $dcFqdn = (Resolve-DnsName -Name "_ldap._tcp.dc._msdcs.$env:USERDNSDOMAIN" -Type SRV -ErrorAction Stop |
                       Where-Object { $_.Type -eq 'SRV' } | Select-Object -First 1).NameTarget
        } catch {}
    }

    if ($dcFqdn) {
        # DNS resolution test
        try {
            $resolved = [System.Net.Dns]::GetHostEntry($dcFqdn)
            if ($resolved) { $dnsOk = $true; $dnsResult = "Resolved: $($resolved.AddressList[0].ToString())" }
        } catch {
            $dnsResult = "Failed: $($_.Exception.Message)"
            Add-Critical "DNS resolution of DC '$dcFqdn' failed"
        }

        # LDAP port 389 test
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $ar  = $tcp.BeginConnect($dcFqdn, 389, $null, $null)
            $wait = $ar.AsyncWaitHandle.WaitOne(3000)
            if ($wait -and $tcp.Connected) {
                $ldapOk     = $true
                $ldapResult = 'Port 389 Open'
            } else {
                $ldapResult = 'Port 389 Unreachable (timeout)'
                Add-Critical "LDAP port 389 unreachable on DC '$dcFqdn'"
            }
            $tcp.Close()
        } catch {
            $ldapResult = "Test failed: $($_.Exception.Message)"
        }
    } else {
        $dnsResult  = 'Could not determine Domain Controller FQDN'
        $ldapResult = 'Skipped - no DC found'
    }

    $dnsColour  = if ($dnsOk)  { 'green' } else { 'red' }
    $ldapColour = if ($ldapOk) { 'green' } else { 'red' }

    $Sec13Html = @"
<div class="kv">
  <span class="k">Domain Controller</span><span class="v">$(if ($dcFqdn) { $dcFqdn } else { 'Not determined' })</span>
  <span class="k">DNS Resolution</span><span class="v">$(Get-StatusBadge $dnsResult $dnsColour)</span>
  <span class="k">LDAP (port 389)</span><span class="v">$(Get-StatusBadge $ldapResult $ldapColour)</span>
</div>
"@
} catch {
    $Sec13Html = "<div class='info-box'>Error performing network check: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 14 - WDS Event Log (last 20 events)
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [14/15] WDS Event Log..." -ForegroundColor Gray
$Sec14Html = ''
try {
    $wdsEvents = $null
    $evRows14  = ''

    try {
        $wdsEvents = Get-WinEvent -LogName 'Windows Deployment Services' -MaxEvents 20 -ErrorAction Stop
    } catch {
        # Log may not exist if WDS is not active
        try {
            $wdsEvents = Get-WinEvent -ProviderName 'Windows-Deployment-Services' -MaxEvents 20 -ErrorAction Stop
        } catch {}
    }

    if ($wdsEvents) {
        foreach ($ev in $wdsEvents) {
            $level = switch ($ev.Level) {
                1 { 'Critical' } 2 { 'Error' } 3 { 'Warning' } 4 { 'Information' } default { 'Verbose' }
            }
            $lColour = switch ($level) {
                'Critical'    { 'badge-red' }
                'Error'       { 'badge-red' }
                'Warning'     { 'badge-yellow' }
                'Information' { 'badge-blue' }
                default       { 'badge-gray' }
            }
            $ts  = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            $msg = [System.Web.HttpUtility]::HtmlEncode($ev.Message.Substring(0, [math]::Min(120, $ev.Message.Length)))
            $evRows14 += "<tr><td>$ts</td><td><span class='badge $lColour'>$level</span></td><td>$($ev.Id)</td><td>$msg</td></tr>"
        }
    }

    if ($evRows14) {
        $Sec14Html = @"
<table>
<thead><tr><th>Time</th><th>Level</th><th>Event ID</th><th>Message</th></tr></thead>
<tbody>$evRows14</tbody>
</table>
"@
    } else {
        $Sec14Html = "<div class='info-box'>No events found in the Windows Deployment Services event log.</div>"
    }
} catch {
    $Sec14Html = "<div class='info-box'>Error reading WDS event log: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 15 - System Event Log (WDS/TFTP/BINL mentions)
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "  [15/15] System Event Log..." -ForegroundColor Gray
$Sec15Html = ''
try {
    $sysEvents  = $null
    $evRows15   = ''
    $keywords   = @('WDS','TFTP','BINL','Multicast','RemoteInstall','WDSServer','WDSTFTP','BINLSVC')

    try {
        $sysEvents = Get-WinEvent -LogName System -MaxEvents 500 -ErrorAction Stop |
                     Where-Object {
                         $ev2 = $_
                         $ev2.Level -le 2 -and ($keywords | Where-Object { $ev2.Message -match $_ }).Count -gt 0
                     } |
                     Select-Object -First 10
    } catch {
        try {
            $sysEvents = Get-EventLog -LogName System -EntryType Error,Warning -Newest 500 -ErrorAction Stop |
                         Where-Object { $kw = $_.Message; ($keywords | Where-Object { $kw -match $_ }).Count -gt 0 } |
                         Select-Object -First 10
        } catch {}
    }

    if ($sysEvents) {
        foreach ($ev in $sysEvents) {
            $level = ''
            $lColour = 'badge-gray'
            if ($ev.PSObject.Properties['Level']) {
                $level = switch ($ev.Level) {
                    1 { 'Critical' } 2 { 'Error' } 3 { 'Warning' } default { 'Info' }
                }
                $lColour = switch ($level) { 'Critical' { 'badge-red' } 'Error' { 'badge-red' } 'Warning' { 'badge-yellow' } default { 'badge-blue' } }
                $ts = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
                $id = $ev.Id
                $msg = [System.Web.HttpUtility]::HtmlEncode($ev.Message.Substring(0, [math]::Min(120, $ev.Message.Length)))
            } else {
                # EventLog object
                $level   = $ev.EntryType.ToString()
                $lColour = switch ($level) { 'Error' { 'badge-red' } 'Warning' { 'badge-yellow' } default { 'badge-blue' } }
                $ts  = $ev.TimeGenerated.ToString('yyyy-MM-dd HH:mm:ss')
                $id  = $ev.EventID
                $msg = [System.Web.HttpUtility]::HtmlEncode($ev.Message.Substring(0, [math]::Min(120, $ev.Message.Length)))
            }
            $evRows15 += "<tr><td>$ts</td><td><span class='badge $lColour'>$level</span></td><td>$id</td><td>$msg</td></tr>"
        }
    }

    if ($evRows15) {
        $Sec15Html = @"
<table>
<thead><tr><th>Time</th><th>Level</th><th>Event ID</th><th>Message</th></tr></thead>
<tbody>$evRows15</tbody>
</table>
"@
    } else {
        $Sec15Html = "<div class='info-box'>No Critical/Error System events mentioning WDS, TFTP, BINL, or Multicast found.</div>"
    }
} catch {
    $Sec15Html = "<div class='info-box'>Error reading System event log: $($_.Exception.Message)</div>"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SUMMARY TILES
# ══════════════════════════════════════════════════════════════════════════════
$isCritical     = $CriticalFindings.Count -gt 0
$overallStatus  = if ($isCritical) { 'CRITICAL' } else { 'HEALTHY' }
$overallColour  = if ($isCritical) { '#ef4444' } else { '#22c55e' }

$pendingClass   = if ($PendingCount -gt 0) { 'tile alert' } else { 'tile ok' }

$SummaryTiles = @"
<div class="tiles">
  <div class="tile $(if ($isCritical) { 'alert' } else { 'ok' })">
    <div class="tv">$overallStatus</div>
    <div class="tl">Overall Status</div>
  </div>
  <div class="$pendingClass">
    <div class="tv">$PendingCount</div>
    <div class="tl">Pending Devices</div>
  </div>
  <div class="tile">
    <div class="tv">$($CriticalFindings.Count)</div>
    <div class="tl">Critical Findings</div>
  </div>
  <div class="tile">
    <div class="tv">$(if ($WdsRoleInstalled) { 'Installed' } else { 'Not Installed' })</div>
    <div class="tl">WDS Role</div>
  </div>
</div>
"@

# ══════════════════════════════════════════════════════════════════════════════
#  CRITICAL FINDINGS SECTION
# ══════════════════════════════════════════════════════════════════════════════
$CritSection = ''
if ($isCritical) {
    $cfRows = ($CriticalFindings | ForEach-Object { "<tr><td><span class='badge badge-red'>CRITICAL</span></td><td>$([System.Web.HttpUtility]::HtmlEncode($_))</td></tr>" }) -join ''
    $CritSection = @"
<div class="section" style="border-color:#ef4444;">
  <div class="sec-head" style="background:linear-gradient(90deg,#450a0a,#7f1d1d);">
    <h2 style="color:#fca5a5;">Critical Findings ($($CriticalFindings.Count))</h2>
  </div>
  <div class="sec-body">
    <table>
    <thead><tr><th>Severity</th><th>Finding</th></tr></thead>
    <tbody>$cfRows</tbody>
    </table>
  </div>
</div>
"@
}

# ══════════════════════════════════════════════════════════════════════════════
#  ASSEMBLE HTML
# ══════════════════════════════════════════════════════════════════════════════
$ScriptEnd  = Get-Date
$Duration   = ($ScriptEnd - $ScriptStart).ToString('mm\:ss')

function New-Section {
    param([string]$Icon, [string]$Title, [string]$Body)
    return @"
<div class="section">
  <div class="sec-head"><h2>$Icon $Title</h2></div>
  <div class="sec-body">$Body</div>
</div>
"@
}

$HtmlBody = $HtmlHead + $HtmlHeader + $SummaryTiles + $CritSection

if (-not $WdsRoleInstalled) {
    $HtmlBody += $WdsRoleHtml
} else {
    $HtmlBody += New-Section '&#128268;' '1. WDS Service Status'              $Sec1Html
    $HtmlBody += New-Section '&#9881;'   '2. WDS Configuration'               $Sec2Html
    $HtmlBody += New-Section '&#128225;' '3. PXE &amp; TFTP Health'           $Sec3Html
    $HtmlBody += New-Section '&#128190;' '4. Boot Images'                     $Sec4Html
    $HtmlBody += New-Section '&#128444;' '5. Install Images'                  $Sec5Html
    $HtmlBody += New-Section '&#127758;' '6. Multicast Sessions'              $Sec6Html
    $HtmlBody += New-Section '&#128203;' '7. Recent Deployment Activity'      $Sec7Html
    $HtmlBody += New-Section '&#9201;'   '8. Pending Devices'                 $Sec8Html
    $HtmlBody += New-Section '&#128279;' '9. WDS DHCP Integration'            $Sec9Html
    $HtmlBody += New-Section '&#128190;' '10. Disk Space'                     $Sec10Html
    $HtmlBody += New-Section '&#9881;'   '11. WDS-related Services'           $Sec11Html
    $HtmlBody += New-Section '&#128194;' '12. Active Directory Integration'   $Sec12Html
    $HtmlBody += New-Section '&#128246;' '13. Network Check'                  $Sec13Html
    $HtmlBody += New-Section '&#128203;' '14. WDS Event Log'                  $Sec14Html
    $HtmlBody += New-Section '&#128203;' '15. System Event Log (WDS)'         $Sec15Html
}

$HtmlBody += @"
<div class="footer">
  WDS Health Check v$ScriptVersion &nbsp;|&nbsp; Generated: $ReportDate &nbsp;|&nbsp; Duration: $Duration
  &nbsp;|&nbsp; Server: $ServerHostname &nbsp;|&nbsp;
  <a href="$CompanyWebsite" target="_blank">$AuthorName</a>
</div>
</div><!-- /wrap -->
</body>
</html>
"@

# ══════════════════════════════════════════════════════════════════════════════
#  WRITE REPORT
# ══════════════════════════════════════════════════════════════════════════════
try {
    [System.IO.File]::WriteAllText($ReportFile, $HtmlBody, [System.Text.Encoding]::UTF8)
    Write-Host ""
    Write-Host "  [OK] Report written: $ReportFile" -ForegroundColor Green
} catch {
    Write-Warning "Failed to write HTML report: $_"
}

# ══════════════════════════════════════════════════════════════════════════════
#  STATUS FILE
# ══════════════════════════════════════════════════════════════════════════════
$sep = '=' * 62
if ($isCritical) {
    $sc  = "$sep`r`n WDS HEALTH CHECK  -  *** CRITICAL ALERT ***`r`n$sep`r`n"
    $sc += " Status    : CRITICAL`r`n Server    : $ServerHostname`r`n"
    $sc += " Generated : $ReportDate`r`n Duration  : $Duration`r`n$sep`r`n"
    $sc += "`r`n CRITICAL FINDINGS ($($CriticalFindings.Count)):`r`n`r`n"
    $sc += (@($CriticalFindings) | ForEach-Object { "  [!] $_" }) -join "`r`n"
    $sc += "`r`n`r`n$sep`r`n Full HTML report : $ReportFile`r`n Status file      : $StatusFile`r`n$sep`r`n"
    $sc += " WDS Health Check v$ScriptVersion  by $AuthorName`r`n$sep`r`n"
} else {
    $sc  = "$sep`r`n WDS HEALTH CHECK  -  HEALTHY STATE`r`n$sep`r`n"
    $sc += " Status    : HEALTHY`r`n Server    : $ServerHostname`r`n"
    $sc += " Generated : $ReportDate`r`n Duration  : $Duration`r`n$sep`r`n"
    $sc += "`r`n No critical findings detected. WDS is in a healthy state.`r`n"
    $sc += "`r`n$sep`r`n Full HTML report : $ReportFile`r`n Status file      : $StatusFile`r`n$sep`r`n"
    $sc += " WDS Health Check v$ScriptVersion  by $AuthorName`r`n$sep`r`n"
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
Write-Host "================================================================" -ForegroundColor DarkCyan
Write-Host "  WDS Health Check complete.  Duration: $Duration" -ForegroundColor DarkCyan
Write-Host "  Report : $ReportFile" -ForegroundColor Yellow
if ($EnableStatusFile) {
    $slabel = if ($isCritical) { 'Status (CRITICAL)' } else { 'Status (HEALTHY)' }
    Write-Host "  $slabel : $StatusFile" -ForegroundColor $(if ($isCritical) { 'Red' } else { 'Green' })
    Write-Host ""
    Write-Host "  Review the HTML report and address any critical findings." -ForegroundColor Cyan
}
Write-Host "================================================================" -ForegroundColor DarkCyan
Write-Host ""
