# Active Directory Health Check Dashboard

A comprehensive PowerShell script that performs a full Active Directory environment health check and exports the results to a single self-contained HTML file with a professional dark/light-themed dashboard UI.

## Features

- **14 Health Check Categories** covering Domain Info, DC Inventory, Services, Replication, DNS, FSMO, Trusts, Tombstone, Privileged Accounts, Password Policy, Stale Objects, Event Logs, and Windows Update Status
- **Dark/Light Mode Toggle** -- switch themes with one click
- **Email Alerts** for critical findings (configurable SMTP)
- **Event Log Analysis** with human-readable suggestions and impact ratings
- **Windows Update Status** per Domain Controller
- **Reports Folder** -- all HTML reports saved to a dedicated `Reports\` subfolder
- **Zero External Dependencies** -- fully self-contained HTML output

## Prerequisites

| Requirement | Details |
|-------------|---------|
| OS | Windows Server 2016 / 2019 / 2022 / 2025 |
| PowerShell | 5.1 or later |
| Modules | RSAT AD DS Tools (`ActiveDirectory` module) |
| Permissions | Domain Admin or equivalent |
| Network | Access to all DCs (WinRM, RPC, SMB) |

## Quick Start

```powershell
# 1. Open PowerShell as Administrator on a DC or RSAT workstation
# 2. Navigate to the script directory
cd "C:\Path\To\Windows-Server-Reports-scripts"

# 3. Run the script
.\AD_HealthCheck.ps1

# 4. Open the report
# Reports are saved to: .\Reports\AD_Health_YYYYMMDD_HHmmss.html
```

## Configuration Variables

Open `AD_HealthCheck.ps1` and edit the variables at the top of the script:

### Company Branding

| Variable | Default | Description |
|----------|---------|-------------|
| `$CompanyLogoURL` | `''` | URL to your company logo (PNG/SVG). Leave blank to skip. |
| `$CompanyWebsite` | `''` | Your company website URL for the logo hyperlink. |

### Author

| Variable | Default | Description |
|----------|---------|-------------|
| `$AuthorName` | `'Tushar Gudde'` | Author name shown in the report footer. |

### Event Log Settings

| Variable | Default | Description |
|----------|---------|-------------|
| `$EventLogHours` | `2` | How many hours back to scan for Warning/Error/Critical events. |

### Stale Object Threshold

| Variable | Default | Description |
|----------|---------|-------------|
| `$StaleThresholdDays` | `90` | Number of days of inactivity before an object is flagged stale. |

### Email Alert Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `$EnableEmailAlert` | `$false` | Set to `$true` to enable SMTP email alerts. |
| `$SMTPServer` | `'smtp.yourdomain.com'` | Your SMTP server hostname. |
| `$SMTPPort` | `587` | SMTP port number. |
| `$SMTPFrom` | `'ad-healthcheck@yourdomain.com'` | Sender email address. |
| `$SMTPTo` | `@('admin@yourdomain.com')` | Array of recipient email addresses. |
| `$SMTPSubject` | `'AD Health Check - CRITICAL ALERT'` | Email subject line. |
| `$SMTPUseSSL` | `$true` | Use SSL/TLS for SMTP connection. |
| `$SMTPCredentialUser` | `''` | SMTP username. Leave blank for anonymous relay. |
| `$SMTPCredentialPass` | `''` | SMTP password. Leave blank for anonymous relay. |

## Report Sections

| # | Section | Description |
|---|---------|-------------|
| 1 | Domain and Forest Info | Domain name, functional levels, FSMO role holders |
| 2 | Domain Controller Inventory | All DCs with IP, OS, Site, GC, RODC status |
| 3 | AD Services Status | NTDS, Netlogon, W32Time, DNS, KDC per DC |
| 4 | Replication Health | repadmin results with failure highlighting |
| 5 | SYSVOL and Netlogon Shares | Accessibility verification per DC |
| 6 | DNS Health | SRV records, forward/reverse zones, forwarders |
| 7 | FSMO Role Holders | All 5 roles with reachability check |
| 8 | AD Trust Relationships | Trust name, direction, type, status |
| 9 | Tombstone and Recycle Bin | Tombstone lifetime, Recycle Bin status |
| 10 | Privileged Account Audit | Domain/Enterprise/Schema Admins with stale flags |
| 11 | Default Domain Password Policy | Length, complexity, lockout settings |
| 12 | Stale Objects | Inactive computers and users (count only) |
| 13 | Directory Service Event Log | Last 2 hours of events with suggestions |
| 14 | Windows Update Status | Installed and pending updates per DC |

## Output

Reports are saved to `.\Reports\AD_Health_YYYYMMDD_HHmmss.html` and can be opened in any modern browser. The HTML file is fully self-contained with no external dependencies.

## License

This project is provided as-is for internal IT administration use.

---

**Created by Tushar Gudde**