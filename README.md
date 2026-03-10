# Windows Server Health Check Scripts

A collection of PowerShell scripts that perform comprehensive health checks for various Windows Server roles and export the results to self-contained HTML reports with a professional dark/light-themed dashboard UI.

## Repository Structure

```
Windows-Server-Reports-scripts/
│
├── AD_HealthCheck/
│   ├── Reports/                          # Auto-created; stores generated .html and .txt reports
│   ├── AD_HealthCheck.ps1                # Main health-check script
│   ├── AD_HealthCheck_EmailAlert.ps1     # Email alert companion
│   └── Run_AD_HealthCheck.ps1            # Launcher (double-click or Task Scheduler)
│
├── ADCS_HealthCheck/
│   ├── Reports/
│   ├── ADCS_HealthCheck.ps1
│   ├── ADCS_HealthCheck_EmailAlert.ps1
│   └── Run_ADCS_HealthCheck.ps1
│
├── DHCP_DNS_HealthCheck/
│   ├── Reports/
│   ├── DHCP_DNS_HealthCheck.ps1
│   ├── DHCP_DNS_HealthCheck_EmailAlert.ps1
│   └── Run_DHCP_DNS_HealthCheck.ps1
│
├── ERP_HealthCheck/
│   ├── Reports/
│   ├── ERP_HealthCheck.ps1
│   ├── ERP_HealthCheck_EmailAlert.ps1
│   └── Run_ERP_HealthCheck.ps1
│
├── Server_HealthCheck/
│   ├── Reports/
│   ├── Server_HealthCheck.ps1
│   ├── Server_HealthCheck_EmailAlert.ps1
│   └── Run_Server_HealthCheck.ps1
│
├── WDS_HealthCheck/
│   ├── Reports/
│   ├── WDS_HealthCheck.ps1
│   ├── WDS_HealthCheck_EmailAlert.ps1
│   └── Run_WDS_HealthCheck.ps1
│
└── WSUS_HealthCheck/
    ├── Reports/
    ├── WSUS_HealthCheck.ps1
    ├── WSUS_HealthCheck_EmailAlert.ps1
    └── Run_WSUS_HealthCheck.ps1
```

> **Note:** The `Reports/` folder inside each subfolder is automatically created on the first run of the health-check script. You do not need to create it manually.

## Features

- **Multiple Role Coverage** — AD, ADCS, DHCP/DNS, ERP, Server baseline, WDS, and WSUS
- **Dark/Light Mode Toggle** — switch themes with one click in any report
- **Email Alerts** for critical findings (configurable SMTP per script)
- **Event Log Analysis** with human-readable suggestions and impact ratings
- **Reports Folder** — all HTML reports saved to a dedicated `Reports\` subfolder inside each script directory
- **Zero External Dependencies** — fully self-contained HTML output

## Prerequisites

| Requirement | Details |
|-------------|---------|
| OS | Windows Server 2016 / 2019 / 2022 / 2025 |
| PowerShell | 5.1 or later |
| Permissions | Domain Admin or equivalent (role-dependent) |
| Network | Access to target servers/services |

## Quick Start

```powershell
# 1. Open PowerShell as Administrator
# 2. Navigate to the desired script subfolder, e.g.:
cd "C:\Path\To\Windows-Server-Reports-scripts\AD_HealthCheck"

# Option A – run directly
.\AD_HealthCheck.ps1

# Option B – use the launcher (works from Task Scheduler or Explorer double-click)
.\Run_AD_HealthCheck.ps1

# 3. Open the report
# Reports are saved to: .\Reports\AD_Health_YYYYMMDD_HHmmss.html
```

## Configuration Variables

Each health-check script has a configuration block near the top. Common variables:

### Company Branding

| Variable | Default | Description |
|----------|---------|-------------|
| `$CompanyLogoURL` | `''` | URL to your company logo (PNG/SVG). Leave blank to skip. |
| `$CompanyWebsite` | `''` | Your company website URL for the logo hyperlink. |
| `$AuthorName` | `'Tushar Gudde'` | Author name shown in the report footer. |

### Email Alert Configuration (in `*_EmailAlert.ps1`)

| Variable | Default | Description |
|----------|---------|-------------|
| `$SmtpServer` | `'smtp.yourdomain.com'` | Your SMTP server hostname. |
| `$SmtpPort` | `25` | SMTP port number. |
| `$MailFrom` | `'healthcheck@yourdomain.com'` | Sender email address. |
| `$MailTo` | `@('admin@yourdomain.com')` | Array of recipient email addresses. |
| `$AlertOnHealthy` | `$false` | Set to `$true` to also send confirmation emails for healthy runs. |

## Output

Reports are saved to `.\Reports\<ScriptName>_YYYYMMDD_HHmmss.html` inside each script's subfolder and can be opened in any modern browser. The HTML file is fully self-contained with no external dependencies.

## License

This project is provided as-is for internal IT administration use.

---

**Created by Tushar Gudde**