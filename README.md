# Windows Server Health Check Scripts

A collection of PowerShell scripts that perform comprehensive health checks for various Windows Server roles and export the results to self-contained HTML reports with a professional dark/light-themed dashboard UI.

## Repository Structure

```
Windows-Server-Reports-scripts/
│
├── Start-HealthCheckMenu.cmd             # Interactive master menu launcher (Start Here!)
├── MasterMenu.ps1                        # Master menu PowerShell UI
├── .github/workflows/psscriptanalyzer.yml# CI/CD pipeline for code linting
│
├── AD_HealthCheck/
│   ├── Reports/                          # Auto-created; stores generated .html and .txt reports
│   ├── AD_HealthCheck.ps1                # Main health-check script
│   ├── AD_HealthCheck_EmailAlert.ps1     # Email alert companion
│   ├── Run_AD_HealthCheck.ps1            # Launcher (Task Scheduler)
│   └── Run_AD_HealthCheck.cmd            # CMD Wrapper (Bypasses execution policy)
│
├── ADCS_HealthCheck/
│   ├── Reports/
│   ├── ADCS_HealthCheck.ps1
│   ├── ADCS_HealthCheck_EmailAlert.ps1
│   ├── Run_ADCS_HealthCheck.ps1
│   └── Run_ADCS_HealthCheck.cmd
│
├── DHCP_DNS_HealthCheck/
│   ├── Reports/
│   ├── DHCP_DNS_HealthCheck.ps1
│   ├── DHCP_DNS_HealthCheck_EmailAlert.ps1
│   ├── Run_DHCP_DNS_HealthCheck.ps1
│   └── Run_DHCP_DNS_HealthCheck.cmd
│
├── ERP_HealthCheck/
│   ├── Reports/
│   ├── ERP_HealthCheck.ps1
│   ├── ERP_HealthCheck_EmailAlert.ps1
│   ├── Run_ERP_HealthCheck.ps1
│   └── Run_ERP_HealthCheck.cmd
│
├── Server_HealthCheck/
│   ├── Reports/
│   ├── Server_HealthCheck.ps1
│   ├── Server_HealthCheck_EmailAlert.ps1
│   ├── Run_Server_HealthCheck.ps1
│   └── Run_Server_HealthCheck.cmd
│
├── WDS_HealthCheck/
│   ├── Reports/
│   ├── WDS_HealthCheck.ps1
│   ├── WDS_HealthCheck_EmailAlert.ps1
│   ├── Run_WDS_HealthCheck.ps1
│   └── Run_WDS_HealthCheck.cmd
│
└── WSUS_HealthCheck/
    ├── Reports/
    ├── WSUS_HealthCheck.ps1
    ├── WSUS_HealthCheck_EmailAlert.ps1
    ├── Run_WSUS_HealthCheck.ps1
    └── Run_WSUS_HealthCheck.cmd
```

> **Note:** The `Reports/` folder inside each subfolder is automatically created on the first run of the health-check script. You do not need to create it manually.

## Features

- **Interactive Master Menu** — Launch any health check from a single, beautiful terminal dashboard (`Start-HealthCheckMenu.cmd`)
- **Air-Gapped Ready** — Fully self-contained HTML output and native CMD wrappers that bypass Execution Policy restrictions without internet.
- **Multiple Role Coverage** — AD, ADCS, DHCP/DNS, ERP, Server baseline, WDS, and WSUS
- **Dark/Light Mode Toggle** — switch themes with one click in any report
- **Email Alerts** for critical findings (configurable SMTP per script)
- **Event Log Analysis** with human-readable suggestions and impact ratings
- **Reports Folder** — all HTML reports saved to a dedicated `Reports\` subfolder inside each script directory
- **CI/CD Integration** — Includes GitHub Actions workflows for automated PowerShell linting.

## Prerequisites

| Requirement | Details |
|-------------|---------|
| OS | Windows Server 2016 / 2019 / 2022 / 2025 |
| PowerShell | 5.1 or later |
| Permissions | Domain Admin or equivalent (role-dependent) |
| Network | Access to target servers/services |

## Quick Start

**Option 1: The Interactive Master Menu (Recommended)**
Simply double-click **`Start-HealthCheckMenu.cmd`** in the root directory to open the interactive dashboard and choose which server role you want to audit.

**Option 2: Individual Wrappers**
If you want to run a specific check without the menu, navigate to its folder and double-click the `.cmd` wrapper:
```cmd
cd "C:\Path\To\Windows-Server-Reports-scripts\AD_HealthCheck"
Run_AD_HealthCheck.cmd
```
*Note: The `.cmd` wrappers automatically bypass PowerShell Execution Policy restrictions, making them perfect for locked-down environments.*

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