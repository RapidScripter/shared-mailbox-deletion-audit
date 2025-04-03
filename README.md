# Exchange Online Shared Mailbox Deletion Audit

![PowerShell](https://img.shields.io/badge/PowerShell-%235391FE.svg?style=for-the-badge&logo=powershell&logoColor=white)
![Microsoft Exchange](https://img.shields.io/badge/Microsoft_Exchange-0078D4?style=for-the-badge&logo=microsoft-exchange&logoColor=white)

Advanced PowerShell script to audit email deletion activities in Exchange Online shared mailboxes with enterprise-grade features.

## Features

- 🔍 **Comprehensive Deletion Tracking**:
  - Soft deletes (MoveToDeletedItems)
  - Hard deletes (permanent removal)
  - Folder-level deletion tracking
- 🔐 **Multiple Authentication Methods**:
  - Interactive login (MFA supported)
  - Certificate-based authentication (CBA)
  - Service account credentials
- 📊 **Advanced Filtering**:
  - Filter by specific shared mailbox
  - Filter by user who performed deletion
  - Filter by email subject
- ⚡ **Performance Optimized**:
  - Parallel processing with configurable throttling
  - Automatic rate limit handling
- 📁 **Enhanced CSV Export**:
  - Timestamped output files
  - Email subject tracking
  - Automatic file opening option

## Prerequisites

- PowerShell 5.1 or later
- Exchange Online PowerShell V2 module
- One of these roles:
  - Global Administrator
  - Compliance Administrator
  - Exchange Administrator

## Installation

1. Clone the repository:
   ```powershell
   git clone https://github.com/RapidScripter/shared-mailbox-deletion-audit.git
   cd shared-mailbox-deletion-audit

2. Install the required module:
   ```powershell
   Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
   ```

## Usage

### Basic Commands

```powershell
# Interactive MFA session (last 180 days)
.\AuditSharedMailboxDeletions.ps1

# Specific date range
.\AuditSharedMailboxDeletions.ps1 -StartDate "2025-01-01" -EndDate "2025-01-31"

# Audit specific shared mailbox
.\AuditSharedMailboxDeletions.ps1 -SharedMBIdentity "shared@domain.com"
```

### Advanced Options

| Parameter               | Description                          | Example                           |
|-------------------------|--------------------------------------|-----------------------------------|
| `-StartDate`            | Report start date                    | `-StartDate "2025-01-01"`         |
| `-EndDate`              | Report end date                      | `-EndDate "2025-01-31"`           |
| `-SharedMBIdentity`     | Specific shared mailbox to audit     | `-SharedMBIdentity "shared@domain.com"` |
| `-UserId`               | Filter by user who performed deletion| `-UserId "user@domain.com"`       |
| `-Subject`              | Filter by email subject              | `-Subject "Confidential"`         |
| `-ClientId`             | App ID for certificate authentication| `-ClientId "xxxxxxxx-xxxx..."`    |
| `-CertificateThumbprint`| Certificate thumbprint for auth      | `-CertificateThumbprint "A1B2..."`|
| `-OutputPath`           | Custom directory for report output   | `-OutputPath "C:\AuditReports"`   |
| `-ThrottleLimit`        | Parallel processing threads (1-100)  | `-ThrottleLimit 10`               |

## Output

The script generates a CSV report with these columns:

- **Activity Time**: When deletion occurred (UTC)
- **Shared Mailbox Name**: Target shared mailbox address
- **Performed By**: User who performed the deletion
- **Activity**: Deletion type (SoftDelete/HardDelete/MoveToDeletedItems)
- **No. of Emails Deleted**: Count of affected messages
- **Email Subjects**: Comma-separated list of subjects
- **Folder**: Source folder where deletion occurred
- **Result Status**: Success or failure status
- **More Info**: Full audit log details (JSON)

Sample output filename: `DeletedEmailsAuditReport_2025-04-03_143022.csv`

## Example Output

| Activity Time       | Shared Mailbox Name | Performed By      | Activity        | No. of Emails Deleted | Email Subjects       | Folder      | Result Status |
|---------------------|---------------------|-------------------|-----------------|-----------------------|----------------------|-------------|---------------|
| 2025-11-15 09:30:22 | shared@contoso.com  | admin@contoso.com | HardDelete      | 1                     | Q3 Budget Review     | Inbox       | Success       |
| 2025-11-16 14:15:41 | finance@contoso.com | user@contoso.com  | MoveToDeletedItems | 3                  | Invoice_2023-11.pdf  | Sent Items  | Success       |
| 2025-11-17 11:20:33 | legal@contoso.com   | extern@partner.com| SoftDelete      | 5                     | NDA, Contract Draft | DeletedItems| Failed        |

## Troubleshooting

| Error/Symptom | Solution |
|--------------|----------|
| "Cannot connect to Exchange Online" | Verify admin permissions and MFA status |
| "No deletion events found" | Check date range and shared mailbox filters |
| "The term 'Connect-ExchangeOnline' is not recognized" | Install module: `Install-Module ExchangeOnlineManagement` |
| Throttling errors | Reduce `-ThrottleLimit` or increase time interval |

## Best Practices

1. **Regular Audits**: Schedule weekly runs for high-risk mailboxes
2. **Certificate Authentication**: Recommended for automated executions
3. **Bulk Deletion Alerts**: Monitor deletions >10 emails
4. **Retention Policies**: Combine with Microsoft 365 retention rules

## Scheduling
Create Windows Task Scheduler job with:

```powershell
# Certificate-based authentication (recommended)
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument `
  "-NoProfile -ExecutionPolicy Bypass -File `"C:\Scripts\AuditSharedMailboxDeletions.ps1`" `
  -ClientId `"your-app-id`" -CertificateThumbprint `"your-thumbprint`" `
  -Organization `"yourtenant.onmicrosoft.com`" -OutputPath `"C:\ScheduledReports`""

# Or with service account credentials
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument `
  "-NoProfile -ExecutionPolicy Bypass -File `"C:\Scripts\AuditSharedMailboxDeletions.ps1`" `
  -UserName `"admin@domain.com`" -Password `"yourpassword`" `
  -OutputPath `"C:\ScheduledReports`""

# Configure trigger (runs every Monday at 3 AM)
$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 3am

# Register the task (run as highest privileges)
Register-ScheduledTask -TaskName "Weekly Shared Mailbox Deletion Audit" `
  -Action $Action -Trigger $Trigger -RunLevel Highest -Description "Automated audit of shared mailbox deletions"
