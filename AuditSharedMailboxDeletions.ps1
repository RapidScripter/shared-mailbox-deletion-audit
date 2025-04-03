<#
=============================================================================================
.SYNOPSIS
    Audits email deletion activities in Exchange Online shared mailboxes.

.DESCRIPTION
    Tracks and reports on email deletions (SoftDelete, HardDelete, MoveToDeletedItems) 
    across shared mailboxes with advanced filtering capabilities.

.VERSION
    2.0

.FEATURES
    - Tracks all deletion types with detailed metadata
    - Supports custom date ranges (up to 180 days)
    - Multiple authentication methods (MFA, Certificate, Basic)
    - Parallel processing for large datasets
    - Comprehensive error handling and logging
    - Scheduled task friendly with parameterized credentials
=============================================================================================
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$EndDate,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern("^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")]
    [string]$SharedMBIdentity,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern("^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")]
    [string]$UserId,
    
    [Parameter(Mandatory = $false)]
    [string]$Subject,
    
    [Parameter(Mandatory = $false)]
    [string]$UserName,
    
    [Parameter(Mandatory = $false)]
    [string]$Password,
    
    [Parameter(Mandatory = $false)]
    [string]$Organization,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    
    [Parameter(Mandatory = $false)]
    [string]$CertificatePath,
    
    [Parameter(Mandatory = $false)]
    [securestring]$CertificatePassword,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10080)]
    [int]$IntervalMinutes = 1440,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$ThrottleLimit = 5,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Get-Location)
)

#region Initialization
$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"
$ProgressPreference = "SilentlyContinue"

$MaxStartDate = ((Get-Date).AddDays(-180)).Date
$OperationNames = "SoftDelete,HardDelete,MoveToDeletedItems"
$ScriptVersion = "2.0"
$ExecutionStartTime = Get-Date
$OutputFileName = "DeletedEmailsAuditReport_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').csv"
$OutputCSV = Join-Path $OutputPath $OutputFileName
#endregion

#region Functions
function Connect-Exchange {
    [CmdletBinding()]
    param()
    
    try {
        if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
            Write-Host "Installing ExchangeOnline module..." -ForegroundColor Yellow
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
        }

        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        
        if ($ClientId -and $CertificateThumbPrint) {
            Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbPrint -Organization $Organization -ShowBanner:$false
        }
        elseif ($ClientId -and $CertificatePath) {
            Connect-ExchangeOnline -AppId $ClientId -CertificateFilePath $CertificatePath -CertificatePassword $CertificatePassword -Organization $Organization -ShowBanner:$false
        }
        elseif ($UserName -and $Password) {
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
            $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword
            Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
        }
        else {
            Connect-ExchangeOnline -ShowBanner:$false
        }
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $_"
        exit 1
    }
}

function Get-SharedMailboxes {
    [CmdletBinding()]
    param(
        [string]$Identity
    )
    
    try {
        if ([string]::IsNullOrEmpty($Identity)) {
            return (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox -ErrorAction Stop).PrimarySMTPAddress
        }
        else {
            $mailbox = Get-Mailbox -Identity $Identity -RecipientTypeDetails SharedMailbox -ErrorAction Stop
            if (-not $mailbox) {
                Write-Host "Shared mailbox '$Identity' not found or is not a shared mailbox" -ForegroundColor Red
                exit 1
            }
            return $mailbox.PrimarySMTPAddress
        }
    }
    catch {
        Write-Error "Failed to retrieve shared mailboxes: $_"
        exit 1
    }
}

function Process-AuditResults {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        $Results,
        [string[]]$SharedMailboxes,
        [string]$FilterSubject,
        [string]$FilterUser
    )
    
    begin {
        $processedCount = 0
        $outputEvents = @()
    }
    
    process {
        $Results | ForEach-Object -Parallel {
            $record = $_
            try {
                $auditData = $record.auditdata | ConvertFrom-Json
                $targetMailbox = $auditData.MailboxOwnerUPN

                # Apply filters
                if ($using:SharedMailboxes -and (-not ($using:SharedMailboxes -contains $targetMailbox))) {
                    return
                }

                if ($using:FilterUser -and $using:FilterUser -ne $auditData.userId) {
                    return
                }

                $emailSubjects = if ($auditData.AffectedItems.Subject) { 
                    $auditData.AffectedItems.Subject -join ", " 
                } else { "-" }

                if ($using:FilterSubject -and $emailSubjects -notmatch $using:FilterSubject) {
                    return
                }

                $folderPath = if ($auditData.Folder.Path) { 
                    $auditData.Folder.Path.Split("\")[1] 
                } else { "Unknown" }

                [PSCustomObject]@{
                    'Activity Time'         = Get-Date $auditData.CreationTime
                    'Shared Mailbox Name'   = $targetMailbox
                    'Activity'              = $record.Operations
                    'Performed By'          = $auditData.userId
                    'No. of Emails Deleted' = ($auditData.AffectedItems.Subject | Measure-Object).Count
                    'Email Subjects'        = $emailSubjects
                    'Folder'                = $folderPath
                    'Result Status'         = $auditData.ResultStatus
                    'More Info'             = $record.auditdata
                }
            }
            catch {
                Write-Warning "Error processing record: $_"
            }
        } -ThrottleLimit $ThrottleLimit | ForEach-Object {
            $outputEvents += $_
            $processedCount++
            if ($processedCount % 100 -eq 0) {
                Write-Progress -Activity "Processing audit records" -Status "Processed $processedCount records"
            }
        }
    }
    
    end {
        return $outputEvents
    }
}
#endregion

#region Main Execution
try {
    # Validate and set date range
    if (-not $StartDate -and -not $EndDate) {
        $EndDate = (Get-Date).Date
        $StartDate = $MaxStartDate
    }

    $StartDate = [DateTime]$StartDate
    $EndDate = [DateTime]$EndDate

    if ($StartDate -lt $MaxStartDate) {
        throw "Audit can only be retrieved for past 180 days. Please select a date after $MaxStartDate"
    }

    if ($EndDate -lt $StartDate) {
        throw "End time should be later than start time"
    }

    # Initialize connection
    Connect-Exchange

    # Load shared mailboxes
    $SharedMailboxes = Get-SharedMailboxes -Identity $SharedMBIdentity

    # Prepare output file
    if (Test-Path $OutputCSV) {
        Remove-Item $OutputCSV -Force
    }

    # Initialize processing
    $CurrentStart = $StartDate
    $CurrentEnd = $CurrentStart.AddMinutes($IntervalMinutes)
    if ($CurrentEnd -gt $EndDate) {
        $CurrentEnd = $EndDate
    }

    # Main processing loop
    $totalRecords = 0
    while ($true) {
        try {
            Write-Host "Retrieving logs from $CurrentStart to $CurrentEnd..." -ForegroundColor Cyan
            
            $searchParams = @{
                StartDate       = $CurrentStart
                EndDate         = $CurrentEnd
                Operations      = $OperationNames
                SessionId       = "SharedMBDeletionAudit"
                SessionCommand  = "ReturnLargeSet"
                ResultSize      = 5000
            }

            if ($UserId) {
                $searchParams["UserIds"] = $UserId
            }

            $Results = Search-UnifiedAuditLog @searchParams

            if ($Results) {
                $processedResults = $Results | Process-AuditResults -SharedMailboxes $SharedMailboxes -FilterSubject $Subject -FilterUser $UserId
                $processedResults | Export-Csv -Path $OutputCSV -NoTypeInformation -Append
                $totalRecords += $processedResults.Count
            }

            # Pagination logic
            if ($Results.Count -lt 5000) {
                if ($CurrentEnd -ge $EndDate) { break }
                $CurrentStart = $CurrentEnd
                $CurrentEnd = $CurrentStart.AddMinutes($IntervalMinutes)
                if ($CurrentEnd -gt $EndDate) { $CurrentEnd = $EndDate }
            }

            # Rate limiting
            Start-Sleep -Milliseconds 500
        }
        catch {
            Write-Warning "Error processing batch: $_"
            if ($_.Exception.Message -match "throttled") {
                $retrySeconds = 30
                Write-Host "Throttling detected. Waiting $retrySeconds seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $retrySeconds
                continue
            }
            break
        }
    }

    # Output results
    if ($totalRecords -gt 0) {
        Write-Host "`nSuccessfully processed $totalRecords audit records" -ForegroundColor Green
        Write-Host "Report saved to: $OutputCSV" -ForegroundColor Cyan
        
        # Option to open file
        $openFile = Read-Host "Open report file now? (Y/N)"
        if ($openFile -match "[yY]") {
            Invoke-Item $OutputCSV
        }
    }
    else {
        Write-Host "No matching records found for the specified criteria" -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Script execution failed: $_"
}
finally {
    # Clean up connection
    try {
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Error disconnecting: $_"
    }

    # Calculate and display execution time
    $executionTime = (Get-Date) - $ExecutionStartTime
    Write-Host "`nScript execution completed in $($executionTime.TotalMinutes.ToString('0.00')) minutes" -ForegroundColor Cyan
}
#endregion