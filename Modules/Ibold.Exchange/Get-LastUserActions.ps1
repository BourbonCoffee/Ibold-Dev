<#
.SYNOPSIS
   Retrieves last user activity information from Exchange Online mailboxes.

.DESCRIPTION
   This script connects to Exchange Online and Microsoft Graph to gather detailed 
   user activity information for all mailboxes. It exports data including last 
   email, calendar, contact usage, and sign-in information to a CSV file.

.PARAMETER AdminUserPrincipalName
   The admin UPN to use for connecting to Exchange Online and Microsoft Graph.
   If not provided, the script will prompt for credentials.

.PARAMETER OutputPath
   The path where the output CSV file will be saved.
   Default is "[Desktop]\User-LastActionsReport.csv"

.EXAMPLE
   .\Get-LastUserActions.ps1
   Connects with default credentials and exports data to the desktop.

.EXAMPLE
   .\Get-LastUserActions.ps1 -AdminUserPrincipalName admin@contoso.com -OutputPath "C:\Reports\UserActivity.csv"
   Connects with the specified admin account and exports data to the specified path.

.NOTES
   Author: Chris Ibold
   Comet Consulting Group
   Version: 2.0
   Date: 2025-04-21
   Requirements: 
   - ExchangeOnlineManagement module
   - Microsoft.Graph module
   - Administrator rights in Exchange Online and Microsoft Entra ID
#>

[CmdletBinding()]
param (
   [Parameter(Mandatory = $false)]
   [string]$AdminUserPrincipalName,

   [Parameter(Mandatory = $false)]
   [string]$OutputPath = (Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath "User-LastActionsReport.csv")
)

function Connect-RequiredServices {
   [CmdletBinding()]
   param()

   try {
      # Connect to Exchange Online
      Write-Verbose "Connecting to Exchange Online..."
      if ($AdminUserPrincipalName) {
         Connect-ExchangeOnline -UserPrincipalName $AdminUserPrincipalName -ErrorAction Stop
      } else {
         Connect-ExchangeOnline -ErrorAction Stop
      }
      Write-Verbose "Successfully connected to Exchange Online"

      # Connect to Microsoft Graph
      Write-Verbose "Connecting to Microsoft Graph..."
      Connect-MgGraph -Scope "User.Read.All", "AuditLog.Read.All" -ErrorAction Stop
      Write-Verbose "Successfully connected to Microsoft Graph"

      return $true
   } catch {
      Write-Error "Failed to connect to required services: $_"
      return $false
   }
}

function Get-UserMailboxActivity {
   [CmdletBinding()]
   param()

   try {
      # Initialize variables
      $now = Get-Date
      $mailboxCounter = 0
      $report = [System.Collections.Generic.List[Object]]::new()

      # Get all user mailboxes
      Write-Verbose "Retrieving user mailboxes..."
      $mailboxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | 
         Select-Object DisplayName, DistinguishedName, UserPrincipalName, ExternalDirectoryObjectId | 
         Sort-Object DisplayName

      Write-Verbose "Found $($mailboxes.Count) mailboxes to process"

      # Process each mailbox
      foreach ($mailbox in $mailboxes) {
         $mailboxCounter++
         Write-Progress -Activity "Processing Mailboxes" -Status "Processing $($mailbox.DisplayName)" -PercentComplete (($mailboxCounter / $mailboxes.count) * 100)
            
         try {
            # Get mailbox diagnostic logs
            $diagnosticLogs = Export-MailboxDiagnosticLogs -Identity $mailbox.DistinguishedName -ExtendedProperties
            $logXml = [xml]($diagnosticLogs.MailboxLog)

            # Extract activity timestamps
            $lastEmail = ($logXml.Properties.MailboxTable.Property | Where-Object { $_.Name -eq "LastEmailTimeCurrentValue" }).Value
            $lastCalendar = ($logXml.Properties.MailboxTable.Property | Where-Object { $_.Name -eq "LastCalendarTimeCurrentValue" }).Value
            $lastContacts = ($logXml.Properties.MailboxTable.Property | Where-Object { $_.Name -eq "LastContactsTimeCurrentValue" }).Value
            $lastFile = ($logXml.Properties.MailboxTable.Property | Where-Object { $_.Name -eq "LastFileTimeCurrentValue" }).Value
            $lastLogonTime = ($logXml.Properties.MailboxTable.Property | Where-Object { $_.Name -eq "LastLogonTime" }).Value
            $lastActive = ($logXml.Properties.MailboxTable.Property | Where-Object { $_.Name -eq "LastUserActionWorkloadAggregateTime" }).Value

            # Calculate days since last activity
            $daysSinceActive = if ($lastActive) {
                  (New-TimeSpan -Start $lastActive -End $now).Days
            } else {
               "N/A"
            }

            # Get mailbox statistics
            $stats = Get-EXOMailboxStatistics -Identity $mailbox.DistinguishedName
            $mailboxSize = ($stats.TotalItemSize.Value.ToString()).Split("(")[0]

            # Get last sign-in information
            $lastUserSignIn = $null
            $lastUserSignIn = (Get-MgAuditLogSignIn -Filter "UserId eq '$($mailbox.ExternalDirectoryObjectId)'" -Top 1).CreatedDateTime
            $lastUserSignInDate = if ($lastUserSignIn) {
               Get-Date($lastUserSignIn) -Format g
            } else {
               "No sign in records found in last 30 days"
            }

            # Get account enabled status
            $accountEnabled = (Get-MgUser -UserId $mailbox.ExternalDirectoryObjectId -Property AccountEnabled).AccountEnabled

            # Create report entry
            $reportLine = [PSCustomObject]@{ 
               Mailbox         = $mailbox.DisplayName 
               UPN             = $mailbox.UserPrincipalName
               Enabled         = $accountEnabled
               Items           = $stats.ItemCount 
               Size            = $mailboxSize 
               LastLogonExo    = $lastLogonTime
               LastLogonAD     = $lastUserSignInDate
               DaysSinceActive = $daysSinceActive
               LastActive      = $lastActive
               LastEmail       = $lastEmail
               LastCalendar    = $lastCalendar
               LastContacts    = $lastContacts
               LastFile        = $lastFile
            }

            $report.Add($reportLine)
         } catch {
            Write-Warning "Error processing mailbox $($mailbox.DisplayName): $_"
         }
      }

      return $report
   } catch {
      Write-Error "Error retrieving mailbox activity: $_"
      return $null
   }
}

function Export-ActivityReport {
   [CmdletBinding()]
   param (
      [Parameter(Mandatory = $true)]
      [System.Collections.Generic.List[Object]]$ReportData,

      [Parameter(Mandatory = $true)]
      [string]$FilePath
   )

   try {
      # Ensure output directory exists
      $outputDirectory = Split-Path -Path $FilePath -Parent
      if (-not (Test-Path -Path $outputDirectory)) {
         New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
      }

      # Export to CSV
      $ReportData | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
      Write-Verbose "Report exported successfully to $FilePath"
      return $true
   } catch {
      Write-Error "Failed to export report to $FilePath : $_"
      return $false
   }
}

# Main script execution
try {
   Write-Host "Starting Exchange Online user activity report..." -ForegroundColor Cyan

   # Connect to required services
   if (-not (Connect-RequiredServices)) {
      throw "Failed to connect to required services. Exiting script."
   }

   # Get user mailbox activity
   Write-Host "Retrieving mailbox activity data..." -ForegroundColor Cyan
   $activityData = Get-UserMailboxActivity

   if ($null -eq $activityData -or $activityData.Count -eq 0) {
      throw "No mailbox activity data retrieved."
   }

   # Export activity report
   Write-Host "Exporting activity report..." -ForegroundColor Cyan
   if (Export-ActivityReport -ReportData $activityData -FilePath $OutputPath) {
      Write-Host "Successfully exported user activity report to: $OutputPath" -ForegroundColor Green
   } else {
      throw "Failed to export activity report."
   }
} catch {
   Write-Host "Error executing script: $_" -ForegroundColor Red
} finally {
   # Clean up connections if needed
   Write-Host "Script execution complete" -ForegroundColor Cyan
}