<#
.SYNOPSIS
    Discovers and optionally removes email addresses with specified domains from Exchange objects.

.DESCRIPTION
    This script identifies Mailboxes, MailUsers, Distribution, Security and Unified Groups using a given SMTP domain 
    in their email addresses. By default, it generates a report of all objects with addresses from the specified domain.
    Optionally, it can remove all email addresses for the given SMTP domain from these objects.

.PARAMETER StartDirectory
    Specifies the default location when asking for the input CSV file.
    The CSV file must contain a column called DomainName with domain values (e.g., contoso.com).

.PARAMETER Type
    Specifies where the SMTP addresses are located: "On-Premises" or "Office365".
    Default: "Office365"

.PARAMETER ObjectType
    Specifies which object types to process: "MailboxOnly", "GroupsOnly", or "All".
    Default: "All"

.PARAMETER IncludeLegacyExchangeDN
    When specified, includes the LegacyExchangeDN in the export file.

.PARAMETER OrgMgmtAdmin
    Specifies the path to the encrypted credentials for Exchange on-premises access.
    Example: "C:\Temp\orgmgmt.xml"

.PARAMETER ExchangeServer
    Specifies the fully qualified domain name of the on-premises Exchange server.
    Example: "server1.contoso.com"

.PARAMETER O365Admin
    Specifies the path to the encrypted credentials for Office 365 access.
    Example: "C:\Temp\o365admin.xml"

.PARAMETER DestroyMethod
    Specifies the method to DELETE email addresses:
    - "Destroy": Uses output from discovery process to remove email addresses
    - "DestroyOnly": Prompts for a CSV of email addresses to remove

.PARAMETER WhatIf
    When specified, shows what would happen if the script runs without making actual changes.

.EXAMPLE
    Proxy-Seek-and-Destroy.ps1 -StartDirectory C:\temp -Type Office365 -ObjectType All -O365Admin C:\temp\O365Creds.xml
    
    Discovers all objects with email addresses from domains specified in a CSV file.

.EXAMPLE
    Proxy-Seek-and-Destroy.ps1 -StartDirectory C:\temp -Type Office365 -ObjectType All -O365Admin C:\temp\O365Creds.xml -DestroyMethod Destroy
    
    Discovers objects and removes email addresses from domains specified in a CSV file.

.NOTES
    AUTHOR: Chris Ibold
    COMPANY: Comet Consulting Group
    CREATED: 2023-01-01
    MODIFIED: 2025-05-15
    
    IMPORTANT NOTES:
    - UserPrincipalName should be changed first when using Office 365, as it will lock certain proxyAddresses
    - This script does not handle scenarios where the license is just removed from the mailbox
#>

[CmdletBinding(SupportsShouldProcess = $true)]
Param(
    [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the starting location for picking a CSV file (ex: C:\Temp)")]
    [ValidateNotNullOrEmpty()]
    [String]$StartDirectory,

    [Parameter(Position = 2, Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Select where the SMTP addresses should be located")]
    [ValidateSet("On-Premises", "Office365")]
    [String]$Type = "Office365",

    [Parameter(Position = 3, Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Select the type of objects with the SMTP addresses")]
    [ValidateSet("MailboxOnly", "GroupsOnly", "All")]
    [String]$ObjectType = "All",

    [Parameter(Position = 4, Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Include LegacyExchangeDN as an entry in the export")]
    [Switch]$IncludeLegacyExchangeDN,

    [Parameter(Position = 5, Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the encrypted credentials for on-premises access")]
    [ValidateNotNullOrEmpty()]
    [String]$OrgMgmtAdmin,

    [Parameter(Position = 6, Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the FQDN of the Exchange Server on-premises")]
    [ValidateNotNullOrEmpty()]
    [String]$ExchangeServer,

    [Parameter(Position = 7, Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the encrypted credentials for Office 365 access")]
    [ValidateNotNullOrEmpty()]
    [String]$O365Admin,

    [Parameter(Position = 8, Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Select the method to DELETE the email addresses")]
    [ValidateSet("Destroy", "DestroyOnly")]
    [String]$DestroyMethod
)

#region Script Initialization
# Set script variables
$scriptPath = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$currentDate = Get-Date -Format "yyyyMMdd_HHmmss"
$logDirectory = Join-Path -Path $scriptPath -ChildPath "Logging"

# Create logging directory if it doesn't exist
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
}

$logFileName = Join-Path -Path $logDirectory -ChildPath "$currentDate-Proxy-Seek-and-Destroy.log"
$globalErrorLogStream = $null
#endregion

#region Helper Functions
Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [String]$LogFile = $logFileName,
        
        [Parameter(Mandatory = $true)]
        [String]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Information", "Warning", "Error", "Verbose", "Debug")]
        [String]$Type = "Information",
        
        [Parameter(Mandatory = $false)]
        [Switch]$WriteToHost = $true
    )
    
    # Ensure log directory exists
    $logDir = Split-Path -Path $LogFile -Parent
    if (-not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    # Build timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Format log entry
    $logEntry = "[$timestamp] [$Type] $Message"
    
    # Write to log file
    Add-Content -Path $LogFile -Value $logEntry
    
    # Optionally write to host with appropriate coloring
    if ($WriteToHost) {
        switch ($Type) {
            "Information" { Write-Host $logEntry -ForegroundColor Cyan }
            "Warning" { Write-Host $logEntry -ForegroundColor Yellow }
            "Error" { Write-Host $logEntry -ForegroundColor Red }
            "Verbose" { 
                if ($VerbosePreference -eq "Continue") {
                    Write-Host $logEntry -ForegroundColor Gray
                }
            }
            "Debug" { 
                if ($DebugPreference -eq "Continue") {
                    Write-Host $logEntry -ForegroundColor Magenta
                }
            }
        }
    }
}

Function Get-InputFilePath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$StartDirectory
    )
    
    try {
        # Prompts the user for the input file starting in $StartDirectory
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.InitialDirectory = $StartDirectory
        $openFileDialog.Filter = "CSV files (*.csv)| *.csv|All files (*.*)| *.*"
        $openFileDialog.ShowDialog() | Out-Null
        
        if ([String]::IsNullOrEmpty($openFileDialog.FileName)) {
            Write-Log -Message "No file was selected by the user" -Type Warning
            return $null
        }
        
        return $openFileDialog.FileName
    } catch {
        Write-Log -Message "Error in file selection dialog: $($_.Exception.Message)" -Type Error
        return $null
    }
}

Function Test-ModuleInstalled {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$ModuleName
    )
    
    $module = Get-Module -Name $ModuleName -ListAvailable -Verbose:$false
    return ($null -ne $module)
}

Function Import-RequiredModules {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]]$ModuleNames
    )
    
    foreach ($moduleName in $ModuleNames) {
        if (-not (Test-ModuleInstalled -ModuleName $moduleName)) {
            Write-Log -Message "$moduleName module is not installed" -Type Warning
            
            # Check if NuGet package provider is available
            $nuget = Get-PackageProvider -ListAvailable | Where-Object { $_.Name -eq "NuGet" }
            if (-not $nuget) {
                Write-Log -Message "NuGet package provider is not installed. Installing now..." -Type Information
                try {
                    Install-PackageProvider -Name NuGet -Force -Scope AllUsers -ErrorAction Stop
                    Write-Log -Message "NuGet package provider installed successfully" -Type Information
                } catch {
                    Write-Log -Message "Failed to install NuGet package provider: $($_.Exception.Message)" -Type Error
                    throw "Failed to install NuGet package provider. Please install it manually."
                }
            }
            
            # Install the required module
            try {
                Write-Log -Message "Installing $moduleName module..." -Type Information
                Install-Module -Name $moduleName -Force -Scope AllUsers -ErrorAction Stop
                Write-Log -Message "$moduleName module installed successfully" -Type Information
            } catch {
                Write-Log -Message "Failed to install $moduleName module: $($_.Exception.Message)" -Type Error
                throw "Failed to install $moduleName module. Please install it manually."
            }
        }
        
        # Import the module
        try {
            Import-Module -Name $moduleName -ErrorAction Stop
            Write-Log -Message "Successfully imported $moduleName module" -Type Information
        } catch {
            Write-Log -Message "Failed to import $moduleName module: $($_.Exception.Message)" -Type Error
            throw "Failed to import $moduleName module. Please check if it's correctly installed."
        }
    }
}
#endregion

#region Connection Functions
Function Connect-ExchangeOnPremises {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 1, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$ExchangeServer,

        [Parameter(Position = 2, Mandatory = $false)]
        [System.Management.Automation.PSCredential]$ExchangeCredential
    )

    $sessionOptions = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -OpenTimeout 20000
    
    try {
        # Attempt #1 http and current logged on user
        Write-Log -Message "Attempting to connect to Exchange On-Premises with Kerberos authentication over HTTP" -Type Information
        $session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell") -Authentication Kerberos -AllowRedirection -SessionOption $sessionOptions -ErrorAction Stop
        Import-PSSession -Session $session -AllowClobber | Out-Null
        Set-ADServerSettings -ViewEntireForest:$true
        Write-Log -Message "Successfully connected to Exchange on-premises over HTTP using Kerberos" -Type Information
        return $session
    } catch {
        Write-Log -Message "Failed to connect using HTTP with Kerberos: $($_.Exception.Message)" -Type Warning
        
        try {
            # Attempt #2 https and current logged on user
            Write-Log -Message "Attempting to connect to Exchange On-Premises with Kerberos authentication over HTTPS" -Type Information
            $session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("https://" + $ExchangeServer + "/PowerShell") -Authentication Kerberos -AllowRedirection -SessionOption $sessionOptions -ErrorAction Stop
            Import-PSSession -Session $session -AllowClobber | Out-Null
            Set-ADServerSettings -ViewEntireForest:$true
            Write-Log -Message "Successfully connected to Exchange on-premises over HTTPS using Kerberos" -Type Information
            return $session
        } catch {
            Write-Log -Message "Failed to connect using HTTPS with Kerberos: $($_.Exception.Message)" -Type Warning
            
            try {
                # Attempt #3 https and provided creds
                if ($null -eq $ExchangeCredential) {
                    Write-Log -Message "Kerberos authentication failed. Please provide credentials." -Type Warning
                    $ExchangeCredential = Get-Credential -Message "Enter credentials for Exchange on-premises"
                }
                
                Write-Log -Message "Attempting to connect to Exchange On-Premises with provided credentials over HTTPS" -Type Information
                $session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("https://" + $ExchangeServer + "/PowerShell") -Credential $ExchangeCredential -AllowRedirection -SessionOption $sessionOptions -ErrorAction Stop
                Import-PSSession -Session $session -AllowClobber | Out-Null
                Set-ADServerSettings -ViewEntireForest:$true
                Write-Log -Message "Successfully connected to Exchange on-premises over HTTPS using provided credentials" -Type Information
                return $session
            } catch {
                Write-Log -Message "Failed to connect using HTTPS with credentials: $($_.Exception.Message)" -Type Warning
                
                try {
                    # Attempt #4 http and provided creds
                    Write-Log -Message "Attempting to connect to Exchange On-Premises with provided credentials over HTTP" -Type Information
                    $session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell") -Credential $ExchangeCredential -AllowRedirection -SessionOption $sessionOptions -ErrorAction Stop
                    Import-PSSession -Session $session -AllowClobber | Out-Null
                    Set-ADServerSettings -ViewEntireForest:$true
                    Write-Log -Message "Successfully connected to Exchange on-premises over HTTP using provided credentials" -Type Information
                    return $session
                } catch {
                    $errorMessage = "Failed to connect to Exchange on-premises server $ExchangeServer after multiple attempts: $($_.Exception.Message)"
                    Write-Log -Message $errorMessage -Type Error
                    throw $errorMessage
                }
            }
        }
    }
}

Function Connect-Office365Services {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [String]$CredentialPath
    )
    
    $credential = $null
    $mfaEnabled = $false
    
    # Try to load credentials if a path is provided
    if (-not [String]::IsNullOrEmpty($CredentialPath)) {
        try {
            $credential = Import-Clixml -Path $CredentialPath -ErrorAction Stop
            Write-Log -Message "Successfully loaded credentials from $CredentialPath" -Type Information
        } catch {
            Write-Log -Message "Failed to load credentials from $CredentialPath: $($_.Exception.Message)" -Type Warning
            Write-Log -Message "Will prompt for credentials or use modern authentication instead" -Type Information
            $mfaEnabled = $true
        }
    } else {
        Write-Log -Message "No credential path provided. Using modern authentication flow" -Type Information
        $mfaEnabled = $true
    }
    
    # Connect to Exchange Online
    try {
        # Import ExchangeOnlineManagement module
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        
        if ($mfaEnabled) {
            Write-Log -Message "Connecting to Exchange Online using modern authentication..." -Type Information
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        } else {
            Write-Log -Message "Connecting to Exchange Online using stored credentials..." -Type Information
            Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ErrorAction Stop
        }
        
        Write-Log -Message "Successfully connected to Exchange Online" -Type Information
    } catch {
        $errorMessage = "Failed to connect to Exchange Online: $($_.Exception.Message)"
        Write-Log -Message $errorMessage -Type Error
        throw $errorMessage
    }
    
    # Connect to Azure AD
    try {
        if ($mfaEnabled) {
            Write-Log -Message "Connecting to Azure AD using modern authentication..." -Type Information
            Connect-AzureAD -ErrorAction Stop | Out-Null
        } else {
            Write-Log -Message "Connecting to Azure AD using stored credentials..." -Type Information
            Connect-AzureAD -Credential $credential -ErrorAction Stop | Out-Null
        }
        
        Write-Log -Message "Successfully connected to Azure AD" -Type Information
    } catch {
        $errorMessage = "Failed to connect to Azure AD: $($_.Exception.Message)"
        Write-Log -Message $errorMessage -Type Error
        
        # This is non-fatal as we might not need Azure AD for all operations
        Write-Log -Message "Continuing script execution, but some Azure AD-dependent operations may fail" -Type Warning
    }
}
#endregion

#region Address Discovery Functions
Function Get-MailboxDetail {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$Domain,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$OutputFile,
        
        [Parameter(Mandatory = $false)]
        [Switch]$IncludeLegacyExchangeDN
    )
    
    # Find all mailboxes with the desired domain name being used
    $domainMatch = "*@" + $Domain.DomainName
    $domainString = ("`'$domainMatch`'").ToString()
    
    try {
        Write-Log -Message "Searching for mailboxes with domain: $domainMatch" -Type Information
        $recipients = Get-Recipient -ResultSize Unlimited -Filter "EmailAddresses -like $domainString -and (RecipientTypeDetails -eq 'UserMailbox' -or RecipientTypeDetails -eq 'SharedMailbox' -or RecipientTypeDetails -eq 'RoomMailbox')" -ErrorAction Stop
        $totalCount = ($recipients | Measure-Object).Count
        Write-Log -Message "Discovered $totalCount mailboxes with $domainMatch used" -Type Information
    } catch {
        Write-Log -Message "Error searching for mailboxes: $($_.Exception.Message)" -Type Error
        return
    }
    
    # Loop through the list of mailboxes and find the exact address that matched
    $i = 0
    foreach ($recipient in $recipients) {
        try {
            if ($recipients.Count -gt 1) {
                $i++
                Write-Progress -Id 1 -Activity "Processing mailboxes" -Status "Processing $i of $($recipients.Count) users: $($recipient.PrimarySmtpAddress)" -PercentComplete (($i / $recipients.Count) * 100)
            }
            
            $addresses = $recipient.EmailAddresses -like $domainMatch
            foreach ($address in $addresses) {
                $record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName
                Add-Content -Path $OutputFile -Value $record
            }

            if ($IncludeLegacyExchangeDN) {
                # Export the Exchange Legacy DN as well
                $legacyDN = if ($recipient.PSObject.Properties.Name -contains "LegacyExchangeDN") { $recipient.LegacyExchangeDN } else { $null }
                
                if ($legacyDN) {
                    $record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName + "`t" + $legacyDN
                    Add-Content -Path $OutputFile -Value $record
                }
            }
        } catch {
            Write-Log -Message "Error processing mailbox $($recipient.DisplayName): $($_.Exception.Message)" -Type Error
            continue
        }
    }
    
    Write-Progress -Id 1 -Activity "Processing mailboxes" -Completed
}

Function Get-InactiveMailboxDetail {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$Domain,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$OutputFile,
        
        [Parameter(Mandatory = $false)]
        [Switch]$IncludeLegacyExchangeDN
    )
    
    # Find all inactive mailboxes with the desired domain name being used
    $domainMatch = "*@" + $Domain.DomainName
    $domainString = ("`'$domainMatch`'").ToString()
    
    try {
        # Check if the Exchange environment supports inactive mailboxes
        $supportsInactiveMailbox = $true
        try {
            Get-Command -Name Get-Mailbox -ParameterName InactiveMailboxOnly -ErrorAction Stop | Out-Null
        } catch {
            $supportsInactiveMailbox = $false
            Write-Log -Message "This Exchange environment does not support inactive mailboxes. Skipping inactive mailbox search." -Type Warning
            return
        }
        
        if ($supportsInactiveMailbox) {
            Write-Log -Message "Searching for inactive mailboxes with domain: $domainMatch" -Type Information
            $recipients = Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited -Filter "EmailAddresses -like $domainString" -ErrorAction Stop
            $totalCount = ($recipients | Measure-Object).Count
            Write-Log -Message "Discovered $totalCount inactive mailboxes with $domainMatch used" -Type Information
        }
    } catch {
        Write-Log -Message "Error searching for inactive mailboxes: $($_.Exception.Message)" -Type Error
        return
    }
    
    # Loop through the list of inactive mailboxes and find the exact address that matched
    $i = 0
    foreach ($recipient in $recipients) {
        try {
            if ($recipients.Count -gt 1) {
                $i++
                Write-Progress -Id 1 -Activity "Processing inactive mailboxes" -Status "Processing $i of $($recipients.Count) users: $($recipient.PrimarySmtpAddress)" -PercentComplete (($i / $recipients.Count) * 100)
            }
            
            $addresses = $recipient.EmailAddresses -like $domainMatch
            foreach ($address in $addresses) {
                $record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + $true + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName
                Add-Content -Path $OutputFile -Value $record
            }

            if ($IncludeLegacyExchangeDN) {
                # Export the Exchange Legacy DN as well
                $legacyDN = if ($recipient.PSObject.Properties.Name -contains "LegacyExchangeDN") { $recipient.LegacyExchangeDN } else { $null }
                
                if ($legacyDN) {
                    $record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + $true + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName + "`t" + $legacyDN
                    Add-Content -Path $OutputFile -Value $record
                }
            }
        } catch {
            Write-Log -Message "Error processing inactive mailbox $($recipient.DisplayName): $($_.Exception.Message)" -Type Error
            continue
        }
    }
    
    Write-Progress -Id 1 -Activity "Processing inactive mailboxes" -Completed
}

Function Get-MailUserDetail {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$Domain,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$OutputFile,
        
        [Parameter(Mandatory = $false)]
        [Switch]$IncludeLegacyExchangeDN
    )
    
    # Find all mail users with the desired domain name being used
    $domainMatch = "*@" + $Domain.DomainName
    $domainString = ("`'$domainMatch`'").ToString()
    
    try {
        Write-Log -Message "Searching for mail users with domain: $domainMatch" -Type Information
        $recipients = Get-Recipient -ResultSize Unlimited -Filter "EmailAddresses -like $domainString -and RecipientTypeDetails -eq 'MailUser'" -ErrorAction Stop
        $totalCount = ($recipients | Measure-Object).Count
        Write-Log -Message "Discovered $totalCount mail users with $domainMatch used" -Type Information
    } catch {
        Write-Log -Message "Error searching for mail users: $($_.Exception.Message)" -Type Error
        return
    }
    
    # Loop through the list of mail users and find the exact address that matched
    $i = 0
    foreach ($recipient in $recipients) {
        try {
            if ($recipients.Count -gt 1) {
                $i++
                Write-Progress -Id 1 -Activity "Processing mail users" -Status "Processing $i of $($recipients.Count) users: $($recipient.PrimarySmtpAddress)" -PercentComplete (($i / $recipients.Count) * 100)
            }
            
            $addresses = $recipient.EmailAddresses -like $domainMatch
            foreach ($address in $addresses) {
                $record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName
                Add-Content -Path $OutputFile -Value $record
            }

            if ($IncludeLegacyExchangeDN) {
                # Export the Exchange Legacy DN as well
                $legacyDN = if ($recipient.PSObject.Properties.Name -contains "LegacyExchangeDN") { $recipient.LegacyExchangeDN } else { $null }
                
                if ($legacyDN) {
                    $record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName + "`t" + $legacyDN
                    Add-Content -Path $OutputFile -Value $record
                }
            }
        } catch {
            Write-Log -Message "Error processing mail user $($recipient.DisplayName): $($_.Exception.Message)" -Type Error
            continue
        }
    }
    
    Write-Progress -Id 1 -Activity "Processing mail users" -Completed
}

Function Get-DistributionGroupDetail {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$Domain,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$OutputFile,
        
        [Parameter(Mandatory = $false)]
        [Switch]$IncludeLegacyExchangeDN
    )
    
    # Find all distribution groups with the desired domain name being used
    $domainMatch = "*@" + $Domain.DomainName
    $domainString = ("`'$domainMatch`'").ToString()
    
    try {
        Write-Log -Message "Searching for distribution groups with domain: $domainMatch" -Type Information
        $recipients = Get-DistributionGroup -ResultSize Unlimited -Filter "EmailAddresses -like $domainString -and (RecipientTypeDetails -eq 'MailUniversalDistributionGroup' -or RecipientTypeDetails -eq 'MailUniversalSecurityGroup')" -ErrorAction Stop
        $totalCount = ($recipients | Measure-Object).Count
        Write-Log -Message "Discovered $totalCount distribution groups with $domainMatch used" -Type Information
    } catch {
        Write-Log -Message "Error searching for distribution groups: $($_.Exception.Message)" -Type Error
        return
    }
    
    # Loop through the list of distribution groups and find the exact address that matched
    $i = 0
    foreach ($recipient in $recipients) {
        try {
            if ($recipients.Count -gt 1) {
                $i++
                Write-Progress -Id 1 -Activity "Processing distribution groups" -Status "Processing $i of $($recipients.Count) groups: $($recipient.DisplayName)" -PercentComplete (($i / $recipients.Count) * 100)
            }
            
            $addresses = $recipient.EmailAddresses -like $domainMatch
            foreach ($address in $addresses) {
                $record = $recipient.Guid.Guid + "`t" + "`t" + $recipient.ExchangeObjectId + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName
                Add-Content -Path $OutputFile -Value $record
            }

            if ($IncludeLegacyExchangeDN) {
                # Export the Exchange Legacy DN as well
                $legacyDN = if ($recipient.PSObject.Properties.Name -contains "LegacyExchangeDN") { $recipient.LegacyExchangeDN } else { $null }
                
                if ($legacyDN) {
                    $record = $recipient.Guid.Guid + "`t" + "`t" + $recipient.ExchangeObjectId + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName + "`t" + $legacyDN
                    Add-Content -Path $OutputFile -Value $record
                }
            }
        } catch {
            Write-Log -Message "Error processing distribution group $($recipient.DisplayName): $($_.Exception.Message)" -Type Error
            continue
        }
    }
    
    Write-Progress -Id 1 -Activity "Processing distribution groups" -Completed
}

Function Get-RoomListDetail {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$Domain,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$OutputFile,
        
        [Parameter(Mandatory = $false)]
        [Switch]$IncludeLegacyExchangeDN
    )
    
    # Find all room lists with the desired domain name being used
    $domainMatch = "*@" + $Domain.DomainName
    $domainString = ("`'$domainMatch`'").ToString()
    
    try {
        Write-Log -Message "Searching for room lists with domain: $domainMatch" -Type Information
        $recipients = Get-DistributionGroup -ResultSize Unlimited -Filter "EmailAddresses -like $domainString -and RecipientTypeDetails -eq 'RoomList'" -ErrorAction Stop
        $totalCount = ($recipients | Measure-Object).Count
        Write-Log -Message "Discovered $totalCount room lists with $domainMatch used" -Type Information
    } catch {
        Write-Log -Message "Error searching for room lists: $($_.Exception.Message)" -Type Error
        return
    }
    
    # Loop through the list of room lists and find the exact address that matched
    $i = 0
    foreach ($recipient in $recipients) {
        try {
            if ($recipients.Count -gt 1) {
                $i++
                Write-Progress -Id 1 -Activity "Processing room lists" -Status "Processing $i of $($recipients.Count) room lists: $($recipient.DisplayName)" -PercentComplete (($i / $recipients.Count) * 100)
            }
            
            $addresses = $recipient.EmailAddresses -like $domainMatch
            foreach ($address in $addresses) {
                $record = $recipient.Guid.Guid + "`t" + "`t" + $recipient.ExchangeObjectId + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName
                Add-Content -Path $OutputFile -Value $record
            }

            if ($IncludeLegacyExchangeDN) {
                # Export the Exchange Legacy DN as well
                $legacyDN = if ($recipient.PSObject.Properties.Name -contains "LegacyExchangeDN") { $recipient.LegacyExchangeDN } else { $null }
                
                if ($legacyDN) {
                    $record = $recipient.Guid.Guid + "`t" + "`t" + $recipient.ExchangeObjectId + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName + "`t" + $legacyDN
                    Add-Content -Path $OutputFile -Value $record
                }
            }
        } catch {
            Write-Log -Message "Error processing room list $($recipient.DisplayName): $($_.Exception.Message)" -Type Error
            continue
        }
    }
    
    Write-Progress -Id 1 -Activity "Processing room lists" -Completed
}

Function Get-UnifiedGroupDetail {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$Domain,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$OutputFile,
        
        [Parameter(Mandatory = $false)]
        [Switch]$IncludeLegacyExchangeDN
    )
    
    # Find all unified groups with the desired domain name being used
    $domainMatch = "*@" + $Domain.DomainName
    $domainString = ("`'$domainMatch`'").ToString()
    
    try {
        Write-Log -Message "Searching for unified groups with domain: $domainMatch" -Type Information
        $recipients = Get-Recipient -ResultSize Unlimited -Filter "EmailAddresses -like $domainString -and RecipientTypeDetails -eq 'GroupMailbox'" -ErrorAction Stop
        $totalCount = ($recipients | Measure-Object).Count
        Write-Log -Message "Discovered $totalCount unified groups with $domainMatch used" -Type Information
    } catch {
        Write-Log -Message "Error searching for unified groups: $($_.Exception.Message)" -Type Error
        return
    }
    
    # Loop through the list of unified groups and find the exact address that matched
    $i = 0
    foreach ($recipient in $recipients) {
        try {
            if ($recipients.Count -gt 1) {
                $i++
                Write-Progress -Id 1 -Activity "Processing unified groups" -Status "Processing $i of $($recipients.Count) unified groups: $($recipient.DisplayName)" -PercentComplete (($i / $recipients.Count) * 100)
            }
            
            $addresses = $recipient.EmailAddresses -like $domainMatch
            foreach ($address in $addresses) {
                $record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName
                Add-Content -Path $OutputFile -Value $record
            }

            if ($IncludeLegacyExchangeDN) {
                # Export the Exchange Legacy DN as well
                $legacyDN = if ($recipient.PSObject.Properties.Name -contains "LegacyExchangeDN") { $recipient.LegacyExchangeDN } else { $null }
                
                if ($legacyDN) {
                    $record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.DisplayName + "`t" + $legacyDN
                    Add-Content -Path $OutputFile -Value $record
                }
            }
        } catch {
            Write-Log -Message "Error processing unified group $($recipient.DisplayName): $($_.Exception.Message)" -Type Error
            continue
        }
    }
    
    Write-Progress -Id 1 -Activity "Processing unified groups" -Completed
}
#endregion

#region Address Removal Functions
Function Remove-EmailAddress {
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$RecipientObject,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$AddressToRemove,
        
        [Parameter(Mandatory = $false)]
        [String]$TenantDomain
    )
    
    $address = $null
    $address = $AddressToRemove.Split(":")
    
    try {
        switch ($RecipientObject.RecipientTypeDetails) {
            { $_ -in @("UserMailbox", "SharedMailbox", "RoomMailbox") } {
                if ($RecipientObject.IsInactiveMailbox -ne $true) {
                    if ($address[0] -clike "SMTP") {
                        # This is a primary address - need to set a new primary first
                        $newPrimary = ($RecipientObject.Alias + "@" + $TenantDomain)
                        
                        if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Set new primary email address to $newPrimary")) {
                            Set-Mailbox -Identity $RecipientObject.Alias -WindowsEmailAddress $newPrimary -ErrorAction Stop
                            Write-Log -Message "Changed primary address from $($RecipientObject.PrimarySmtpAddress) to $newPrimary" -Type Information
                            
                            # Now remove the old address which was added as secondary automatically
                            Set-Mailbox -Identity $RecipientObject.Alias -EmailAddresses @{Remove = $address[1] } -ErrorAction Stop
                            Write-Log -Message "Removed $($address[1]) from $($RecipientObject.Alias)" -Type Information
                        }
                    } else {
                        # This is a secondary address - we can simply remove it
                        if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Remove secondary email address $($address[1])")) {
                            Set-Mailbox -Identity $RecipientObject.Alias -EmailAddresses @{Remove = $address[1] } -ErrorAction Stop
                            Write-Log -Message "Removed $($address[1]) from $($RecipientObject.Alias)" -Type Information
                        }
                    }
                } else {
                    # Handle inactive mailboxes
                    if ($address[0] -clike "SMTP") {
                        # This is a primary address - need to set a new primary first
                        $newPrimary = ($RecipientObject.Alias + "@" + $TenantDomain)
                        
                        if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Set new primary email address to $newPrimary (inactive mailbox)")) {
                            Set-Mailbox -Identity $RecipientObject.Alias -WindowsEmailAddress $newPrimary -InactiveMailbox:$True -ErrorAction Stop
                            Write-Log -Message "Changed primary address from $($RecipientObject.PrimarySmtpAddress) to $newPrimary (inactive mailbox)" -Type Information
                            
                            # Now remove the old address which was added as secondary automatically
                            Set-Mailbox -Identity $RecipientObject.Alias -InactiveMailbox:$True -EmailAddresses @{Remove = $address[1] } -ErrorAction Stop
                            Write-Log -Message "Removed $($address[1]) from $($RecipientObject.Alias) (inactive mailbox)" -Type Information
                        }
                    } else {
                        # This is a secondary address - we can simply remove it
                        if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Remove secondary email address $($address[1]) (inactive mailbox)")) {
                            Set-Mailbox -Identity $RecipientObject.Alias -InactiveMailbox:$True -EmailAddresses @{Remove = $address[1] } -ErrorAction Stop
                            Write-Log -Message "Removed $($address[1]) from $($RecipientObject.Alias) (inactive mailbox)" -Type Information
                        }
                    }
                }
            }
            "MailUser" {
                if ($address[0] -clike "SMTP") {
                    # This is a primary address - need to set a new primary first
                    $newPrimary = ($RecipientObject.Alias + "@" + $TenantDomain)
                    
                    if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Set new primary email address to $newPrimary (mail user)")) {
                        Set-MailUser -Identity $RecipientObject.Alias -WindowsEmailAddress $newPrimary -ErrorAction Stop
                        Write-Log -Message "Changed primary address from $($RecipientObject.PrimarySmtpAddress) to $newPrimary (mail user)" -Type Information
                        
                        # Now remove the old address which was added as secondary automatically
                        Set-MailUser -Identity $RecipientObject.Alias -EmailAddresses @{Remove = $address[1] } -ErrorAction Stop
                        Write-Log -Message "Removed $($address[1]) from $($RecipientObject.Alias) (mail user)" -Type Information
                    }
                } else {
                    # This is a secondary address - we can simply remove it
                    if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Remove secondary email address $($address[1]) (mail user)")) {
                        Set-MailUser -Identity $RecipientObject.Alias -EmailAddresses @{Remove = $address[1] } -ErrorAction Stop
                        Write-Log -Message "Removed $($address[1]) from $($RecipientObject.Alias) (mail user)" -Type Information
                    }
                }
            }
            { $_ -in @("MailUniversalDistributionGroup", "MailUniversalSecurityGroup") } {
                if ($address[0] -clike "SMTP") {
                    # This is a primary address - need to set a new primary first
                    $newPrimary = ($RecipientObject.Alias + "@" + $TenantDomain)
                    
                    if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Set new primary email address to $newPrimary (group)")) {
                        Set-DistributionGroup -Identity $RecipientObject.Alias -WindowsEmailAddress $newPrimary -ErrorAction Stop
                        Write-Log -Message "Changed primary address from $($RecipientObject.PrimarySmtpAddress) to $newPrimary (group)" -Type Information
                        
                        # Now remove the old address which was added as secondary automatically
                        Set-DistributionGroup -Identity $RecipientObject.Alias -EmailAddresses @{Remove = $address[1] } -ErrorAction Stop
                        Write-Log -Message "Removed $($address[1]) from $($RecipientObject.Alias) (group)" -Type Information
                    }
                } else {
                    # This is a secondary address - we can simply remove it
                    if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Remove secondary email address $($address[1]) (group)")) {
                        Set-DistributionGroup -Identity $RecipientObject.Alias -EmailAddresses @{Remove = $address[1] } -ErrorAction Stop
                        Write-Log -Message "Removed $($address[1]) from $($RecipientObject.Alias) (group)" -Type Information
                    }
                }
            }
            "GroupMailbox" {
                if ($address[0] -clike "SMTP") {
                    # This is a primary address - need to set a new primary first
                    $newPrimary = ($RecipientObject.Alias + "@" + $TenantDomain)
                    
                    if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Set new primary email address to $newPrimary (unified group)")) {
                        Set-UnifiedGroup -Identity $RecipientObject.Alias -PrimarySmtpAddress $newPrimary -ErrorAction Stop
                        Write-Log -Message "Changed primary address from $($RecipientObject.PrimarySmtpAddress) to $newPrimary (unified group)" -Type Information
                        
                        # Now remove the old address which was added as secondary automatically
                        Set-UnifiedGroup -Identity $RecipientObject.Alias -EmailAddresses @{Remove = $address[1] } -ErrorAction Stop
                        Write-Log -Message "Removed $($address[1]) from $($RecipientObject.Alias) (unified group)" -Type Information
                    }
                } else {
                    # This is a secondary address - we can simply remove it
                    if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Remove secondary email address $($address[1]) (unified group)")) {
                        Set-UnifiedGroup -Identity $RecipientObject.Alias -EmailAddresses @{Remove = $address[1] } -ErrorAction Stop
                        Write-Log -Message "Removed $($address[1]) from $($RecipientObject.Alias) (unified group)" -Type Information
                    }
                }
            }
            "RoomList" {
                if ($address[0] -clike "SMTP") {
                    # This is a primary address - need to set a new primary first
                    $newPrimary = ($RecipientObject.Alias + "@" + $TenantDomain)
                    
                    if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Set new primary email address to $newPrimary (room list)")) {
                        Set-DistributionGroup -Identity $RecipientObject.Alias -PrimarySmtpAddress $newPrimary -ErrorAction Stop
                        Write-Log -Message "Changed primary address from $($RecipientObject.PrimarySmtpAddress) to $newPrimary (room list)" -Type Information
                    }
                } else {
                    # This is a secondary address - we can simply remove it
                    if ($PSCmdlet.ShouldProcess($RecipientObject.PrimarySmtpAddress, "Remove secondary email address $($address[1]) (room list)")) {
                        Set-DistributionGroup -Identity $RecipientObject.Alias -EmailAddresses @{Remove = $address[1] } -ErrorAction Stop
                        Write-Log -Message "Removed $($address[1]) from $($RecipientObject.Alias) (room list)" -Type Information
                    }
                }
            }
            default {
                Write-Log -Message "Unknown recipient type: $($RecipientObject.RecipientTypeDetails) for $($RecipientObject.PrimarySmtpAddress)" -Type Warning
            }
        }
        return $true
    } catch {
        Write-Log -Message "Failed to remove $($address[1]) from $($RecipientObject.PrimarySmtpAddress): $($_.Exception.Message)" -Type Error
        return $false
    }
}
#endregion

#region Main Script Logic
try {
    # Initialize logging and announce script start
    Write-Log -Message "============ SCRIPT STARTED ============" -Type Information
    Write-Log -Message "Script running with Type: $Type, ObjectType: $ObjectType" -Type Information
    
    # Determine operation mode
    $operationMode = if ($DestroyMethod) { "Destroy" } else { "Discover" }
    Write-Log -Message "Operation mode: $operationMode" -Type Information
    
    # Connect to appropriate Exchange environment
    if ($Type -eq "Office365") {
        # Check and load Exchange Online and Azure AD modules
        Write-Log -Message "Checking required modules for Office 365..." -Type Information
        Import-RequiredModules -ModuleNames @("ExchangeOnlineManagement", "AzureAD")
        
        # Connect to Office 365 services
        Connect-Office365Services -CredentialPath $O365Admin
    } else {
        # Connect to Exchange On-Premises
        Write-Log -Message "Connecting to Exchange On-Premises..." -Type Information
        $exchangeCredential = $null
        if ($OrgMgmtAdmin) {
            try {
                $exchangeCredential = Import-Clixml -Path $OrgMgmtAdmin -ErrorAction Stop
                Write-Log -Message "Successfully loaded credentials from $OrgMgmtAdmin" -Type Information
            } catch {
                Write-Log -Message "Failed to load credentials from $OrgMgmtAdmin: $($_.Exception.Message)" -Type Error
                $exchangeCredential = Get-Credential -Message "Enter credentials for Exchange on-premises"
            }
        }
        
        Connect-ExchangeOnPremises -ExchangeServer $ExchangeServer -ExchangeCredential $exchangeCredential
    }
    
    # Process based on operation mode
    if ($operationMode -eq "Discover") {
        # Gather mode - collect addresses but don't remove them
        
        # Prompt for input file containing domain names
        Write-Log -Message "Prompting for input CSV file containing domain names..." -Type Information
        $filePath = Get-InputFilePath -StartDirectory $StartDirectory
        
        if (-not $filePath) {
            Write-Log -Message "No file was selected. Exiting script." -Type Warning
            exit
        }
        
        $fileDirectory = Split-Path -Path $filePath -Parent
        $fileName = Split-Path -Path $filePath -Leaf
        
        # Prepare output file
        $outputFile = Join-Path -Path $fileDirectory -ChildPath "$($fileName.Substring(0, $fileName.Length - 4))_Domain-Results-$currentDate.txt"
        
        if (Test-Path -Path $outputFile) {
            Remove-Item -Path $outputFile -Force
        }
        
        Add-Content -Path $outputFile -Value "Guid`tExchangeGuid`tExchangeObjectId`tRecipientTypeDetails`tIsInactiveMailbox`tAlias`tEmailAddress`tPrimarySmtpAddress`tDisplayName`tLegacyExchangeDN"
        
        # Process each domain in the input file
        $domains = Import-Csv -Path $filePath
        $domainCount = $domains.Count
        
        Write-Log -Message "Found $domainCount domains to process" -Type Information
        
        $progressCounter = 0
        foreach ($domain in $domains) {
            $progressCounter++
            Write-Progress -Id 1 -Activity "Processing domains" -Status "Domain $progressCounter of $domainCount" -PercentComplete (($progressCounter / $domainCount) * 100)
            
            Write-Log -Message "Processing domain: $($domain.DomainName)" -Type Information
            
            # Execute appropriate functions based on ObjectType
            if ($ObjectType -eq "MailboxOnly" -or $ObjectType -eq "All") {
                Get-MailboxDetail -Domain $domain -OutputFile $outputFile -IncludeLegacyExchangeDN:$IncludeLegacyExchangeDN
                Get-InactiveMailboxDetail -Domain $domain -OutputFile $outputFile -IncludeLegacyExchangeDN:$IncludeLegacyExchangeDN
                Get-MailUserDetail -Domain $domain -OutputFile $outputFile -IncludeLegacyExchangeDN:$IncludeLegacyExchangeDN
            }
            
            if ($ObjectType -eq "GroupsOnly" -or $ObjectType -eq "All") {
                Get-DistributionGroupDetail -Domain $domain -OutputFile $outputFile -IncludeLegacyExchangeDN:$IncludeLegacyExchangeDN
                Get-RoomListDetail -Domain $domain -OutputFile $outputFile -IncludeLegacyExchangeDN:$IncludeLegacyExchangeDN
                
                # Unified (Office 365) Groups only exist in Exchange Online
                if ($Type -ne "On-Premises") {
                    Get-UnifiedGroupDetail -Domain $domain -OutputFile $outputFile -IncludeLegacyExchangeDN:$IncludeLegacyExchangeDN
                }
            }
        }
        
        Write-Progress -Id 1 -Activity "Processing domains" -Completed
        Write-Log -Message "Domain processing completed. Results saved to: $outputFile" -Type Information
        
        # Return the output file path for use by other scripts
        return $outputFile
    } else {
        # Destroy mode - remove addresses
        
        # Source the list of addresses to remove
        if ($DestroyMethod -eq "Destroy") {
            # Get the list from an existing output file
            Write-Log -Message "Prompting for input file containing addresses to remove..." -Type Information
            $inputFile = Get-InputFilePath -StartDirectory $StartDirectory
            
            if (-not $inputFile) {
                Write-Log -Message "No file was selected. Exiting script." -Type Warning
                exit
            }
            
            try {
                $users = Import-Csv -Path $inputFile -Delimiter "`t"
                Write-Log -Message "Loaded $($users.Count) addresses to process from $inputFile" -Type Information
            } catch {
                Write-Log -Message "Error loading address file: $($_.Exception.Message)" -Type Error
                exit
            }
        } elseif ($DestroyMethod -eq "DestroyOnly") {
            # Prompt for a CSV from a previous run
            Write-Log -Message "Prompting for previously generated address file..." -Type Information
            $inputFile = Get-InputFilePath -StartDirectory $StartDirectory
            
            if (-not $inputFile) {
                Write-Log -Message "No file was selected. Exiting script." -Type Warning
                exit
            }
            
            try {
                $users = Import-Csv -Path $inputFile -Delimiter "`t"
                Write-Log -Message "Loaded $($users.Count) addresses to process from $inputFile" -Type Information
            } catch {
                Write-Log -Message "Error loading address file: $($_.Exception.Message)" -Type Error
                exit
            }
        }
        
        # Get tenant domain for creating replacement addresses
        $tenantDomain = ""
        if ($Type -eq "Office365") {
            try {
                $tenantDomain = Get-AzureADDomain | Where-Object { $_.IsInitial -eq $true }
                Write-Log -Message "Found tenant domain: $($tenantDomain.Name)" -Type Information
            } catch {
                Write-Log -Message "Error retrieving tenant domain: $($_.Exception.Message)" -Type Error
                Write-Log -Message "Using a placeholder domain for new addresses. Update manually if needed." -Type Warning
                $tenantDomain = @{ Name = "yourtenant.onmicrosoft.com" }
            }
        } else {
            # For on-premises, try to determine the default domain
            try {
                $acceptedDomain = Get-AcceptedDomain | Where-Object { $_.Default -eq $true }
                if ($acceptedDomain) {
                    $tenantDomain = @{ Name = $acceptedDomain.DomainName.ToString() }
                    Write-Log -Message "Found default accepted domain: $($tenantDomain.Name)" -Type Information
                } else {
                    Write-Log -Message "Could not determine default domain. Will prompt for input." -Type Warning
                    $defaultDomain = Read-Host "Enter the default domain name to use for new primary addresses"
                    $tenantDomain = @{ Name = $defaultDomain }
                }
            } catch {
                Write-Log -Message "Error retrieving accepted domains: $($_.Exception.Message)" -Type Error
                $defaultDomain = Read-Host "Enter the default domain name to use for new primary addresses"
                $tenantDomain = @{ Name = $defaultDomain }
            }
        }
        
        # Process address removal
        $progressCounter = 0
        $successCount = 0
        $failureCount = 0
        
        foreach ($user in $users) {
            $progressCounter++
            Write-Progress -Id 1 -Activity "Processing addresses" -Status "Address $progressCounter of $($users.Count)" -PercentComplete (($progressCounter / $users.Count) * 100)
            
            # Skip entries without email addresses (might be LegacyExchangeDN entries)
            if ([string]::IsNullOrEmpty($user.EmailAddress)) {
                Write-Log -Message "Skipping entry without email address: $($user.PrimarySmtpAddress)" -Type Warning
                continue
            }
            
            # Get the actual recipient object for processing
            try {
                $recipient = $null
                
                # Different retrieval methods based on recipient type
                switch ($user.RecipientTypeDetails) {
                    { $_ -in @("UserMailbox", "SharedMailbox", "RoomMailbox") } {
                        if ($user.IsInactiveMailbox -eq $true) {
                            $recipient = Get-Mailbox -Identity $user.Alias -InactiveMailboxOnly -ErrorAction Stop
                        } else {
                            $recipient = Get-Mailbox -Identity $user.Alias -ErrorAction Stop
                        }
                    }
                    "MailUser" {
                        $recipient = Get-MailUser -Identity $user.Alias -ErrorAction Stop
                    }
                    { $_ -in @("MailUniversalDistributionGroup", "MailUniversalSecurityGroup", "RoomList") } {
                        $recipient = Get-DistributionGroup -Identity $user.Alias -ErrorAction Stop
                    }
                    "GroupMailbox" {
                        $recipient = Get-UnifiedGroup -Identity $user.Alias -ErrorAction Stop
                    }
                    default {
                        Write-Log -Message "Unknown recipient type: $($user.RecipientTypeDetails) for $($user.PrimarySmtpAddress)" -Type Warning
                        $failureCount++
                        continue
                    }
                }
                
                if ($null -eq $recipient) {
                    Write-Log -Message "Could not find recipient $($user.PrimarySmtpAddress)" -Type Warning
                    $failureCount++
                    continue
                }
                
                # Perform the address removal
                $result = Remove-EmailAddress -RecipientObject $recipient -AddressToRemove $user.EmailAddress -TenantDomain $tenantDomain.Name
                
                if ($result) {
                    $successCount++
                } else {
                    $failureCount++
                }
            } catch {
                Write-Log -Message "Error processing $($user.PrimarySmtpAddress): $($_.Exception.Message)" -Type Error
                $failureCount++
                continue
            }
        }
        
        Write-Progress -Id 1 -Activity "Processing addresses" -Completed
        
        # Summarize results
        Write-Log -Message "Address removal processing completed" -Type Information
        Write-Log -Message "Successfully processed $successCount addresses" -Type Information
        
        if ($failureCount -gt 0) {
            Write-Log -Message "Failed to process $failureCount addresses" -Type Warning
        }
    }
} catch {
    Write-Log -Message "Critical error: $($_.Exception.Message)" -Type Error
    Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Type Error
} finally {
    # Cleanup and disconnect sessions
    try {
        Get-PSSession | Remove-PSSession
        Write-Log -Message "Disconnected all PowerShell sessions" -Type Information
    } catch {
        Write-Log -Message "Error during session cleanup: $($_.Exception.Message)" -Type Warning
    }
    
    Write-Log -Message "============ SCRIPT COMPLETED ============" -Type Information
}
#endregion