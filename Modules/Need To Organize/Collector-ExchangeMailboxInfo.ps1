#Requires -version 5.0
#Requires -Modules Sterling

<#
    .SYNOPSIS
        Collects mailbox data from Exchange (on-premises or online) environments to be exported as various formats.
        
        This script cannot be ran automatically in an environment that requires MFA. The account used
        must bypass the MFA requirement.
        
        In Office 365, Global Reader DOES NOT WORK!
            
    .DESCRIPTION
        This script contains proprietary information owned by Core BTS (“Core BTS”) and should be 
        regarded as confidential. This script, any output and files, related information, and all copies of same remain the
        confidential property of Core BTS and shall be returned to Core BTS upon request.
        
        These materials and the information contained herein are not to be duplicated or used, in whole or in part, for 
        any purpose other than for the purposes for Core BTS and their clients.
        
        Author: 
        - Ryan Weaver, Principal Architect, ryan.weaver@corebts.com

        Requirements: 
        - PowerShell 5.0+
        - Exchange 2010/2013/2016/2019/Online Management Shell
        - Assuming no RBAC limiting functionality, Org Management Role
        
    .PARAMETER ExchangeOrgName
        This parameter is deprecated and no longer used.

    .PARAMETER MailboxFilter
        Specify the filter to be used for the mailboxes to collect. It will default to "'*'" if not specified.
        
        If Exchange 2010, '*' does not work and a filter must be specified. If all mailboxes are desired, 
        use something like: -MailboxFilter \"DisplayName -like '*'\"

    .PARAMETER ObjectDomain
        Specify the mailbox domain name to be used in the output. This should be the Active Directory Forest or Office 365 Tenant name.

    .PARAMETER StageBatchSize
        Specify only if a different staging threshold should be used for pushing changes to SQL. It will default to every 5000 records if not specified.

    .PARAMETER IncludeGroupMailboxes
        Specify only if the returned data should include Office 365 group mailboxes. It will default to not being included if not specified.

    .PARAMETER IncludeCASInfo
        Specify only if the returned data should include CAS mailbox information/settings. It will default to not being included if not specified.

    .PARAMETER IncludeMobileInfo
        Specify only if the returned data should include mobile device information/settings. It will default to not being included if not specified.

    .PARAMETER IncludeMobileStats
        Specify only if the returned data should include mobile device statistics, this will include mobile device information/settings. It will default to not being included if not specified

    .PARAMETER IncludeInboxRules
        Specify only if the returned data should include the count of inbox rules. It will default to not being included if not specified.

    .PARAMETER IncludeRegionalInfo
        Specify only if the returned data should include the regional mailbox configuration. It will default to not being included if not specified.
        NOT RECOMMENDED FOR LARGE ENVIRONMENTS AND EXCHANGE ONLINE. IT IS VERY SLOW CURRENTLY.

    .PARAMETER SessionTimeout
        Specify only if a different Exchange session timeout. It will default to every 60 minutes if not specified.

    .PARAMETER Help
        Specify to display help content

    .EXAMPLE
        Collector-ExchangeMailboxInfo.ps1 -ObjectDomain 'contoso.onmicrosoft.com' -MailboxFilter 'HiddenFromAddressListsEnabled -eq $true'

        This would run the mailbox info collector using the Migrator for saved credentials of contoso.onmicrosoft.com
        and filter the mailboxes to those that are hidden from the address list

    .EXAMPLE
        Collector-ExchangeMailboxInfo.ps1 -ObjectDomain contoso.onmicrosoft.com

        This would run the mailbox info collector using the Migrator for saved credentials of contoso.onmicrosoft.com,
        and export the results to the Migrator

    .NOTES
        Version:
            - 5.1.2024.0516:    New script
#>

#region Parameters
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify FQDN of an available on-premises Exchange server. It will default to 'Online' if not specified.")]
    [string]$ExchangeServer = "Online",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the file path for the Exchange Server credentials.")]
    [string]$ExchangeCredentialPath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify only if you need to force SSL on for the remote PowerShell connection. It will default to false if not specified.")]
    [switch]$ExchangeForceSSL,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the remote PowerShell the specific authentication method. It will default to Basic if not specified.")]
    [ValidateSet("Basic", "Negotiate", "Kerberos")]
    [string]$ExchangeAuthMethod = "Basic",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the export format to use. It will default to DB if not specified.")]
    [ValidateSet("CSV", "TSV", "CLIXML", "JSON")]
    [string]$ExportFormat = "DB",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the MySQL server name.")]
    [string]$MySQLServer,
	
    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify if specific MySQL port is needed. It will default to 3306 if not specified.")]
    [int]$MySQLPort = 3306,
			
    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the MySQL database name.")]
    [string]$Database,
	
    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify if specific MySQL credentials are needed.")]
    [string]$MySQLCredentialFile,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify a delimiter to use for CSV file. It will default to comma separated if not specified.")]
    [char]$Delimiter = ",",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the export location to use. It will default to the current path if not specified.")]
    [string]$DataFolder = "",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the filter to be used for the mailboxes to collect. It will default to ''*'' if not specified.")]
    [string]$MailboxFilter = "*",

    [Parameter(Mandatory = $true, ValueFromPipeline = $false,
        HelpMessage = "Specify the mailbox domain name to be used in the output.")]
    [string]$ObjectDomain,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify only if the returned data should include Office 365 group mailboxes. It will default to not being included if not specified.")]
    [switch]$IncludeGroupMailboxes,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify only if the returned data should include CAS mailbox information/settings. It will default to not being included if not specified.")]
    [switch]$IncludeCASInfo,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify only if the returned data should include mobile device information/settings. It will default to not being included if not specified.")]
    [switch]$IncludeMobileInfo,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify only if the returned data should include mobile device statistics, this will include mobile device information/settings. It will default to not being included if not specified.")]
    [switch]$IncludeMobileStats,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify only if the returned data should include the count of inbox rules. It will default to not being included if not specified.")]
    [switch]$IncludeInboxRules,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify only if the returned data should include the regional mailbox configuration. It will default to not being included if not specified.")]
    [switch]$IncludeRegionalInfo
)
#endregion

#region User Variables
#Very little reason to change these
$InformationPreference = "Continue"

if ($DebugPreference -eq "Confirm" -or $DebugPreference -eq "Inquire") { $DebugPreference = "Continue" }
#endregion

#region Static Variables
#Don't change these
Set-Variable -Name strBaseLocation -Option AllScope -Scope Script
Set-Variable -Name dateStartTimeStamp -Option AllScope -Scope Script -Value (Get-Date).ToUniversalTime()
Set-Variable -Name strLogTimeStamp -Option AllScope -Scope Script -Value $dateStartTimeStamp.ToString("MMddyyyy_HHmmss")
Set-Variable -Name strLogFile -Option ReadOnly -Scope Script
Set-Variable -Name htLoggingPreference -Option AllScope -Scope Script -Value @{"InformationPreference" = $InformationPreference; `
        "WarningPreference" = $WarningPreference; "ErrorActionPreference" = $ErrorActionPreference; `
        "VerbosePreference" = $VerbosePreference; "DebugPreference" = $DebugPreference
}
Set-Variable -Name verScript -Option AllScope -Scope Script -Value "5.1.2024.0516"

Set-Variable -Name boolScriptIsModulesLoaded -Option AllScope -Scope Script -Value $false
Set-Variable -Name boolScriptIsEMS -Option AllScope -Scope Script -Value $false
Set-Variable -Name boolScriptIsEMSOnline -Option AllScope -Scope Script -Value $false
Set-Variable -Name ExitCode -Option AllScope -Scope Script -Value 1

Set-Variable -Name MySQLMailboxTable -Option AllScope -Scope Script -Value "ex_mailboxes"
Set-Variable -Name MySQLEmailAddressesTable -Option AllScope -Scope Script -Value "ex_mailboxes_mva_emailaddresses"
Set-Variable -Name MySQLExtensionCustomAttribute1Table -Option AllScope -Scope Script -Value "ex_mailboxes_mva_extensioncustomattribute1"
Set-Variable -Name MySQLExtensionCustomAttribute2Table -Option AllScope -Scope Script -Value "ex_mailboxes_mva_extensioncustomattribute2"
Set-Variable -Name MySQLExtensionCustomAttribute3Table -Option AllScope -Scope Script -Value "ex_mailboxes_mva_extensioncustomattribute3"
Set-Variable -Name MySQLExtensionCustomAttribute4Table -Option AllScope -Scope Script -Value "ex_mailboxes_mva_extensioncustomattribute4"
Set-Variable -Name MySQLExtensionCustomAttribute5Table -Option AllScope -Scope Script -Value "ex_mailboxes_mva_extensioncustomattribute5"
Set-Variable -Name MySQLMobileTable -Option AllScope -Scope Script -Value "ex_mobiledevices"
Set-Variable -Name MySQLStagingSize -Option AllScope -Scope Script -Value 100
Set-Variable -Name MySQLCredential -Option AllScope -Scope Script
Set-Variable -Name ExportLocation -Option AllScope -Scope Script

New-Object System.Data.DataTable | Set-Variable dtEmailAddresses -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute1 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute2 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute3 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute4 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute5 -Option AllScope -Scope Script

Set-Variable -Name arrMailboxAttribs -Option AllScope -Scope Script -Value 'Guid', 'ExchangeGuid', 'SamAccountName', 'UserPrincipalName', 'RecipientType', 'Alias', `
    'ArchiveGuid', 'ArchiveStatus', 'CustomAttribute1', 'CustomAttribute10', 'CustomAttribute11', 'CustomAttribute12', 'CustomAttribute13', `
    'CustomAttribute14', 'CustomAttribute15', 'CustomAttribute2', 'CustomAttribute3', 'CustomAttribute4', 'CustomAttribute5', 'CustomAttribute6', `
    'CustomAttribute7', 'CustomAttribute8', 'CustomAttribute9', 'DisplayName', 'ForwardingAddress', 'ForwardingSmtpAddress', 'HiddenFromAddressListsEnabled', `
    'LitigationHoldEnabled', 'LitigationHoldOwner', 'MaxReceiveSize', 'MaxSendSize', 'PrimarySmtpAddress', 'RecipientTypeDetails', `
    'MailboxGuid', 'EmailAddresses', 'IsDirSynced', 'DistinguishedName', 'EmailAddressPolicyEnabled', 'EndDateForRetentionHoldUTC', `
    'Identity', 'IsInactiveMailbox', 'LegacyExchangeDN', 'LinkedMasterAccount', 'LitigationHoldDate', 'LitigationHoldDuration', 'MailboxMoveBatchName', `
    'MailboxMoveStatus', 'MailboxPlan', 'MailTip', 'MessageCopyForSendOnBehalfEnabled', 'MessageCopyForSentAsEnabled', 'MessageTrackingReadStatusEnabled', `
    'ModerationEnabled', 'Name', 'Office', 'OrganizationalUnit', 'ProhibitSendQuota', 'ProhibitSendReceiveQuota', 'RemoteRecipientType', 'ResourceCapacity', `
    'ResourceType', 'RetentionComment', 'RetentionHoldEnabled', 'RetentionPolicy', 'RetentionUrl', 'StartDateForRetentionHoldUTC', 'UsageLocation', `
    'WhenChangedUTC', 'WhenCreatedUTC', 'WhenMailboxCreated', 'WhenSoftDeletedUTC', 'WhenAddedUTC', 'WhenModifiedUTC', 'AccountDisabled', `
    'IsDeleted', 'ExternalDirectoryObjectId', 'MailboxRegion', 'AutoExpandingArchiveEnabled', 'ImmutableID', 'SourceAnchor', 'ArchiveQuota' 

Set-Variable -Name arrMailboxCASAttribs -Option AllScope -Scope Script -Value 'Guid', 'OWAEnabled', 'PopEnabled', 'ImapEnabled', 'MAPIEnabled'
Set-Variable -Name arrMobileAttribs -Option AllScope -Scope Script -Value 'Guid', 'FriendlyName', 'DeviceId', 'DeviceImei', 'DeviceMobileOperator', `
    'DeviceOS', 'DeviceOSLanguage', 'DeviceTelephoneNumber', 'DeviceType', 'DeviceUserAgent', 'DeviceModel', 'FirstSyncTime', 'UserDisplayName', 'DeviceAccessState', `
    'DeviceAccessStateReason', 'DeviceAccessControlRule', 'ClientVersion', 'ClientType', 'IsManaged', 'IsCompliant', 'IsDisabled', 'ExchangeVersion', `
    'Name', 'DistinguishedName', 'Identity', 'WhenChangedUTC', 'WhenCreatedUTC', 'ExchangeGuid'

Set-Variable -Name arrMobileStatsArribs -Option AllScope -Scope Script -Value 'LastPolicyUpdateTime', 'LastSyncAttemptTime', `
    'LastSuccessSync', 'DeviceWipeSentTime', 'DeviceWipeRequestTime', 'DeviceWipeAckTime', 'AccountOnlyDeviceWipeSentTime', `
    'AccountOnlyDeviceWipeRequestTime', 'AccountOnlyDeviceWipeAckTime', 'LastPingHeartbeat', 'MailboxLogReport', 'DeviceEnableOutboundSMS', `
    'Guid', 'IsRemoteWipeSupported', 'Status', 'StatusNote', 'DevicePolicyApplied', 'DevicePolicyApplicationStatus', `
    'LastDeviceWipeRequestor', 'LastAccountOnlyDeviceWipeRequestor', 'NumberOfFoldersSynced', 'SyncStateUpgradeTime'

Set-Variable -Name MailboxPropertyMap -Option AllScope -Scope Script -Value @{
    "ObjectDomain"                      = "ObjectDomain"
    "SamAccountName"                    = "SamAccountName"
    "UserPrincipalName"                 = "UserPrincipalName"
    "RecipientType"                     = "RecipientType"
    "Alias"                             = "Alias"
    "ArchiveGuid"                       = "ArchiveGuid"
    "ArchiveStatus"                     = "ArchiveStatus"
    "ArchiveQuota"                      = "ArchiveQuotaMB"
    "CustomAttribute1"                  = "CustomAttribute1"
    "CustomAttribute10"                 = "CustomAttribute10"
    "CustomAttribute11"                 = "CustomAttribute11"
    "CustomAttribute12"                 = "CustomAttribute12"
    "CustomAttribute13"                 = "CustomAttribute13"
    "CustomAttribute14"                 = "CustomAttribute14"
    "CustomAttribute15"                 = "CustomAttribute15"
    "CustomAttribute2"                  = "CustomAttribute2"
    "CustomAttribute3"                  = "CustomAttribute3"
    "CustomAttribute4"                  = "CustomAttribute4"
    "CustomAttribute5"                  = "CustomAttribute5"
    "CustomAttribute6"                  = "CustomAttribute6"
    "CustomAttribute7"                  = "CustomAttribute7"
    "CustomAttribute8"                  = "CustomAttribute8"
    "CustomAttribute9"                  = "CustomAttribute9"
    "DisplayName"                       = "DisplayName"
    "ForwardingAddress"                 = "ForwardingAddress"
    "ForwardingSmtpAddress"             = "ForwardingSmtpAddress"
    "HiddenFromAddressListsEnabled"     = "HiddenFromAddressListsEnabled"
    "LitigationHoldEnabled"             = "LitigationHoldEnabled"
    "LitigationHoldOwner"               = "LitigationHoldOwner"
    "MaxReceiveSize"                    = "MaxReceiveSizeMB"
    "MaxSendSize"                       = "MaxSendSizeMB"
    "PrimarySmtpAddress"                = "PrimarySmtpAddress"
    "RecipientTypeDetails"              = "RecipientTypeDetails"
    "MailboxGuid"                       = "MailboxGuid"
    "IsDirSynced"                       = "IsDirSynced"
    "ExchangeGuid"                      = "ExchangeGuid"
    "DistinguishedName"                 = "DistinguishedName"
    "EmailAddressPolicyEnabled"         = "EmailAddressPolicyEnabled"
    "EndDateForRetentionHoldUTC"        = "EndDateForRetentionHoldUTC"
    "Guid"                              = "Guid"
    "Identity"                          = "Identity"
    "IsInactiveMailbox"                 = "IsInactiveMailbox"
    "LegacyExchangeDN"                  = "LegacyExchangeDN"
    "LinkedMasterAccount"               = "LinkedMasterAccount"
    "LitigationHoldDate"                = "LitigationHoldDateUTC"
    "LitigationHoldDuration"            = "LitigationHoldDuration"
    "MailboxMoveBatchName"              = "MailboxMoveBatchName"
    "MailboxMoveStatus"                 = "MailboxMoveStatus"
    "MailboxPlan"                       = "MailboxPlan"
    "MailTip"                           = "MailTip"
    "MessageCopyForSendOnBehalfEnabled" = "MessageCopyForSendOnBehalfEnabled"
    "MessageCopyForSentAsEnabled"       = "MessageCopyForSentAsEnabled"
    "MessageTrackingReadStatusEnabled"  = "MessageTrackingReadStatusEnabled"
    "ModerationEnabled"                 = "ModerationEnabled"
    "Name"                              = "Name"
    "Office"                            = "Office"
    "OrganizationalUnit"                = "OrganizationalUnit"
    "ProhibitSendQuota"                 = "ProhibitSendQuotaMB"
    "ProhibitSendReceiveQuota"          = "ProhibitSendReceiveQuotaMB"
    "RemoteRecipientType"               = "RemoteRecipientType"
    "ResourceCapacity"                  = "ResourceCapacity"
    "ResourceType"                      = "ResourceType"
    "RetentionComment"                  = "RetentionComment"
    "RetentionHoldEnabled"              = "RetentionHoldEnabled"
    "RetentionPolicy"                   = "RetentionPolicy"
    "RetentionUrl"                      = "RetentionUrl"
    "StartDateForRetentionHoldUTC"      = "StartDateForRetentionHoldUTC"
    "UsageLocation"                     = "UsageLocation"
    "WhenChangedUTC"                    = "WhenChangedUTC"
    "WhenCreatedUTC"                    = "WhenCreatedUTC"
    "WhenMailboxCreated"                = "WhenMailboxCreatedUTC"
    "WhenSoftDeletedUTC"                = "WhenSoftDeletedUTC"
    "WhenAddedUTC"                      = "WhenAddedUTC"
    "WhenModifiedUTC"                   = "WhenModifiedUTC"
    "AccountDisabled"                   = "AccountDisabled"
    "IsDeleted"                         = "IsDeleted"
    "ExternalDirectoryObjectId"         = "ExternalDirectoryObjectId"
    "MailboxRegion"                     = "MailboxRegion"
    "AutoExpandingArchiveEnabled"       = "AutoExpandingArchiveEnabled"
    "ImmutableID"                       = "ImmutableID"
    "SourceAnchor"                      = "SourceAnchor"
}

Set-Variable -Name EmailAddressesPropertyMap -Option AllScope -Scope Script -Value @{
    "ObjectDomain"   = "ObjectDomain"
    "ExchangeGuid"   = "ExchangeGuid"
    "EmailAddresses" = "EmailAddresses"
}

Set-Variable -Name ExtensionCustomAttribute1PropertyMap -Option AllScope -Scope Script -Value @{
    "ObjectDomain"              = "ObjectDomain"
    "ExchangeGuid"              = "ExchangeGuid"
    "ExtensionCustomAttribute1" = "ExtensionCustomAttribute1"
}

Set-Variable -Name ExtensionCustomAttribute2PropertyMap -Option AllScope -Scope Script -Value @{
    "ObjectDomain"              = "ObjectDomain"
    "ExchangeGuid"              = "ExchangeGuid"
    "ExtensionCustomAttribute2" = "ExtensionCustomAttribute2"
}

Set-Variable -Name ExtensionCustomAttribute3PropertyMap -Option AllScope -Scope Script -Value @{
    "ObjectDomain"              = "ObjectDomain"
    "ExchangeGuid"              = "ExchangeGuid"
    "ExtensionCustomAttribute3" = "ExtensionCustomAttribute3"
}

Set-Variable -Name ExtensionCustomAttribute4PropertyMap -Option AllScope -Scope Script -Value @{
    "ObjectDomain"              = "ObjectDomain"
    "ExchangeGuid"              = "ExchangeGuid"
    "ExtensionCustomAttribute4" = "ExtensionCustomAttribute4"
}

Set-Variable -Name ExtensionCustomAttribute5PropertyMap -Option AllScope -Scope Script -Value @{
    "ObjectDomain"              = "ObjectDomain"
    "ExchangeGuid"              = "ExchangeGuid"
    "ExtensionCustomAttribute5" = "ExtensionCustomAttribute5"
}
                                       
Set-Variable -Name CASPropertyMap -Option AllScope -Scope Script -Value @{
    "OWAEnabled"  = "OWAEnabled"
    "PopEnabled"  = "PopEnabled"
    "ImapEnabled" = "ImapEnabled"
    "MAPIEnabled" = "MAPIEnabled"
}

Set-Variable -Name InboxRulePropertyMap -Option AllScope -Scope Script -Value @{"InboxRuleCount" = "InboxRuleCount" }
Set-Variable -Name RegionalPropertyMap -Option AllScope -Scope Script -Value @{
    "Language" = "Language"
    "TimeZone" = "TimeZone"
}
Set-Variable -Name MobilePropertyMap -Option AllScope -Scope Script -Value @{
    "MobileDomain"                       = "MobileDomain"
    "Guid"                               = "Guid"
    "ExchangeGuid"                       = "ExchangeGuid"
    "LastPolicyUpdateTime"               = "LastPolicyUpdateTimeUTC"
    "LastSyncAttemptTime"                = "LastSyncAttemptTimeUTC"
    "LastSuccessSync"                    = "LastSuccessSyncUTC"
    "DeviceWipeSentTime"                 = "DeviceWipeSentTimeUTC"
    "DeviceWipeRequestTime"              = "DeviceWipeRequestTimeUTC"
    "DeviceWipeAckTime"                  = "DeviceWipeAckTimeUTC"
    "AccountOnlyDeviceWipeSentTime"      = "AccountOnlyDeviceWipeSentTimeUTC"
    "AccountOnlyDeviceWipeRequestTime"   = "AccountOnlyDeviceWipeRequestTimeUTC"
    "AccountOnlyDeviceWipeAckTime"       = "AccountOnlyDeviceWipeAckTimeUTC"
    "LastPingHeartbeat"                  = "LastPingHeartbeat"
    "MailboxLogReport"                   = "MailboxLogReport"
    "DeviceEnableOutboundSMS"            = "DeviceEnableOutboundSMS"
    "IsRemoteWipeSupported"              = "IsRemoteWipeSupported"
    "Status"                             = "Status"
    "StatusNote"                         = "StatusNote"
    "DevicePolicyApplied"                = "DevicePolicyApplied"
    "DevicePolicyApplicationStatus"      = "DevicePolicyApplicationStatus"
    "LastDeviceWipeRequestor"            = "LastDeviceWipeRequestor"
    "LastAccountOnlyDeviceWipeRequestor" = "LastAccountOnlyDeviceWipeRequestor"
    "NumberOfFoldersSynced"              = "NumberOfFoldersSynced"
    "SyncStateUpgradeTime"               = "SyncStateUpgradeTimeUTC"
    "FriendlyName"                       = "FriendlyName"
    "DeviceId"                           = "DeviceId"
    "DeviceImei"                         = "DeviceImei"
    "DeviceMobileOperator"               = "DeviceMobileOperator"
    "DeviceOS"                           = "DeviceOS"
    "DeviceOSLanguage"                   = "DeviceOSLanguage"
    "DeviceTelephoneNumber"              = "DeviceTelephoneNumber"
    "DeviceType"                         = "DeviceType"
    "DeviceUserAgent"                    = "DeviceUserAgent"
    "DeviceModel"                        = "DeviceModel"
    "FirstSyncTime"                      = "FirstSyncTimeUTC"
    "UserDisplayName"                    = "UserDisplayName"
    "DeviceAccessState"                  = "DeviceAccessState"
    "DeviceAccessStateReason"            = "DeviceAccessStateReason"
    "DeviceAccessControlRule"            = "DeviceAccessControlRule"
    "ClientVersion"                      = "ClientVersion"
    "ClientType"                         = "ClientType"
    "IsManaged"                          = "IsManaged"
    "IsCompliant"                        = "IsCompliant"
    "IsDisabled"                         = "IsDisabled"
    "ExchangeVersion"                    = "ExchangeVersion"
    "Name"                               = "Name"
    "DistinguishedName"                  = "DistinguishedName"
    "Identity"                           = "Identity"
    "WhenChangedUTC"                     = "WhenChangedUTC"
    "WhenCreatedUTC"                     = "WhenCreatedUTC"
}
    
#endregion

#region Complete Functions
Function _ConfirmScriptRequirements {
    <#
    .SYNOPSIS
        Verifies that all necessary requirements are present for the script and return true/false
    .EXAMPLE
        $valid = _ConfirmScriptRequirements

        This would check the script requirements and set $valid to true/false based on the results
    .NOTES
        Version:
        - 5.1.2024.0516:    New function
    #>
    [CmdletBinding()]
    Param()

    begin {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        Write-Debug -Message "Starting _ConfirmScriptRequirements"
        try {
            Write-Host "Loading Sterling PowerShell module`r"

            if (Get-Module -ListAvailable Sterling -Verbose:$false) {
                Import-Module Sterling -ErrorAction Stop -Verbose:$false
                $script:boolScriptIsModulesLoaded = $true
            } else {
                Write-Warning "Missing Sterlin PowerShell module`r"
                $script:boolScriptIsModulesLoaded = $false
            }#if/else
        } catch {
            Write-Error "Unable to load Sterling PowerShell module`r"
            Write-Error $_

            $script:boolScriptIsModulesLoaded = $false
        }#try/catch

        Set-Variable -Name strBaseLocation -Option AllScope -Scope Script -Value $(_GetScriptDirectory -Path)
        Set-Variable -Name strLogFile -Option ReadOnly -Force -Scope Script -Value "$script:strBaseLocation\Logging\$script:strLogTimeStamp-$((_GetScriptDirectory -Leaf).Replace(".ps1",'')).log"

        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Debug -WriteBackToHost -Message "Starting _ConfirmScriptRequirements"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Script version $verScript starting"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: InformationPreference = $InformationPreference"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ErrorActionPreference = $ErrorActionPreference"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: VerbosePreference = $VerbosePreference"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: DebugPreference = $DebugPreference"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ExchangeServer = $ExchangeServer"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ExchangeForceSSL = $ExchangeForceSSL"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ExchangeAuthMethod = $ExchangeAuthMethod"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ExchangeCredentialPath = $ExchangeCredentialPath"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: MailboxFilter = $MailboxFilter"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ObjectDomain = $ObjectDomain"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ExportFormat = $ExportFormat"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: Delimiter = $Delimiter"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: DataFolder = $DataFolder"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: MailboxFilter = $MailboxFilter"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludeCASInfo = $IncludeCASInfo"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludeMobileInfo = $IncludeMobileInfo"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludeMobileStats = $IncludeMobileStats"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludeInboxRules = $IncludeInboxRules"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludeRegionalInfo = $IncludeRegionalInfo"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludeGroupMailboxes = $IncludeGroupMailboxes"
    }#begin
    
    process {
        if ($script:boolScriptIsModulesLoaded) {
            try {
                if ($ExportFormat -eq "DB") {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: MySQLServer = $MySQLServer"
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: MySQLPort = $MySQLPort"
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: Database = $Database"

                    if ((Test-Path -Path $MySQLCredentialFile)) {
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: MySQLCredentialFile = $MySQLCredentialFile"
                        
                        $MySQLCredential = Import-Clixml $MySQLCredentialFile
                    } else {
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "MySQLCredentialFile specified but does not exist. Prompting for credentails"
                        $MySQLCredential = Get-Credential
                    }
                }

                $global:VerbosePreference = "SilentlyContinue"
                
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Connecting to Exchange"
                $ExchangePSSessionID = _ConfirmExchangeSession -Server $ExchangeServer -ForceSSL:$ExchangeForceSSL -AuthMethod $ExchangeAuthMethod -Creds $ExchangeCreds -CredentialFile $ExchangeCredentialPath

                if ($htLoggingPreference['VerbosePreference'] -eq "Continue") { $global:VerbosePreference = "Continue" }#if
                
                if ($DataFolder -eq "") {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "DataFolder not supplied, setting to current script location"
                    $script:ExportLocation = $strBaseLocation + "\Exchange"
                } else { $script:ExportLocation = $DataFolder }
            } catch {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error verifying script requirements"
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
                return $false
            }#try/catch
        }#if

        #Final check
        if ($script:boolScriptIsModulesLoaded -and $script:boolScriptIsEMS) { return $true }
        else { return $false }
    }#process

    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Debug -WriteBackToHost -Message "Finishing _ConfirmScriptRequirements"
    }#end
}#function _ConfirmScriptRequirements

Function _GetMailboxInfo {
    <#
    .SYNOPSIS
        Collects the necessary recipient cache and returns a datatable with the results
    .PARAMETER Filter
        The optional parameter for a filter to be used when querying recipient
   .PARAMETER MailboxAttributes
        Specify the array of mailbox attributes to return with the DataTable.
   .PARAMETER EmailAddresses,
        Specify the EmailAddresses datatable to update with found information
   .PARAMETER ExtensionCustomAttribute1,
        Specify the ExtensionCustomAttribute1 datatable to update with found information
   .PARAMETER ExtensionCustomAttribute2,
        Specify the ExtensionCustomAttribute2 datatable to update with found information
   .PARAMETER ExtensionCustomAttribute3,
        Specify the ExtensionCustomAttribute3 datatable to update with found information
   .PARAMETER ExtensionCustomAttribute4,
        Specify the ExtensionCustomAttribute4 datatable to update with found information
    .PARAMETER ExtensionCustomAttribute5
        Specify the ExtensionCustomAttribute5 datatable to update with found information
    .EXAMPLE
        $dtRecipients = _GetMailboxInfo  -MailboxAttributes $MBXAttribs
    
        This would get all mailboxes and return attributes $MBXAttribs to $dtRecipients
    .NOTES
        Version:
            - 5.1.2024.0516:    New function
    #>
 
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the filter to use for the recipients.")]
        [string]$Filter = "*",
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the array of mailbox attributes to return with the DataTable.")]
        [array]$MailboxAttributes,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify if group mailboxes should be returned with the DataTable.")]
        [switch]$IncludeGroups,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the EmailAddresses datatable to update with found information")]
        [System.Data.DataTable]$EmailAddresses,
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute1 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute1,
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute2 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute2,
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute3 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute3,
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute4 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute4,
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute5 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute5
    )
    
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetMailboxInfo"
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Building base datatable"

            $dtMailboxes = Get-Mailbox -ResultSize 1 -WarningAction SilentlyContinue -Verbose:$false | Select-Object -Property $MailboxAttributes | ConvertTo-DataTable
            $dtMailboxes.PrimaryKey = $dtMailboxes.Columns["Guid"]

            #Stupid blank attributes that default columns to strings
            [void]$dtMailboxes.Columns.Remove("LitigationHoldDate")
            [void]$dtMailboxes.Columns.Add("LitigationHoldDate", [datetime])
            $dtMailboxes.Clear()

            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with mailbox information"
            Get-Mailbox -ResultSize Unlimited -Filter $Filter -Verbose:$false -ErrorAction Stop | Select-Object -Property $MailboxAttributes | ForEach-Object {
        
                $drNewRow = $dtMailboxes.NewRow()
                ForEach ($element in $_.PSObject.Properties) {
                    $columnName = $element.Name
                    $columnValue = $element.Value
                    
                    if ([string]::IsNullorEmpty($columnValue) -or $columnValue.ToString() -eq "Unlimited") {
                        $columnValue = [DBNull]::Value
                    } else {
                        switch ($columnName) {
                            "ProhibitSendQuota" { $drNewRow["ProhibitSendQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "ProhibitSendReceiveQuota" { $drNewRow["ProhibitSendReceiveQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "RecoverableItemsQuota" { $drNewRow["RecoverableItemsQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "RecoverableItemsWarningQuota" { $drNewRow["RecoverableItemsWarningQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "CalendarLoggingQuota" { $drNewRow["CalendarLoggingQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "IssueWarningQuota" { $drNewRow["IssueWarningQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "RulesQuota" { $drNewRow["RulesQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1KB), 2) }
                            "MaxSendSize" { $drNewRow["MaxSendSize"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "MaxReceiveSize" { $drNewRow["MaxReceiveSize"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "ArchiveQuota" { $drNewRow["ArchiveQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "ArchiveWarningQuota" { $drNewRow["ArchiveWarningQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "EmailAddresses" {
                                ForEach ($entry in $columnValue) {
                                    $drNewAddressRow = $EmailAddresses.NewRow()
                                    $drNewAddressRow["ObjectDomain"] = [string]$ObjectDomain
                                    $drNewAddressRow["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                    $drNewAddressRow["EmailAddresses"] = [string]$entry
                                    [void]$EmailAddresses.Rows.Add($drNewAddressRow)
                                }#foreach
                            }#EmailAddresses
                            "ExtensionCustomAttribute1" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute1Row = $ExtensionCustomAttribute1.NewRow()
                                    $drExtCustomAttribute1Row["ObjectDomain"] = [string]$ObjectDomain
                                    $drExtCustomAttribute1Row["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                    $drExtCustomAttribute1Row["ExtensionCustomAttribute1"] = [string]$entry
                                    [void]$ExtensionCustomAttribute1.Rows.Add($drExtCustomAttribute1Row)
                                }#foreach
                            }#ExtensionCustomAttribute1
                            "ExtensionCustomAttribute2" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute2Row = $ExtensionCustomAttribute2.NewRow()
                                    $drExtCustomAttribute2Row["ObjectDomain"] = [string]$ObjectDomain
                                    $drExtCustomAttribute2Row["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                    $drExtCustomAttribute2Row["ExtensionCustomAttribute2"] = [string]$entry
                                    [void]$ExtensionCustomAttribute2.Rows.Add($drExtCustomAttribute2Row)
                                }#foreach
                            }#ExtensionCustomAttribute2
                            "ExtensionCustomAttribute3" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute3Row = $ExtensionCustomAttribute3.NewRow()
                                    $drExtCustomAttribute3Row["ObjectDomain"] = [string]$ObjectDomain
                                    $drExtCustomAttribute3Row["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                    $drExtCustomAttribute3Row["ExtensionCustomAttribute3"] = [string]$entry
                                    [void]$ExtensionCustomAttribute3.Rows.Add($drExtCustomAttribute3Row)
                                }#foreach
                            }#ExtensionCustomAttribute3
                            "ExtensionCustomAttribute4" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute4Row = $ExtensionCustomAttribute4.NewRow()
                                    $drExtCustomAttribute4Row["ObjectDomain"] = [string]$ObjectDomain
                                    $drExtCustomAttribute4Row["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                    $drExtCustomAttribute4Row["ExtensionCustomAttribute4"] = [string]$entry
                                    [void]$ExtensionCustomAttribute4.Rows.Add($drExtCustomAttribute4Row)
                                }#foreach
                            }#ExtensionCustomAttribute4
                            "ExtensionCustomAttribute5" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute5Row = $ExtensionCustomAttribute5.NewRow()
                                    $drExtCustomAttribute5Row["ObjectDomain"] = [string]$ObjectDomain
                                    $drExtCustomAttribute5Row["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                    $drExtCustomAttribute5Row["ExtensionCustomAttribute5"] = [string]$entry
                                    [void]$ExtensionCustomAttribute5.Rows.Add($drExtCustomAttribute5Row)
                                }#foreach
                            }#ExtensionCustomAttribute5
                            default {
                                if ($columnValue.gettype().Name -eq "ArrayList") {
                                    $drNewRow["$columnName"] = $columnValue.Clone()
                                } else {
                                    $drNewRow["$columnName"] = $columnValue
                                }#if/else
                            }#default
                        }#switch
                    }#if/else
                }#loop through each property

                [void]$dtMailboxes.Rows.Add($drNewRow)
            }#get-mailbox/foreach

            if ($IncludeGroups) {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with group mailbox information"
                
                Get-Mailbox -GroupMailbox -ResultSize Unlimited -Filter $Filter -Verbose:$false -ErrorAction Stop | Select-Object -Property $MailboxAttributes | ForEach-Object {
                    $drNewRow = $dtMailboxes.NewRow()
                    ForEach ($element in $_.PSObject.Properties) {
                        $columnName = $element.Name
                        $columnValue = $element.Value
                        
                        if ([string]::IsNullorEmpty($columnValue) -or $columnValue.ToString() -eq "Unlimited") {
                            $columnValue = [DBNull]::Value
                        } else {
                            switch ($columnName) {
                                "ProhibitSendQuota" { $drNewRow["ProhibitSendQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                                "ProhibitSendReceiveQuota" { $drNewRow["ProhibitSendReceiveQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                                "RecoverableItemsQuota" { $drNewRow["RecoverableItemsQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                                "RecoverableItemsWarningQuota" { $drNewRow["RecoverableItemsWarningQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                                "CalendarLoggingQuota" { $drNewRow["CalendarLoggingQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                                "IssueWarningQuota" { $drNewRow["IssueWarningQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                                "RulesQuota" { $drNewRow["RulesQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1KB), 2) }
                                "MaxSendSize" { $drNewRow["MaxSendSize"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                                "MaxReceiveSize" { $drNewRow["MaxReceiveSize"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                                "ArchiveQuota" { $drNewRow["ArchiveQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                                "ArchiveWarningQuota" { $drNewRow["ArchiveWarningQuota"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                                "EmailAddresses" {
                                    ForEach ($entry in $columnValue) {
                                        $drNewAddressRow = $EmailAddresses.NewRow()
                                        $drNewAddressRow["ObjectDomain"] = [string]$ObjectDomain
                                        $drNewAddressRow["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                        $drNewAddressRow["EmailAddresses"] = [string]$entry
                                        [void]$EmailAddresses.Rows.Add($drNewAddressRow)
                                    }#foreach
                                }#EmailAddresses
                                "ExtensionCustomAttribute1" {
                                    ForEach ($entry in $columnValue) {
                                        $drExtCustomAttribute1Row = $ExtensionCustomAttribute1.NewRow()
                                        $drExtCustomAttribute1Row["ObjectDomain"] = [string]$ObjectDomain
                                        $drExtCustomAttribute1Row["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                        $drExtCustomAttribute1Row["ExtensionCustomAttribute1"] = [string]$entry
                                        [void]$ExtensionCustomAttribute1.Rows.Add($drExtCustomAttribute1Row)
                                    }#foreach
                                }#ExtensionCustomAttribute1
                                "ExtensionCustomAttribute2" {
                                    ForEach ($entry in $columnValue) {
                                        $drExtCustomAttribute2Row = $ExtensionCustomAttribute2.NewRow()
                                        $drExtCustomAttribute2Row["ObjectDomain"] = [string]$ObjectDomain
                                        $drExtCustomAttribute2Row["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                        $drExtCustomAttribute2Row["ExtensionCustomAttribute2"] = [string]$entry
                                        [void]$ExtensionCustomAttribute2.Rows.Add($drExtCustomAttribute2Row)
                                    }#foreach
                                }#ExtensionCustomAttribute2
                                "ExtensionCustomAttribute3" {
                                    ForEach ($entry in $columnValue) {
                                        $drExtCustomAttribute3Row = $ExtensionCustomAttribute3.NewRow()
                                        $drExtCustomAttribute3Row["ObjectDomain"] = [string]$ObjectDomain
                                        $drExtCustomAttribute3Row["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                        $drExtCustomAttribute3Row["ExtensionCustomAttribute3"] = [string]$entry
                                        [void]$ExtensionCustomAttribute3.Rows.Add($drExtCustomAttribute3Row)
                                    }#foreach
                                }#ExtensionCustomAttribute3
                                "ExtensionCustomAttribute4" {
                                    ForEach ($entry in $columnValue) {
                                        $drExtCustomAttribute4Row = $ExtensionCustomAttribute4.NewRow()
                                        $drExtCustomAttribute4Row["ObjectDomain"] = [string]$ObjectDomain
                                        $drExtCustomAttribute4Row["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                        $drExtCustomAttribute4Row["ExtensionCustomAttribute4"] = [string]$entry
                                        [void]$ExtensionCustomAttribute4.Rows.Add($drExtCustomAttribute4Row)
                                    }#foreach
                                }#ExtensionCustomAttribute4
                                "ExtensionCustomAttribute5" {
                                    ForEach ($entry in $columnValue) {
                                        $drExtCustomAttribute5Row = $ExtensionCustomAttribute5.NewRow()
                                        $drExtCustomAttribute5Row["ObjectDomain"] = [string]$ObjectDomain
                                        $drExtCustomAttribute5Row["ExchangeGuid"] = [guid]($drNewRow["ExchangeGuid"]).Guid
                                        $drExtCustomAttribute5Row["ExtensionCustomAttribute5"] = [string]$entry
                                        [void]$ExtensionCustomAttribute5.Rows.Add($drExtCustomAttribute5Row)
                                    }#foreach
                                }#ExtensionCustomAttribute5
                                default {
                                    if ($columnValue.gettype().Name -eq "ArrayList") {
                                        $drNewRow["$columnName"] = $columnValue.Clone()
                                    } else {
                                        $drNewRow["$columnName"] = $columnValue
                                    }#if/else
                                }#default
                            }#switch
                        }#if/else
                    }#loop through each property

                    [void]$dtMailboxes.Rows.Add($drNewRow)
                }#get-mailbox/foreach
            }#if

            return @(, $dtMailboxes)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather mailbox information"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
    
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetMailboxInfo"
    }#end
}#function _GetMailboxInfo

Function _GetCASMailboxInfo {
    <#
    .SYNOPSIS
        Collects the CAS mailbox info cache and returns a datatable with the results
    .PARAMETER Session
        The PSSession to run the command against
    .PARAMETER Mailboxes
        Specify the Mailboxes datatable to update with found informationa
    .PARAMETER CASAttributes
        Specify the array of CAS attributes to return with the DataTable
    .EXAMPLE
        _GetCASMailboxInfo -Mailboxes $dtRecipients -CASAttributes $CASAttribs | Out-Null
    
        This would get all CAS mailboxes and return attributes $CASAttribs to $dtRecipients
    .NOTES
        Version:
            - 5.1.2024.0516:    New function
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the Mailboxes datatable to update with found information")]
        [System.Data.DataTable]$Mailboxes,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the array of CAS attributes to return with the DataTable.")]
        [array]$CASAttributes
    )
        
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetCASMailboxInfo"
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Appending CAS attributes to datatable"
            
            $dtCAS = Get-CASMailbox -ResultSize 1 -WarningAction SilentlyContinue -Verbose:$false | Select-Object -Property $CASAttributes | ConvertTo-DataTable
            
            foreach ($column in $dtCAS.Columns) { if ($column.ColumnName -ne "Guid") { $Mailboxes.Columns.Add($column.ColumnName, $column.DataType) } }

            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with related CAS mailbox information"
            
            Get-CASMailbox -ResultSize Unlimited -Verbose:$false | Select-Object -Property $CASAttributes | ForEach-Object {
                $row = $Mailboxes.Rows.Find($_.Guid)

                if ($row) {
                    foreach ($element in $_.PSObject.Properties) {
                        $columnName = $element.Name
                        $columnValue = $element.Value

                        if ([string]::IsNullorEmpty($columnValue)) { $columnValue = [DBNull]::Value }
                        
                        if ($columnValue.gettype().Name -eq "ArrayList") {
                            $row["$columnName"] = $columnValue.Clone()
                        } else {
                            $row["$columnName"] = $columnValue
                        }
                    }#loop through each property
                }#if
            }#invoke/foreach
            
            return @(, $Mailboxes)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather CAS mailbox information"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
          
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetCASMailboxInfo"
    }#end
}#function _GetCASMailboxInfo

Function _GetMobileInfo {
    <#
    .SYNOPSIS
        Collects the mobile device info and returns a datatable with the results
    .PARAMETER Mailboxes
        Specify the Mailboxes datatable to use as a query
    .PARAMETER MobileAttributes
        Specify the array of mobile attributes to return with the DataTable
    .EXAMPLE
        $dtMobile = _GetMobileInfo -Mailboxes $dtRecipients -MobileAttributes $MobileAttribs
    
        This would get all mobile devices and return attributes $MobileAttribs to $dtMobile
    .NOTES
        Version:
            - 5.1.2024.0517:    New function
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the Mailboxes datatable to use as a query")]
        [System.Data.DataTable]$Mailboxes,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the array of mobile attributes to return with the DataTable.")]
        [array]$MobileAttributes
    )
        
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetMobileInfo"
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Building base mobile device datatable"
            $dtMobile = Get-MobileDevice -ResultSize 1 -WarningAction SilentlyContinue -Verbose:$false | Select-Object -Property $MobileAttributes | ConvertTo-DataTable

            $dtMobile.PrimaryKey = $dtMobile.Columns["Guid"]
            $dtMobile.Clear()
            
            foreach ($mailbox in $Mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' }) {
                $MBXGuid = $mailbox.Guid.ToString()
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with related mobile device information for $MBXGuid"

                Get-MobileDevice -Mailbox $MBXGuid -Verbose:$false | Select-Object -Property $MobileAttributes | ForEach-Object {
                    $drNewRow = $dtMobile.NewRow()

                    ForEach ($element in $_.PSObject.Properties) {
                        $columnName = $element.Name
                        $columnValue = $element.Value
                        
                        if ([string]::IsNullorEmpty($columnValue)) { $columnValue = [DBNull]::Value }

                        if ($columnValue.gettype().Name -eq "ArrayList") {
                            $drNewRow["$columnName"] = $columnValue.Clone()
                        } else {
                            switch ($columnName) {
                                "ExchangeGuid" { $drNewRow["ExchangeGuid"] = $MBXGuid }
                                "ObjectDomain" { $drNewRow["ObjectDomain"] = $ObjectDomain }
                                default {
                                    $drNewRow["$columnName"] = $columnValue
                                }#default
                            }#switch
                        }#if/else
                    }#loop through each property

                    [void]$dtMobile.Rows.Add($drNewRow)
                }#invoke/foreach
            }#foreach

            return @(, $dtMobile)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather mobile device information"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
          
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetMobileInfo"
    }#end
}#function _GetMobileInfo

Function _GetMobileStats {
    <#
    .SYNOPSIS
        Collects the mobile device statistics and returns a datatable with the results
    .PARAMETER MobileDevices
        Specify the Mailboxes datatable to update with found information
    .PARAMETER MobileStatsArribs
        Specify the array of mobile statistics attributes to return with the DataTable
    .EXAMPLE
        _GetMobileStats -MobileDevices $dtMobileDevices -MobileStatsArribs $MobileStatAttribs | Out-Null
    
        This would get all mobile device stats and return attributes $MobileStatAttribs to $dtMobileDevices
    .NOTES
        Version:
            - 5.1.2024.0517:    New function
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the mobile devices datatable to update with found information")]
        [System.Data.DataTable]$MobileDevices,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the array of mobile device statistics attributes to return with the DataTable.")]
        [array]$MobileStatsArribs
    )
        
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetMobileStats"
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Appending mobile device statistics attributes to datatable"
            foreach ($device in $MobileDevices) {
                #$deviceID = $device.Identity
                $deviceID = $device.Guid
                
                $dtStats = Get-MobileDeviceStatistics -Identity $deviceID -ErrorAction Stop -Verbose:$false | Select-Object -Property $MobileStatsArribs | ConvertTo-DataTable
                
                if ($dtStats) {
                    foreach ($column in $dtStats.Columns) {
                        if ($column.ColumnName -ne "Guid") {
                            $MobileDevices.Columns.Add($column.ColumnName, $column.DataType)
                        }#if
                    }#foreach
                    break
                }#if
            }#foreach
            
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with related mobile device statistics information"
            foreach ($device in $MobileDevices) {
                #$deviceID = $device.Identity
                $deviceID = $device.Guid
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Debug -WriteBackToHost -Message "Processing mobile device statistics for $deviceID"

                Get-MobileDeviceStatistics -Identity $deviceID -ErrorAction Stop -Verbose:$false | Select-Object -Property $MobileStatsArribs | ForEach-Object {
                    $row = $MobileDevices.Rows.Find($_.Guid)

                    if ($row) {
                        foreach ($element in $_.PSObject.Properties) {
                            $columnName = $element.Name
                            $columnValue = $element.Value

                            if ([string]::IsNullorEmpty($columnValue)) { $columnValue = [DBNull]::Value }
                            
                            if ($columnValue.gettype().Name -eq "ArrayList") {
                                $row["$columnName"] = $columnValue.Clone()
                            } else {
                                $row["$columnName"] = $columnValue
                            }
                        }#loop through each property
                    }#if
                }#invoke/foreach
                
            }#foreach

            return @(, $MobileDevices)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather mobile device statistics"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
          
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetMobileStats"
    }#end
}#function _GetMobileStats

Function _GetInboxRules {
    <#
    .SYNOPSIS
        Collects the inbox rule count for a mailbox
    .PARAMETER Mailboxes
        Specify the Mailboxes datatable to use as a query
    .EXAMPLE
        $dtInbox = _GetInboxRules -Mailboxes $dtRecipients
    
        This would get all inbox rule counts and return as another column in Mailboxes
    .NOTES
        Version:
            - 5.1.2024.0517:    New function
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the Mailboxes datatable to use as a query")]
        [System.Data.DataTable]$Mailboxes
    )
        
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetInboxRules"
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Appending Inbox rule count attribute to datatable"
            $Mailboxes.Columns.Add("InboxRuleCount", [int])

            foreach ($mailbox in $Mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' }) {
                $MBXGuid = $mailbox.Guid.ToString()
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with related inbox rule information for $MBXGuid"
                $RuleCount = (Get-InboxRule -Mailbox $MBXGuid -Verbose:$false | Select-Object -Property Name).Count
                
                if ([string]::IsNullorEmpty($RuleCount)) { $mailbox.InboxRuleCount = [DBNull]::Value }
                else { $mailbox.InboxRuleCount = $RuleCount }
            }#foreach

            return $true
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather inbox rules"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
          
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetInboxRules"
    }#end
}#function _GetInboxRules

Function _GetRegionalInfo {
    <#
    .SYNOPSIS
        Collects the regional mailbox info for a mailbox
    .PARAMETER Mailboxes
        Specify the Mailboxes datatable to use as a query
    .EXAMPLE
        $dtInbox = _GetRegionalInfo -Session $session -Mailboxes $dtRecipients
    
        This would get all regional mailbox information and return as another column in dtRecipients
    .NOTES
        Version:
            - 5.1.2024.0517:    New function
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the Mailboxes datatable to use as a query")]
        [System.Data.DataTable]$Mailboxes
    )
        
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetRegionalInfo"
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Appending regional info attributes to datatable"
            $Mailboxes.Columns.Add("Language", [string])
            $Mailboxes.Columns.Add("TimeZone", [string])

            foreach ($mailbox in $Mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' }) {
                $MBXGuid = $mailbox.Guid.ToString()
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with related regional information for $MBXGuid"
                
                Get-MailboxRegionalConfiguration -Identity $MBXGuid -Verbose:$false | Select-Object -Property Language, TimeZone, DateFormat | ForEach-Object {
                    foreach ($element in $_.PSObject.Properties) {
                        $columnName = $element.Name
                        $columnValue = $element.Value
                        if ([string]::IsNullorEmpty($columnValue)) { $columnValue = [DBNull]::Value }
                        
                        if ($columnValue.gettype().Name -eq "CultureInfo") {
                            $mailbox["$columnName"] = $columnValue.Name
                        } else {
                            $mailbox["$columnName"] = $columnValue
                        }
                    }#loop through each property
                }#invoke/foreach
            }#foreach
            return $true
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather regional info"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
          
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetRegionalInfo"
    }#end
}#function _GetRegionalInfo

Function _ConfirmExchangeSession {
    <#
    .SYNOPSIS
        Check on age of Exchange session and re-establishes if it is has expired
    .EXAMPLE

    .NOTES
        Version:
        - 5.1.2024.0516:    New function
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify FQDN of an available on-premises Exchange server. It will default to 'Online' if not specified.")]
        [string]$Server = "Online",

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the file path for the Exchange Server credentials.")]
        [string]$CredentialFile,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify only if you need to force SSL on for the remote PowerShell connection. It will default to false if not specified.")]
        [switch]$ForceSSL,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the remote PowerShell the specific authentication method. It will default to Basic if not specified.")]
        [ValidateSet("Basic", "Negotiate", "Kerberos")]
        [string]$AuthMethod = "Basic",

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the credentials for connecting to Exchange.")]
        [PSCredential]$Creds
    )

    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _ConfirmExchangeSession"
        $NeedToConnect = $true
    }#begin
    
    process {
        try {
            $PSSessionID = Get-PSSession -ErrorAction SilentlyContinue | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }
            
            if ($PSSessionID) {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Exchange PS session found"

                if (-not [string]::IsNullOrEmpty($PSSessionID.State)) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Exchange PS session state is $($PSSessionID.State)"
                    $NeedToConnect = $false
                } elseif ($PSSessionID.State -ne "Opened") {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Warning -WriteBackToHost -Message "Exchange PS session is not valid. Establishing new PS session"
                    
                    if ($PSSessionID) {
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Disconnecting existing PS session"
                        $PSSessionID | Remove-PSSession -ErrorAction SilentlyContinue -Verbose:$false
                    }#if session already connected

                    $NeedToConnect = $true
                }#if/elseif
            } else {
                if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue -Verbose:$false) { $RESTSession = Get-ConnectionInformation -ErrorAction SilentlyContinue -Verbose:$false }

                if ($RESTSession) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Exchange REST connection found"

                    if (-not [string]::IsNullOrEmpty($RESTSession.State)) {
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Exchange REST connection state is $($RESTSession.State)"
                        $NeedToConnect = $false
                    } elseif ($RESTSession.State -ne "Connected") {
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Warning -WriteBackToHost -Message "Exchange REST connection is not valid. Establishing new REST connection"
                        
                        if ($RESTSession) {
                            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Disconnecting existing REST connection"
                            Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
                        }#if session already connected

                        $NeedToConnect = $true
                    }#if
                }#if REST session already connected
            }#if/else

            if ($NeedToConnect) {
                $global:VerbosePreference = "SilentlyContinue"

                $ConnectionSplat = @{
                    Server  = $Server
                    Verbose = $false
                }

                if ($Server -ne "Online") {
                    $ConnectionSplat.Add("UseSSL", $ForceSSL)
                    $ConnectionSplat.Add("AuthenticationMethod", $AuthMethod)
                }

                if ($Creds) {
                    $ConnectionSplat.Add("Credential", $Creds)
                } elseif ($CredentialFile) {
                    $ConnectionSplat.Add("CredentialFile", $CredentialFile)
                }#if/else
                
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message $($ConnectionSplat | Out-String)
                Connect-Exchange @ConnectionSplat -ErrorAction Stop
            }#if

            if ($htLoggingPreference['VerbosePreference'] -eq "Continue") { $global:VerbosePreference = "Continue" }#if

            $Session = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }
            if (-not $Session) {
                if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue -Verbose:$false) { $Session = Get-ConnectionInformation -ErrorAction SilentlyContinue -Verbose:$false }
            }#if/else

            if ($Session) {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Connection established"
            } else { throw "Unable to establish connect" }#if/else

            $mod = (Get-Module | Where-Object { $_.ExportedFunctions.ContainsKey("Get-Mailbox") -eq $true }).Name
            if ($mod -and $Session) {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Exchange module found"
                $script:boolScriptIsEMS = $true

                if (Get-Command -Module $mod -Name "Get-ReportSubmissionPolicy" -ErrorAction SilentlyContinue) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Exchange module contained Get-ReportSubmissionPolicy, must be Exchange Online"
                    $script:boolScriptIsEMSOnline = $true
                }#if online
            }#if module and pssession
            
            if ($script:boolScriptIsEMS -and -not $script:boolScriptIsEMSOnline) {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Exchange On-Premises found so setting ViewEntireForest to true"
                if (-not (Get-AdServerSettings).ViewEntireForest) {
                    Set-AdServerSettings -ViewEntireForest $true | Out-Null
                }#if
            }#if on-premises

            return $Session
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error resetting Exchange session"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            Exit $ExitCode
        }#try/catch
    }#process

    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _ConfirmExchangeSession"
    }#end
}#function _ConfirmExchangeSession

Function _CreateMVATable {
    <#
    .SYNOPSIS
        Create a blank MVA datatable
    .PARAMETER Attribute
        Specify the MVA attribute build the MVA DataTable.
    .EXAMPLE
        $dtGroups = _CreateMVATable Attribute "ManagedBy"
    
        This would create a blank MVA datatable for "ManagedBy" attribute
    .NOTES
        Version:
            - 5.0.2022.0329:    Initial function
    #>
 
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the attribute for the MVA datatable")]
        [string]$Attribute
    )
    
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _CreateMVATable"
    }#begin
    
    process	{
        try	{
            $dtMVA = New-Object System.Data.DataTable
            $col1 = New-Object System.Data.DataColumn ObjectDomain, ([string])
            $col2 = New-Object System.Data.DataColumn ExchangeGuid, ([Guid])
            $col3 = New-Object System.Data.DataColumn $Attribute, ([string])
            $dtMVA.Columns.Add($col1)
            $dtMVA.Columns.Add($col2)
            $dtMVA.Columns.Add($col3)
            [System.Data.DataColumn[]]$KeyColumn = ($dtMVA.Columns["ObjectDomain"], $dtMVA.Columns["ExchangeGuid"], $dtMVA.Columns[$Attribute])
            $dtMVA.PrimaryKey = $KeyColumn
            return @(, $dtMVA)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to create MVA datatable"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $false
        }#try/catch
    }#process
    
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _CreateMVATable"
    }#end
}#function _CreateMVATable

function _GetScriptDirectory {
    <#
    .SYNOPSIS
        _GetScriptDirectory returns the proper location of the script.
 
    .OUTPUTS
        System.String
   
    .NOTES
        Returns the correct path within a packaged executable.
    #>
    [OutputType([string])]
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [switch]$Path,

        [Parameter(Mandatory = $false)]
        [switch]$Leaf,

        [Parameter(Mandatory = $false)]
        [switch]$LeafBase
    )

    if ($null -ne $hostinvocation) {
        if ($Leaf) {
            Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf
        } elseif ($LeafBase) {
            (Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf).Split(".")[0]
        } elseif ($Path) {
            Split-Path $hostinvocation.MyCommand.path
        } else {
            Split-Path $hostinvocation.MyCommand.path
        }#if/else
    } elseif ($null -ne $script:MyInvocation.MyCommand.Path) {
        if ($Leaf) {
            Split-Path $script:MyInvocation.MyCommand.Path -Leaf
        } elseif ($LeafBase) {
            (Split-Path $script:MyInvocation.MyCommand.Path -Leaf).Split(".")[0]
        } elseif ($Path) {
            Split-Path $script:MyInvocation.MyCommand.Path
        } else {
            (Get-Location).Path + "\" + (Split-Path $script:MyInvocation.MyCommand.Path -Leaf)
        }#if/else
    } else {
        if ($Leaf) {
            Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf
        } elseif ($LeafBase) {
            (Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf).Split(".")[0]
        } elseif ($Path) {
            (Get-Location).Path
        } else {
            (Get-Location).Path + "\" + (Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf)
        }#if/else
    }#if/else
}#function _GetScriptDirectory
#endregion

#region Active Development
Function _SubmitMySQLChanges {
    <#
    .SYNOPSIS
        Process data that needs to be saved to SQL and returns true/false for success
    .PARAMETER ConnectionString
        Specify the MySQL Server connection string
    .PARAMETER TargetTable
        Specify the name of the target MySQL table
    .PARAMETER PropertyMap
        Specify the hashtable for the MySQL column to property map
    .PARAMETER DataToMerge
        Specify the datatable that should be merged into MySQL
    .PARAMETER BatchSize
        Specify the staging threshold that should be used for pushing changes to MySQL
    .PARAMETER DataDomain
        Specify the data domain/primary key that will be associated with the data
    .EXAMPLE
        $results = _SubmitMySQLChanges -ConnectionString $ConnectionString -TargetTable $targetTable -PropertyMap $PropMap -BatchSize $BatchSize -DataToMerge $dtMailboxes -DataDomain "contoso.com"
    
        This would process all data in $dtMailboxes against property map $propMap, submit those changes to $ConnectionString is batches of $batchsize
        to $targetTable
    .NOTES
        Version:
            - 5.1.2024.0517:    New function
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [string]$ConnectionString,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [string]$TargetTable,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [System.Data.DataTable]$DataToMerge,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [Hashtable]$PropertyMap,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [int]$BatchSize,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [string]$DataDomain
    )
        
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _SubmitMySQLChanges"
    }#begin
    
    process	{
        try	{
            $dtMerge = New-MySQLDatatable -ConnectionString $ConnectionString -TableName $TargetTable -IncludeConstraints
            
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Local data table based on MySQL created"

            #Merge changes
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Processing data"
            foreach ($row in $DataToMerge.Rows) {
                $targetrow = $dtMerge.NewRow()
                
                foreach ($SourceField in $PropertyMap.Keys) {
                    switch ($SourceField) {
                        "ExchangeGuid" { $targetrow[$PropertyMap[$SourceField]] = [guid]($row["ExchangeGuid"]) }
                        "Guid" { $targetrow[$PropertyMap[$SourceField]] = [guid]($row["Guid"]) }
                        "MobileDomain" {
                            $targetrow[$PropertyMap[$SourceField]] = $DataDomain
                        }
                        "WhenModifiedUTC" { $targetrow[$PropertyMap[$SourceField]] = (Get-Date).ToUniversalTime() }
                        "ObjectDomain" {
                            $targetrow[$PropertyMap[$SourceField]] = $DataDomain
                        }
                        default {
                            if ($null -eq $row[$SourceField] -or "" -eq "$($row[$SourceField])") {
                                $targetrow[$PropertyMap[$SourceField]] = [DBNull]::Value
                            } else {
                                $targetrow[$PropertyMap[$SourceField]] = $row.$SourceField
                            }#if/else
                        }#default
                    }#switch
                }#foreach

                [void]$dtMerge.Rows.Add($targetrow)
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Debug -WriteBackToHost -Message "Add row to merge data table"
                
                # Incremental bulk copy of ETL data into dynamically created temporary staging table to significantly reduce local memory consumption
                if ($dtMerge.GetChanges().Rows.Count -ge $BatchSize) {
                    #Manually clear local data table
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting export of batch data to MySQL"
                        
                    $dtChanges = $dtMerge.GetChanges()
                    $Updated = Invoke-MySQLMerge -ConnectionString $ConnectionString -TargetTable $TargetTable -ColumnList $PropertyMap.Values -Data $dtChanges -LastSeenDatetimeColumn 'LastSeenUTC'

                    if ($Updated) {
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished export of batch data to MySQL"
                        $dtMerge.Clear()
                        $dtMerge.AcceptChanges()
                            
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Cleared merge data table"
                        $results = $true
                    } else {
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Unable to batch changes to MySQL. Will retry on next loop."
                        $results = $false
                    }#if/else
                }#if
            }#foreach

            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting export of batch data and writing data to MySQL"

            if ($dtMerge.Rows.Count -gt 0) {
                $dtChanges = $dtMerge.GetChanges()
                $Updated = Invoke-MySQLMerge -ConnectionString $ConnectionString -TargetTable $TargetTable -ColumnList $PropertyMap.Values -Data $dtChanges -LastSeenDatetimeColumn 'LastSeenUTC'

                if ($Updated -ge 1) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished export of final data to MySQL"
                    $dtMerge.Clear()
                    $dtMerge.AcceptChanges()
                        
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Cleared merge data table"
                    $results = $true
                } else {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Unable to export final changes to MySQL"
                    $results = $false
                }#if/else

            }#if
            
            return $results
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to write data to MySQL"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $false
        }#try/catch
    }#process
          
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _SubmitMySQLChanges"
    }#end
}#function _SubmitMySQLChanges
#endregion

#region Main Program

Write-Host "`r"
Write-Host "Script Written by Sterling Consulting`r"
Write-Host "All rights reserved. Proprietary and Confidential Material`r"
Write-Host "Exchange Mailbox Information Collector Script`r"
Write-Host "`r"

Write-Host "Script starting`r"

if (_ConfirmScriptRequirements) {
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Script requirements met"

    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Creating MVA datatables"
    $dtEmailAddresses = _CreateMVATable -Attribute "EmailAddresses"
    $dtExtensionCustomAttribute1 = _CreateMVATable -Attribute "ExtensionCustomAttribute1"
    $dtExtensionCustomAttribute2 = _CreateMVATable -Attribute "ExtensionCustomAttribute2"
    $dtExtensionCustomAttribute3 = _CreateMVATable -Attribute "ExtensionCustomAttribute3"
    $dtExtensionCustomAttribute4 = _CreateMVATable -Attribute "ExtensionCustomAttribute4"
    $dtExtensionCustomAttribute5 = _CreateMVATable -Attribute "ExtensionCustomAttribute5"
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Finished creating MVA datatables"

    #Get mailboxes from system
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Retrieving mailbox information"
    $dtMailboxes = _GetMailboxInfo -Filter $MailboxFilter -MailboxAttributes $arrMailboxAttribs -IncludeGroups:$IncludeGroupMailboxes `
        -EmailAddresses $dtEmailAddresses -ExtensionCustomAttribute1 $dtExtensionCustomAttribute1 -ExtensionCustomAttribute2 $dtExtensionCustomAttribute2 `
        -ExtensionCustomAttribute3 $dtExtensionCustomAttribute3 -ExtensionCustomAttribute4 $dtExtensionCustomAttribute4 -ExtensionCustomAttribute5 $dtExtensionCustomAttribute5
    
    if ($dtMailboxes.DefaultView.Count -le 0) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "No mailbox information found. Unable to continue without mailbox information. Exiting script"
        Exit $ExitCode
    }#if
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished retrieving mailbox information"
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "$($dtMailboxes.Rows.Count) Mailbox entries collected"

    #Get CAS mailbox info from system
    if ($IncludeCASInfo) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Retrieving CAS mailbox information"
        _GetCASMailboxInfo -Mailboxes $dtMailboxes -CASAttributes $arrMailboxCASAttribs | Out-Null
        $MailboxPropertyMap += $CASPropertyMap
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished retrieving CAS mailbox information"
    }#if

    #Get Inbox rule count from system
    if ($IncludeInboxRules) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Retrieving Inbox rule information"
        _GetInboxRules -Mailboxes $dtMailboxes | Out-Null
        $MailboxPropertyMap += $InboxRulePropertyMap
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished Inbox rules information"
    }#if

    #Get Regional mailbox configuration from system
    if ($IncludeRegionalInfo) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Retrieving regional mailbox information"
        _GetRegionalInfo -Mailboxes $dtMailboxes | Out-Null
        $MailboxPropertyMap += $RegionalPropertyMap
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished retrieving regional mailbox information"
    }#if

    #Get mobile device info from system
    if ($IncludeMobileInfo -or $IncludeMobileStats) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Retrieving mobile device information"
        $dtMobileDevices = _GetMobileInfo -Mailboxes $dtMailboxes -MobileAttributes $arrMobileAttribs
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished retrieving mobile device information"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "$($dtMobileDevices.Rows.Count) Mobile devices collected"
    }#if

    #Get mobile device statistics info from system
    if ($IncludeMobileStats -and $dtMobileDevices.DefaultView.Count -gt 0) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Retrieving mobile device statistics information"
        _GetMobileStats -MobileDevices $dtMobileDevices -MobileStatsArribs $arrMobileStatsArribs | Out-Null
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished retrieving mobile device statistics information"
    } elseif ($IncludeMobileStats -and $dtMobileDevices.DefaultView.Count -le 0) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Warning -WriteBackToHost -Message "No mobile device information found. Unable to collect mobile device statistics"
    }#if/elseif


    if ($dtMailboxes.DefaultView.Count -gt 0) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "$($dtMailboxes.DefaultView.Count) mailboxes found"
        #Export data to requested destination
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Exporting data to $($ExportFormat)"
        switch ($ExportFormat) {
            "Clixml" {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting CliXML to $script:ExportLocation"
                
                #Check for path/folder
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Checking for $script:ExportLocation"
                if (-not (Test-Path -Path $script:ExportLocation)) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Creating folder structure for $script:ExportLocation"
                    New-Item -Path $script:ExportLocation -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
                }#if
                
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting export of CliXML"
                Export-Clixml -Path "$script:ExportLocation\$($ObjectDomain)_Mailboxes_$strLogTimeStamp.xml" -InputObject $dtMailboxes -Depth 3 | Out-Null
                Export-Clixml -Path "$script:ExportLocation\$($ObjectDomain)_EmailAddresses_$strLogTimeStamp.xml" -InputObject $dtEmailAddresses -Depth 3 | Out-Null

                if ($dtExtensionCustomAttribute1.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 1 data"
                    Export-Clixml -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute1_$strLogTimeStamp.xml" -InputObject $dtExtensionCustomAttribute1 -Depth 3 | Out-Null
                }

                if ($dtExtensionCustomAttribute2.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 2 data"
                    Export-Clixml -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute2_$strLogTimeStamp.xml" -InputObject $dtExtensionCustomAttribute2 -Depth 3 | Out-Null
                }

                if ($dtExtensionCustomAttribute3.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 3 data"
                    Export-Clixml -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute3_$strLogTimeStamp.xml" -InputObject $dtExtensionCustomAttribute3 -Depth 3 | Out-Null
                }

                if ($dtExtensionCustomAttribute4.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 4 data"
                    Export-Clixml -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute4_$strLogTimeStamp.xml" -InputObject $dtExtensionCustomAttribute4 -Depth 3 | Out-Null
                }

                if ($dtExtensionCustomAttribute5.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 5 data"
                    Export-Clixml -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute5_$strLogTimeStamp.xml" -InputObject $dtExtensionCustomAttribute5 -Depth 3 | Out-Null
                }

                if ($dtMobileDevices.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting mobile devices data"
                    Export-Clixml -Path "$script:ExportLocation\$($ObjectDomain)_MobileDevices_$strLogTimeStamp.xml" -InputObject $dtMobileDevices -Depth 3 | Out-Null
                }
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished export of CliXML"
            }#clixml
            "CSV" {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting CSV to $script:ExportLocation with $Delimiter delimiter"
                
                #Check for path/folder
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Checking for $script:ExportLocation"
                if (-not (Test-Path -Path $script:ExportLocation)) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Creating folder structure for $script:ExportLocation"
                    New-Item -Path $script:ExportLocation -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
                }#if
                
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting export of CSV"
                $dtMailboxes | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_Mailboxes_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                $dtEmailAddresses | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_EmailAddresses_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null

                if ($dtExtensionCustomAttribute1.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 1 data"
                    $dtExtensionCustomAttribute1 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute1_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }

                if ($dtExtensionCustomAttribute2.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 2 data"
                    $dtExtensionCustomAttribute2 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute2_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }

                if ($dtExtensionCustomAttribute3.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 3 data"
                    $dtExtensionCustomAttribute3 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute3_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }

                if ($dtExtensionCustomAttribute4.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 4 data"
                    $dtExtensionCustomAttribute4 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute4_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }

                if ($dtExtensionCustomAttribute5.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 5 data"
                    $dtExtensionCustomAttribute5 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute5_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }

                if ($dtMobileDevices.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting mobile devices data"
                    $dtMobileDevices | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_MobileDevices_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished export of CSV"
            }#csv
            "TSV" {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting TSV to $script:ExportLocation with tab delimiter"
                
                #Check for path/folder
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Checking for $script:ExportLocation"
                if (-not (Test-Path -Path $script:ExportLocation)) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Creating folder structure for $script:ExportLocation"
                    New-Item -Path $script:ExportLocation -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
                }#if
                
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting export of TSV"
                $dtMailboxes | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_Mailboxes_$strLogTimeStamp.csv" -Delimiter "`t" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                $dtEmailAddresses | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_EmailAddresses_$strLogTimeStamp.csv" -Delimiter "`t" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null

                if ($dtExtensionCustomAttribute1.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 1 data"
                    $dtExtensionCustomAttribute1 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute1_$strLogTimeStamp.csv" -Delimiter "`t" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }

                if ($dtExtensionCustomAttribute2.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 2 data"
                    $dtExtensionCustomAttribute2 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute2_$strLogTimeStamp.csv" -Delimiter "`t" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }

                if ($dtExtensionCustomAttribute3.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 3 data"
                    $dtExtensionCustomAttribute3 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute3_$strLogTimeStamp.csv" -Delimiter "`t" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }

                if ($dtExtensionCustomAttribute4.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 4 data"
                    $dtExtensionCustomAttribute4 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute4_$strLogTimeStamp.csv" -Delimiter"`t" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }

                if ($dtExtensionCustomAttribute5.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 5 data"
                    $dtExtensionCustomAttribute5 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute5_$strLogTimeStamp.csv" -Delimiter "`t" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }

                if ($dtMobileDevices.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting mobile devices data"
                    $dtMobileDevices | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_MobileDevices_$strLogTimeStamp.csv" -Delimiter "`t" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
                }
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished export of TSV"
            }#tsv
            "JSON" {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting JSON to $script:ExportLocation"
                
                #Check for path/folder
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Checking for $script:ExportLocation"
                if (-not (Test-Path -Path $script:ExportLocation)) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Creating folder structure for $script:ExportLocation"
                    New-Item -Path $script:ExportLocation -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
                }#if
                
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting export of JSON"
                Export-Json -Path "$script:ExportLocation\$($ObjectDomain)_Mailboxes_$strLogTimeStamp.json" -InputObject $dtMailboxes -Depth 3 | Out-Null
                Export-Json -Path "$script:ExportLocation\$($ObjectDomain)_EmailAddresses_$strLogTimeStamp.json" -InputObject $dtEmailAddresses -Depth 3 | Out-Null

                if ($dtExtensionCustomAttribute1.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 1 data"
                    Export-Json -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute1_$strLogTimeStamp.json" -InputObject $dtExtensionCustomAttribute1 -Depth 3 | Out-Null
                }

                if ($dtExtensionCustomAttribute2.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 2 data"
                    Export-Json -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute2_$strLogTimeStamp.json" -InputObject $dtExtensionCustomAttribute2 -Depth 3 | Out-Null
                }

                if ($dtExtensionCustomAttribute3.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 3 data"
                    Export-Json -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute3_$strLogTimeStamp.json" -InputObject $dtExtensionCustomAttribute3 -Depth 3 | Out-Null
                }

                if ($dtExtensionCustomAttribute4.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 4 data"
                    Export-Json -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute4_$strLogTimeStamp.json" -InputObject $dtExtensionCustomAttribute4 -Depth 3 | Out-Null
                }

                if ($dtExtensionCustomAttribute5.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting extension attribute 5 data"
                    Export-Json -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute5_$strLogTimeStamp.json" -InputObject $dtExtensionCustomAttribute5 -Depth 3 | Out-Null
                }

                if ($dtMobileDevices.DefaultView.Count -gt 0) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting mobile devices data"
                    Export-Json -Path "$script:ExportLocation\$($ObjectDomain)_MobileDevices_$strLogTimeStamp.json" -InputObject $dtMobileDevices -Depth 3 | Out-Null
                }
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished export of JSON"
            }#json
            "DB" {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Processing mailbox information data"

                try {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Processing data for MySQL export"
                    $MySQLConnectionString = New-MySQLConnectionString -Server $MySQLServer -Port $MySQLPort -Database $Database -Credential $MySQLCredential

                    $SuccessToDB = _SubmitMySQLChanges -ConnectionString $MySQLConnectionString -TargetTable $MySQLMailboxTable -PropertyMap $MailboxPropertyMap `
                        -BatchSize $MySQLStagingSize -DataToMerge $dtMailboxes -DataDomain $ObjectDomain
                    $SuccessToDB = _SubmitMySQLChanges -ConnectionString $MySQLConnectionString -TargetTable $MySQLEmailAddressesTable -PropertyMap $EmailAddressesPropertyMap `
                        -BatchSize $MySQLStagingSize -DataToMerge $dtEmailAddresses -DataDomain $ObjectDomain

                    if ($dtExtensionCustomAttribute1.DefaultView.Count -gt 0) {
                        $SuccessToDB = _SubmitMySQLChanges -ConnectionString $MySQLConnectionString -TargetTable $MySQLExtensionCustomAttribute1Table -PropertyMap $ExtensionCustomAttribute1PropertyMap `
                            -BatchSize $MySQLStagingSize -DataToMerge $dtExtensionCustomAttribute1 -DataDomain $ObjectDomain
                    }

                    if ($dtExtensionCustomAttribute2.DefaultView.Count -gt 0) {
                        $SuccessToDB = _SubmitMySQLChanges -ConnectionString $MySQLConnectionString -TargetTable $MySQLExtensionCustomAttribute2Table -PropertyMap $ExtensionCustomAttribute2PropertyMap `
                            -BatchSize $MySQLStagingSize -DataToMerge $dtExtensionCustomAttribute2 -DataDomain $ObjectDomain
                    }

                    if ($dtExtensionCustomAttribute3.DefaultView.Count -gt 0) {
                        $SuccessToDB = _SubmitMySQLChanges -ConnectionString $MySQLConnectionString -TargetTable $MySQLExtensionCustomAttribute3Table -PropertyMap $ExtensionCustomAttribute3PropertyMap `
                            -BatchSize $MySQLStagingSize -DataToMerge $dtExtensionCustomAttribute3 -DataDomain $ObjectDomain
                    }

                    if ($dtExtensionCustomAttribute4.DefaultView.Count -gt 0) {
                        $SuccessToDB = _SubmitMySQLChanges -ConnectionString $MySQLConnectionString -TargetTable $MySQLExtensionCustomAttribute4Table -PropertyMap $ExtensionCustomAttribute4PropertyMap `
                            -BatchSize $MySQLStagingSize -DataToMerge $dtExtensionCustomAttribute4 -DataDomain $ObjectDomain
                    }

                    if ($dtExtensionCustomAttribute5.DefaultView.Count -gt 0) {
                        $SuccessToDB = _SubmitMySQLChanges -ConnectionString $MySQLConnectionString -TargetTable $MySQLExtensionCustomAttribute5Table -PropertyMap $ExtensionCustomAttribute5PropertyMap `
                            -BatchSize $MySQLStagingSize -DataToMerge $dtExtensionCustomAttribute5 -DataDomain $ObjectDomain
                    }

                    if ($dtMobileDevices.DefaultView.Count -gt 0) {
                        $SuccessToDB = _SubmitMySQLChanges -ConnectionString $MySQLConnectionString -TargetTable $MySQLMobileTable -PropertyMap $MobilePropertyMap `
                            -BatchSize $MySQLStagingSize -DataToMerge $dtMobileDevices -DataDomain $ObjectDomain
                    }

                    $ExitCode = 0
                } catch { $SuccessToDB = $false }
            }#DB
        }#switch

        #In case export to database fails, fallback to CSV
        if ($ExportFormat -eq "DB" -and -not $SuccessToDB) {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Failed to export to database. Falling back to exporting CSV to $script:ExportLocation with $Delimiter delimiter"

            #Check for path/folder
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Checking for $script:ExportLocation"
            if (-not (Test-Path -Path $script:ExportLocation)) {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Creating folder structure for $script:ExportLocation"
                New-Item -Path $script:ExportLocation -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
            }

            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting export of CSV"
            $dtMailboxes | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_Mailboxes_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
            $dtEmailAddresses | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_EmailAddresses_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
            if ($dtExtensionCustomAttribute1.DefaultView.Count -gt 0) {
                $dtExtensionCustomAttribute1 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute1_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
            }

            if ($dtExtensionCustomAttribute2.DefaultView.Count -gt 0) {
                $dtExtensionCustomAttribute2 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute2_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
            }

            if ($dtExtensionCustomAttribute3.DefaultView.Count -gt 0) {
                $dtExtensionCustomAttribute3 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute3_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
            }

            if ($dtExtensionCustomAttribute4.DefaultView.Count -gt 0) {
                $dtExtensionCustomAttribute4 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute4_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
            }

            if ($dtExtensionCustomAttribute5.DefaultView.Count -gt 0) {
                $dtExtensionCustomAttribute5 | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_ExtensionAttribute5_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
            }

            if ($dtMobileDevices.DefaultView.Count -gt 0) {
                $dtMobileDevices | Export-Csv -Path "$script:ExportLocation\$($ObjectDomain)_MobileDevices_$strLogTimeStamp.csv" -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
            }
        }
    } else {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "No mailbox data to export"
    }#if/else

    $RunTime = ((Get-Date).ToUniversalTime() - $dateStartTimeStamp)
    $RunTime = '{0:00}:{1:00}:{2:00}:{3:00}.{4:00}' -f $RunTime.Days, $RunTime.Hours, $RunTime.Minutes, $RunTime.Seconds, $RunTime.Milliseconds
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Run time was $RunTime"
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exit code is $ExitCode"
} else {
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Script requirements not met:"
    
    if (-not $script:boolScriptIsModulesLoaded) { 
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Missing required PowerShell module(s) or could not load modules" 
    } elseif (-not $boolScriptIsEMS) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Not able to connect to $ObjectDomain"
    }#if/else
}#if/else

$VerbosePreference = "SilentlyContinue"
Exit $ExitCode

#endregion