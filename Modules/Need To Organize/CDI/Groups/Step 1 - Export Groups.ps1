#Requires -version 5.0


#region Parameters
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $false,
        HelpMessage = "Specify the distribution group domain name to be used in the output.")]
    [string]$GroupDomain,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the filter to be used for the distribution groups to collect. It will default to ''*'' if not specified.")]
    [string]$GroupFilter = "*",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify FQDN of an available on-premises Exchange server. It will default to 'Online' if not specified.")]
    [string]$ExchangeServer = "Online",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify only if you need to force SSL on for the remote PowerShell connection. It will default to false if not specified.")]
    [switch]$ExchangeForceSSL,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the remote PowerShell the specific authentication method. It will default to Basic if not specified.")]
    [ValidateSet("Basic", "Negotiate", "Kerberos")]
    [string]$ExchangeAuthMethod = "Basic",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the credential XML file to be used")]
    [string]$CredentialsPath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify to include unified groups")]
    [switch]$IncludeUnifiedGroups,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify to include group membership")]
    [switch]$IncludeMembership,
    
    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify to include group send-as permissions")]
    [switch]$IncludePermissions,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify FQDN of an Active Directory Domain Controller to use. This should be a DC in the root domain. It will default to use the current logon server if not specified.")]
    [string]$DomainController = "",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify to connect to GCC High tenants")]
    [switch]$GCCHigh,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the export location to be used")]
    [string]$ExportLocation
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
Set-Variable -Name ADModule -Option AllScope -Scope Script
Set-Variable -Name htLoggingPreference -Option AllScope -Scope Script -Value @{"InformationPreference" = $InformationPreference; `
        "WarningPreference" = $WarningPreference; "ErrorActionPreference" = $ErrorActionPreference; "VerbosePreference" = $VerbosePreference; "DebugPreference" = $DebugPreference
}
Set-Variable -Name verScript -Option AllScope -Scope Script -Value "5.1.2024.0109"

Set-Variable -Name boolScriptIsModulesLoaded -Option AllScope -Scope Script -Value $false
Set-Variable -Name ExitCode -Option AllScope -Scope Script -Value 1

Set-Variable -Name SendAsGUID -Option AllScope -Scope Script -Value "ab721a54-1e2f-11d0-9819-00aa0040529b"
Set-Variable -Name RecipientFilter -Option AllScope -Scope Script -Value "((RecipientType -eq 'UserMailbox') -or (RecipientType -eq 'MailUniversalSecurityGroup') `
    -or (RecipientType -eq 'MailUser'))"

New-Object System.Data.DataTable | Set-Variable dtRecipients -Option AllScope -Scope Script
New-Object System.Collections.ArrayList | Set-Variable arrPermissionsException -Option AllScope -Scope Script
New-Object System.Collections.ArrayList | Set-Variable arrPermissions -Option AllScope -Scope Script

New-Object System.Data.DataTable | Set-Variable dtGroups -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtEmailAddresses -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute1 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute2 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute3 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute4 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute5 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtManagedBy -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtAcceptMessagesOnlyFrom -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtAcceptMessagesOnlyFromDLMembers -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtAcceptMessagesOnlyFromSendersOrMembers -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtBypassModerationFromSendersOrMembers -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtGrantSendOnBehalfTo -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtModeratedBy -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtRejectMessagesFrom -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtRejectMessagesFromDLMembers -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtRejectMessagesFromSendersOrMembers -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtMembers -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtSendAsPermissions -Option AllScope -Scope Script

#Includes Groups, DLs, and DDLs
Set-Variable -Name arrGroupAttribs -Option AllScope -Scope Script -Value 'Guid', 'GroupType', 'SamAccountName', 'BypassNestedModerationEnabled', `
    'ManagedBy', 'MemberJoinRestriction', 'MemberDepartRestriction', 'ReportToManagerEnabled', 'ReportToOriginatorEnabled', `
    'SendOofMessageToOriginatorEnabled', 'AcceptMessagesOnlyFrom', 'AcceptMessagesOnlyFromDLMembers', 'AcceptMessagesOnlyFromSendersOrMembers', `
    'Alias', 'ArbitrationMailbox', 'BypassModerationFromSendersOrMembers', 'OrganizationalUnit', 'CustomAttribute1', 'CustomAttribute10', `
    'CustomAttribute11', 'CustomAttribute12', 'CustomAttribute13', 'CustomAttribute14', 'CustomAttribute15', 'CustomAttribute2', `
    'CustomAttribute3', 'CustomAttribute4', 'CustomAttribute5', 'CustomAttribute6', 'CustomAttribute7', 'CustomAttribute8', 'CustomAttribute9', `
    'ExtensionCustomAttribute1', 'ExtensionCustomAttribute2', 'ExtensionCustomAttribute3', 'ExtensionCustomAttribute4', `
    'ExtensionCustomAttribute5', 'DisplayName', 'EmailAddresses', 'GrantSendOnBehalfTo', 'ExternalDirectoryObjectId', `
    'HiddenFromAddressListsEnabled', 'LegacyExchangeDN', 'MaxSendSize', 'MaxReceiveSize', 'ModeratedBy', 'ModerationEnabled', `
    'EmailAddressPolicyEnabled', 'PrimarySmtpAddress', 'RecipientType', 'RecipientTypeDetails', 'RejectMessagesFrom', 'RejectMessagesFromDLMembers', `
    'RejectMessagesFromSendersOrMembers', 'RequireSenderAuthenticationEnabled', 'SimpleDisplayName', 'SendModerationNotifications', `
    'WindowsEmailAddress', 'MailTip', 'Identity', 'Name', 'DistinguishedName', 'WhenChangedUTC', `
    'WhenCreatedUTC', 'Id', 'IncludedRecipients', 'LdapRecipientFilter', 'Notes', 'RecipientContainer', 'RecipientFilter', `
    'BccBlocked', 'Description', 'ExchangeObjectId', 'HiddenGroupMembershipEnabled', 'DirectMembershipOnly', 'IsDirSynced'
#endregion

#region Complete Functions

Function ConvertTo-DataTable {
    <#
        .SYNOPSIS
            Converts an object into a PowerShell DataTable

        .DESCRIPTION
            This function will convert a PSObject into a PowerShell DataTable for interactions similar to SQL Server

        .PARAMETER InputObject
            This is a mandatory parameter which is the input object that will be converted into a DataTable

        .INPUTS
            [psobject]. You can pipe objects to this script.

        .OUTPUTS
            [System.Data.DataTable] This command will return a PowerShell DataTable

        .EXAMPLE
            $dtSourceData = ConvertTo-DataTable -Inputobject $csv_Import

            The preceding example creates a DataTable from the import CSV content

        .NOTES
            Version:
                - 5.1.2023.0725: 	New function. Adopted from https://github.com/RamblingCookieMonster/PowerShell
    #>
    [CmdletBinding()]
    [OutputType([System.Data.DataTable])]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [PSObject[]]$InputObject
    )
    
    begin {
        Function _gettype {
            param($type)
            
            $types = @(
                'System.Boolean',
                'System.Byte[]',
                'System.Byte',
                'System.Char',
                'System.Datetime',
                'System.Decimal',
                'System.Double',
                'System.Guid',
                'System.Int16',
                'System.Int32',
                'System.Int64',
                'System.Single',
                'System.UInt16',
                'System.UInt32',
                'System.UInt64')
            
            If ($types -contains $type) { return $type }
            ElseIf ($type -match "System.Collections.Generic.List") { return "System.Array" }
            ElseIf ($type -match "System.Collections.ArrayList") { return "System.Collections.ArrayList" }
            ElseIf ($type -match "MultiValuedProperty") { return "System.Collections.ArrayList" }
            Else { return "System.String" }
        }#function _gettype
        
        $NewDatatable = New-Object System.Data.DataTable
        $First = $true
    }#begin
    
    process {
        foreach ($Object in $InputObject) {
            $NewDataTableRow = $NewDatatable.NewRow()
            
            # Create columns
            foreach ($Property in $Object.PsObject.Properties) {
                $Name = $Property.Name
                $Value = $Property.Value
                if ($Property.TypeNameOfValue -match "List") {
                    $Value = @()
                    $Property.Value | ForEach-Object { $Value += $_ }
                }
                
                if ($First) {
                    $Column = New-Object System.Data.DataColumn
                    $Column.ColumnName = $Name
                    
                    $Column.DataType = [System.Type]::GetType($(_gettype $Property.TypeNameOfValue))
                    $Column.AllowDBNull = $true

                    #old way
                    #If it's not DBNull or Null, get the type
                    #if ($Value -isnot [System.DBNull] -and $null -ne $Value) {
                    #	$Column.DataType = [System.Type]::GetType($(_gettype $Property.TypeNameOfValue))
                    #} else {$Column.AllowDBNull = $true}
                    
                    [void]$NewDatatable.Columns.Add($Column)
                }#if first row

                if ($null -eq $Value) { $Value = [DBNull]::Value }
                
                if ($Value.getType().ToString() -eq "System.Collections.ArrayList") {
                    $NewDataTableRow.Item($Name) = [System.Collections.ArrayList]$Value
                } else { $NewDataTableRow.Item($Name) = $Value }
            }#foreach property
            
            [void]$NewDatatable.Rows.Add($NewDataTableRow)
            
            $First = $false
        }#foreach row
    }#process

    end {
        # Because PowerShell handles returning objects stupidly
        return @(, $NewDatatable)
    }#end
}#Function ConvertTo-DataTable

Function Connect-Exchange {
    <#
        .SYNOPSIS
            This cmdlet will assist with connections for Exchange

        .DESCRIPTION
            This cmdlet will help shortcut the ability to connect to Exchange on-premises and Online. 
            You can leverage saved credential file, prompt for credentials, or a certificate based connection.

        .PARAMETER CredentialFile
            This is an optional parameter for the file path for the Exchange credentials. 
            If not specified, it will prompt for credentials.

        .PARAMETER Credential
            This is an optional parameter for the Exchange credentials. 
            If not specified, it will prompt for credentials.

        .PARAMETER Server
            This is an optional parameter for the FQDN of an available Exchange on-premises server.
            If not specified, it will default to 'Online'.
            
        .PARAMETER UseSSL
            This is an optional parameter to force use of HTTPS. 
            If not specified, it will try both but may fail/hang depending on the environment.

        .PARAMETER AuthenticationMethod
            This is an optional parameter to force authentication method to use. 
            If not specified, it will try all of them but may fail/hang depending on the environment.

        .PARAMETER Modern
            This is an optional parameter to force modern authentication to Exchange Online.  
            Requires the Exchange Online V2 module which will be installed if missing.
            If not specified, it will not prompt for modern authentication and may fail/hang depending on the environment.

        .PARAMETER Certificate
            This is an optional parameter to force certificate authentication to Exchange Online. 
            Requires the Exchange Online V2 module which will be installed if missing.

        .PARAMETER ConnectionFilePath
            This is an optional parameter to use a predefined certificate connection.
            File can be created by New-EXOCerConnection.

        .PARAMETER AppID
            This is an optional parameter to specify the Azure AD application ID to use for the certificate connection.

        .PARAMETER TenantName
            This is an optional parameter to specify the tenant name to use for the certificate connection.
                
        .PARAMETER CertificateThumbprint
            This is an optional parameter to specify the certificate thumbprint to use for the certificate connection.

        .PARAMETER ConnectionPrefix
            This is an optional parameter to configure the connection to use a custom prefix.
            This should be used in very specific cases and expert level only.

        .PARAMETER IgnoreExistingSessions
            This is an optional parameter to leave existing unrelated PowerShell sessions connected. 
            This will still remove existing Exchange sessions to the same server.
            This should be used in very specific cases and expert level only.
        
        .PARAMETER GCCHigh
            This is an optional parameter to connect to Microsoft 365 GCC High Tenant

        .INPUTS
            None. You cannot pipe objects to this script.

        .OUTPUTS
            None.

        .EXAMPLE
            Connect-Exchange -Server server1.contoso.com -UseSSL

            The preceding example will attempt a connection to server1.contoso.com using only SSL.

        .EXAMPLE
            Connect-Exchange -CredentialFile C:\Temp\creds.xml -Server server1.contoso.com

            The preceding example will attempt a connection using all authentication methods and SSL/non-SSL to server1.contoso.com using the credentials in the supplied file.

        .EXAMPLE
            Connect-Exchange -Certificate -ConnectionFilePath C:\Temp\contoso_EXOCBA.xml

            The preceding example will attempt a certificate based connection to Exchange Online using the information in the supplied connection file.

        .EXAMPLE
            Connect-Exchange -Certificate -AppID <APPID> -TenantName contoso -CertificateThumbprint <THUMBPRINT>

            The preceding example will attempt a certificate based connection to Exchange Online for contoso.onmicrosoft.com using the <APPID> and <THUMBPRINT> information.

        .NOTES
            Version:
                - 5.1.2023.0815:    New function
                - 5.1.2023.1002:    Updated version check
                - 5.1.2023.1019:    Added parameter and logic for GCC High environment
    #>
    [CmdletBinding(DefaultParameterSetName = 'Modern')]
    [OutputType([System.Void])]
    Param(
        [Parameter(ParameterSetName = 'Basic', Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the file path for the Exchange credentials. It will prompt for credentials if not specified.")]
        [string]$CredentialFile,

        [Parameter(ParameterSetName = 'Basic', Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the Exchange credentials. It will prompt for credentials if not specified.")]
        [PSCredential]$Credential,
        
        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify FQDN of an available on-premises Exchange server. It will default to Exchange Online if not specified.")]
        [string]$Server = "Online",
        
        [Parameter(ParameterSetName = 'Basic', Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify to force use of HTTPS. If not specified, it will try both but may fail/hang depending on the environment")]
        [switch]$UseSSL,
    
        [Parameter(ParameterSetName = 'Basic', Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify authentication method to use. If not specified, it will try all of them but may fail/hang depending on the environment.")]
        [ValidateSet("Basic", "Negotiate", "Kerberos")]
        [string]$AuthenticationMethod,
    
        [Parameter(ParameterSetName = 'Modern', Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify to use Modern for Exchange Online if a modern authentication prompt is required. If not specified, it will try to connect without modern authentication but may fail/hang depending on the environment.")]
        [Alias('MFA')]
        [switch]$Modern,

        [Parameter(ParameterSetName = 'CertFile', Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify to use a certificate for Exchange Online")]
        [Parameter(ParameterSetName = 'CertInfo', Mandatory = $true)]
        [switch]$Certificate,

        [Parameter(ParameterSetName = 'CertFile', Mandatory = $true, ValueFromPipeline = $false, 
            HelpMessage = "Specify the certificate connection file to use for the connection")]
        [string]$ConnectionFilePath,

        [Parameter(ParameterSetName = 'CertInfo', Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the Azure AD application ID to use for the connection")]
        [string]$AppID,

        [Parameter(ParameterSetName = 'CertInfo', Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the tenant name to use for the connection")]
        [string]$TenantName,

        [Parameter(ParameterSetName = 'CertInfo', Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the certificate thumbprint to use for the connection")]
        [string]$CertificateThumbprint,
        
        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify if the connection needs to include a prefix. This would allow connection to multiple environments within the same PowerShell window. WARNING: EXPERT ONLY!")]
        [string]$ConnectionPrefix = "",
        
        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify to leave existing unrelated sessions connected. This will still remove existing Exchange sessions to the same server. WARNING: EXPERT ONLY!")]
        [switch]$IgnoreExistingSessions,
        
        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify to connect to Microsoft 365 GCC High Tenant")]
        [switch]$GCCHigh
    )

    begin {
        $InformationPreference = "Continue"
        if ($DebugPreference -eq "Confirm" -or $DebugPreference -eq "Inquire") { $DebugPreference = "Continue" }
        $htLoggingPreference = @{"InformationPreference" = $InformationPreference; "WarningPreference" = $WarningPreference; `
                "ErrorActionPreference" = $ErrorActionPreference; "VerbosePreference" = $VerbosePreference; "DebugPreference" = $DebugPreference
        }
        
        $MinimumModuleVersion = "3.2.0"

        if ($IgnoreExistingSessions) {
            Out-Log -LoggingPreference $htLoggingPreference -Type Verbose -WriteBackToHost -Message "Ignoring existing sessions"
            $sessionID = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.ComputerName -eq "$Server" }
            
            if ($sessionID) {
                Out-Log -LoggingPreference $htLoggingPreference -Type Verbose -WriteBackToHost -Message "Found RPS session to same server. Must remove to continue."
                $sessionID | Remove-PSSession -ErrorAction Continue
                Get-Module | Where-Object { $_.Description -like "*$Server*" } | Remove-Module
            }#if session already connected to same server but broken
        } else {
            Out-Log -LoggingPreference $htLoggingPreference -Type Verbose -WriteBackToHost -Message "Looking at existing sessions or connections"
            
            $PSSessionID = Get-PSSession -ErrorAction SilentlyContinue | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }
            
            if ($PSSessionID) {
                Out-Log -LoggingPreference $htLoggingPreference -Type Verbose -WriteBackToHost -Message "Existing RPS session found, must remove to continue"
                
                $PSSessionID | Remove-PSSession -ErrorAction Continue
            }#if RPS session already connected
            
            if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) { $RESTSession = Get-ConnectionInformation -ErrorAction SilentlyContinue }

            if ($RESTSession) {
                Out-Log -LoggingPreference $htLoggingPreference -Type Verbose -WriteBackToHost -Message "Existing REST session found, must remove to continue"
                
                Disconnect-ExchangeOnline -ErrorAction Continue -Confirm:$false
            }#if REST session already connected
        }#if/else

        #Connection parameter splat
        $ConnectionParams = @{}

        if ($ConnectionPrefix -ne "") {
            $ConnectionParams.Add("Prefix", $ConnectionPrefix)
        }#if

        if ($Server -eq "Online") {
            $ModuleImported = Get-Module -Name "ExchangeOnlineManagement" -ErrorAction Stop -Verbose:$false

            if ($ModuleImported) {
                if ($ModuleImported.Version -le $MinimumModuleVersion) {
                    Out-Log -LoggingPreference $htLoggingPreference -Type Verbose -WriteBackToHost -Message "ExchangeOnlineManagement module version($($ModuleImported.Version)) meets minimum version($MinimumModuleVersion) requirements and is loaded"
                }#if
            } else {
                Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message "ExchangeOnlineManagement module missing"
                Break
            }#if/else

            if ($Certificate) {
                if ($ConnectionFilePath -and (Test-Path -Path $ConnectionFilePath)) {
                    $CertCredentials = Import-Clixml $ConnectionFilePath
    
                    $ConnectionParams.Add("AppID", $CertCredentials.AppID)
                    $ConnectionParams.Add("CertificateThumbprint", $CertCredentials.CertificateThumbprint)
                    $ConnectionParams.Add("Organization", $CertCredentials.Organization)
                } else {      
                    if ($TenantName -notlike "*.onmicrosoft.com") {
                        $TenantName = "$TenantName.onmicrosoft.com"
                    }#if
                    
                    $ConnectionParams.Add("AppID", $AppID)
                    $ConnectionParams.Add("CertificateThumbprint", $CertificateThumbprint)
                    $ConnectionParams.Add("Organization", $TenantName)
                }#if/else
            }#if

            $ConnectionParams.Add("ShowBanner", $false)
        } else {
            $SessionOptionsParams = @{
                "SkipCACheck"         = $true
                "SkipCNCheck"         = $true
                "SkipRevocationCheck" = $true
                "OpenTimeout"         = 20000
            }

            $ConnectionParams = @{
                "ConfigurationName" = "Microsoft.Exchange"
                "AllowRedirect"     = $true
                "ErrorAction"       = "SilentlyContinue"
            }

            if ($UseSSL) {
                $ConnectionParams["ConnectionUri"] = "https://$Server/powershell/"
            } else {
                $ConnectionParams["ConnectionUri"] = "http://$Server/powershell/"
            }#if/else
            
            if ($AuthenticationMethod) {
                $ConnectionParams["Authentication"] = $AuthenticationMethod
            } else {
                $ConnectionParams["Authentication"] = "Basic"
            }#if/else
        } #if/else

        if (-not ($PSCmdlet.ParameterSetName -eq "Modern" -or $Certificate)) {
            if ($CredentialFile -and (Test-Path -Path $CredentialFile)) {
                $ExchangeCredentials = Import-Clixml $CredentialFile
            } elseif ($Credential) {
                $ExchangeCredentials = $Credential
            } else {
                Out-Log -LoggingPreference $htLoggingPreference -Type Warning -WriteBackToHost -Message "Credential file not specified or invalid"
                $ExchangeCredentials = Get-Credential
            }#otherwise prompt

            $ConnectionParams.Add("Credential", $ExchangeCredentials)
        }#if

        if ($GCCHigh) {
            $ConnectionParams.Add("ExchangeEnvironmentName", "O365USGovGCCHigh")
        }

        #Required for situations where no authentication type is specified, it will fail on the import-module command otherwise
        if ($AuthenticationMethod -eq '' -or $AuthenticationMethod -eq $null) { Remove-Variable AuthenticationMethod }
    }#begin
     
    process {
        if ($Server -eq "Online") {
            try {
                if ($Certificate) {
                    try {
                        Connect-ExchangeOnline @ConnectionParams -ErrorAction Stop
                        Out-Log -LoggingPreference $htLoggingPreference -Type Information -WriteBackToHost -Message "You have successfully connected to Exchange Online with certificate"
                    } catch {
                        Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message "No or invalid Exchange Online Certificate-Based connection found"
                        Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message "To create the connection, run New-EXOCertConnection first"
                    }#try/catch
                } else {
                    #Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message "$($ConnectionParams | Out-String)"
                    #$global:Splat = $ConnectionParams
                    Connect-ExchangeOnline @ConnectionParams
                }#if/else
            } catch {
                Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message "Error while trying to connect to Exchange Online"
                Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message $_
            }#try/catch
        } else {
            try	{
                $SessionOptions = New-PSSessionOption @SessionOptionsParams
                $ConnectionParams.Add("SessionOption", $SessionOptions)

                Out-Log -LoggingPreference $htLoggingPreference -Type Information -WriteBackToHost -Message "Attempting to $($ConnectionParams["ConnectionUri"]) and $($ConnectionParams["Authentication"])"
                $session = New-PSSession @ConnectionParams
                
                #This will try Basic auth twice if auth method isn't specified
                if (-not $session -and -not ($UseSSL -or $AuthenticationMethod)) {
                    $AuthTypes = ($MyInvocation.MyCommand.Parameters['AuthenticationMethod'].Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }).ValidValues
                    ForEach ($type in $AuthTypes) {
                        $ConnectionParams["Authentication"] = $type
                        Out-Log -LoggingPreference $htLoggingPreference -Type Information -WriteBackToHost -Message "Attempting to $($ConnectionParams["ConnectionUri"]) and $($ConnectionParams["Authentication"])"
                        
                        $session = New-PSSession @ConnectionParams
                        if ($session) { break }
                    }#foreach
                    
                    if (-not $session) {
                        $ConnectionParams["ConnectionUri"] = "https://$Server/powershell/"

                        ForEach ($type in $AuthTypes) {
                            $ConnectionParams["Authentication"] = $type
                            Out-Log -LoggingPreference $htLoggingPreference -Type Information -WriteBackToHost -Message "$(Get-Date -Format 'MM/dd/yyyy HH:mm:ss:fff') Attempting to $($ConnectionParams["ConnectionUri"]) and $($ConnectionParams["Authentication"])"
                            
                            $session = New-PSSession @ConnectionParams
                            if ($session) { break }
                        }#foreach
                    }#if
                }#if/elseif/else
                
                if ($session) {
                    $ImportParams = @{
                        "Global"      = $true
                        "ErrorAction" = "Stop"
                    }

                    if ($ConnectionPrefix -ne "") {
                        $ImportParams["Prefix"] = $ConnectionPrefix
                    }
                    Import-Module (Import-PSSession $session -AllowClobber) @ImportParams
                } else {
                    Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message "Unable to establish a session with supplied parameters. Please check server and parameters"
                    Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message "If Exchange On-premises, it may be necessary to enable an authentication method"
                    Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message "Example:  Set-PowerShellVirtualDirectory -Identity ""PowerShell (Default Web Site)"" -WindowsAuthentication:`$true"
                    Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message "It is also possible to enable the authentication directly on the virtual directory within IIS"
                }#if/else
            } catch {
                Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message "Error while trying to connect to Exchange On-Premises"
                Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message $_
            }#try/catch
        }#if/else
        return
    }#process

    end {
    }#end
}#Function Connect-Exchange

Function Out-Log {
    <# 
        .SYNOPSIS 
            Write to a log file in a format that takes advantage of the CMTrace.exe log viewer that comes with SCCM.
        
        .DESCRIPTION 
            Output strings to a log file that is formatted for use with CMTrace.exe and also writes back to the host.
            
            The severity of the logged line can be set as: 
            
                    2-Error
                    3-Warning
                    4-Verbose
                    5-Debug
                    6-Information

            Warnings will be highlighted in yellow. Errors are highlighted in red. 
            
            ** Verbose and Debug will not output to the log file unless the script was ran with the -Verbose/-Debug parameter **

            The tools to view the log: 
            CM Trace - https://www.microsoft.com/en-us/download/details.aspx?id=50012 or the Installation directory on Configuration Manager 2012 Site Server - <Install Directory>\tools\ 
            OneTrace - new preview feature. Link provided once it is available

        .PARAMETER Logfile
            Specify the file path for the log file. It will default to %temp%\Sterling.log if not specified.

        .PARAMETER Message
            Specify the object or string for the log

        .PARAMETER Type
            Specify the severity of message to log. It will default to Information. Options include 'Warning','Error','Verbose','Debug', 'Information'

        .PARAMETER WriteBackToHost
            Specify whether or not to write the message back to the console. It will default to false if not specified.

        .PARAMETER LoggingPreference
            Specify a hashtable of the preference variables so that they can be honored. It will use the default PowerShell values if not specified.
            Example: @{"InformationPreference"=$InformationPreference;"WarningPreference"=$WarningPreference;"ErrorActionPreference"=$ErrorActionPreference;"VerbosePreference"=$VerbosePreference;"DebugPreference"=$DebugPreference}

        .PARAMETER ForceWriteToLog
            Specify whether or not to force the message to be written to the log. It will default to false if not specified.

        .INPUTS
            None. You cannot pipe objects to this script.

        .OUTPUTS
            None.

        .EXAMPLE 
            Try {
                Get-Process -Name DoesnotExist -ea stop
            }
            Catch {
                Out-Log -Logfile "C:\output\logfile.log -Message $_ -Type Error
            }
            
            This will write a line to the logfile.log file in c:\output\logfile.log. It will state the errordetails in the log file 
            and highlight the line in Red. It will also write back to the host in a friendlier red on black message than
            the normal error record.
        
        .EXAMPLE
            $VerbosePreference = Continue
            Out-Log -Message "This is a verbose message." -Type Verbose -VerbosePreference $VerbosePreference

            This example will write a verbose entry into the Sterling.log log file and also write back to the host. The Out-Log will obey
            the preference variables.

        .EXAMPLE
            Out-Log -Message "This is an informational message" -WritebacktoHost

            This example will write the informational message to the log but write back to the host.

        .EXAMPLE
            Function Test{
                [CmdletBinding()]
                Param()
                Out-Log -VerbosePreference $VerbosePreference -Message "This is a verbose message" -Type Verbose
            }
            Test -Verbose

            This example shows how to use Out-Log inside a function and then call the function with the -verbose switch.
            The Out-Log function will then print the verbose message.

        .NOTES
            Version:
                - 5.1.2023.0725:	New function. Adopted from
                                        https://wolffhaven.gitlab.io/wolffhaven_icarus_test/powershell/write-cmtracelog-dropping-logs-like-a-boss/
                                        https://adamtheautomator.com/building-logs-for-cmtrace-powershell/
                - 5.1.2024.0215:    New format for module
                - 5.1.2024.0318:    Removed output definition to fix writetohost
                - 5.1.2024.0501:    Fixed bug with warning
    #>
    [CmdletBinding()]
    Param( 
        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the file path for the log file. It will default to %temp%\Sterling.log if not specified.")]
        [string]$Logfile = "$env:TEMP\Sterling.log",
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the object or string for the log")]
        $Message,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the severity of message to log.")]
        [ValidateSet('Warning', 'Error', 'Verbose', 'Debug', 'Information')] 
        [string]$Type = "Information",

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify whether or not to write the message back to the console. It will default to false if not specified.")]
        [switch]$WriteBackToHost,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify a hashtable of the preference variables so that they can be honored. It will use the default PowerShell values if not specified.")]
        [hashtable]$LoggingPreference = @{"InformationPreference" = $InformationPreference; `
                "WarningPreference" = $WarningPreference; "ErrorActionPreference" = $ErrorActionPreference; "VerbosePreference" = $VerbosePreference; "DebugPreference" = $DebugPreference
        },

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify whether or not to force the message to be written to the log. It will default to false if not specified.")]
        [switch]$ForceWriteToLog
    )#Param

    begin {
        #Update variables based on parameters
        $Type = $Type.ToUpper()
        
        #Set the order 
        switch ($Type) {
            'Warning' { $severity = 2 }#Warning
            'Error' { $severity = 3 }#Error
            'Verbose' { $severity = 4 }#Verbose
            'Debug' { $severity = 5 }#Debug
            'Information' { $severity = 6 }#Information
        }#switch

        #Script defaults
        $TimeGenerated = (Get-Date).ToUniversalTime()
        $userContext = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $LineTemplate = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="" file="{6}">'
        
        #Attempt password redaction
        if ($Message -match "Password=") {
            $Message = $Message -replace "Password=(?<Password>((\`"[\S\s]+\`")|(\{[\S\s]+\})|([\S-[;$]]+)(;|$)))", "Password=********;"
        }

        #open log stream
        #if ($null -eq $global:ErrorLogStream -or -not $global:ErrorLogStream.BaseStream) {
        If (-not (Test-Path -Path (Split-Path $LogFile))) { New-Item (Split-Path $LogFile) -ItemType Directory | Out-Null }
            
        $global:ErrorLogStream = New-Object System.IO.StreamWriter $Logfile, $true, ([System.Text.Encoding]::UTF8)
        $global:ErrorLogStream.AutoFlush = $true
            
        #}

        #Need the callstack information to get the details about the calling script
        $CallStack = Get-PSCallStack | Select-Object -Property Command, Location, ScriptName, ScriptLineNumber
        if (($null -ne $CallStack.Count) -or (($CallStack.Command -ne '<ScriptBlock>') -and ($CallStack.Location -ne '<No file>') -and ($null -ne $CallStack.ScriptName))) {
            if ($CallStack.Count -eq 1) {
                $CallingInfo = $CallStack[0]
            } elseif ($CallStack.Count -eq 2) {
                $CallingInfo = $CallStack[($CallStack.Count - 1)]
            } else {
                $CallingInfo = $CallStack[($CallStack.Count - 2)]
            }#need only or the second to the last one if multiple returned
        } else {
            Write-Error -Message 'No callstack detected' -Category 'InvalidData'
        }#if callstack info found
        
        $global:TestStack = $CallStack
        $global:CallInfo = $CallingInfo
    }

    process {
        #Switch statement to write out to the log and/or back to the host.
        switch ($severity) {
            2 {
                if ($LoggingPreference['WarningPreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                    #Build log message
                    $LogMessage = $Type + ": " + $Message
                    $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                        "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName

                    $logline = $LineTemplate -f $LineContent
                    $global:ErrorLogStream.WriteLine($logline)
                }#if silentlycontinue and not forced to write, don't write to the log
                
                #Write back to the host if $Writebacktohost is true.
                if ($WriteBackToHost) {
                    switch ($LoggingPreference['WarningPreference']) {
                        'Continue' { $WarningPreference = 'Continue'; Write-Warning -Message "$Message"; $WarningPreference = '' }
                        'Stop' { $WarningPreference = 'Stop'; Write-Warning -Message "$Message"; $WarningPreference = '' }
                        'Inquire' { $WarningPreference = 'Inquire'; Write-Warning -Message "$Message"; $WarningPreference = '' }
                    }#switch
                }#if writeback
            }#Warning
            3 {  
                #This if statement is to catch the two different types of errors that may come through. 
                #A normal terminating exception will have all the information that is needed, if it's a user generated error by using Write-Error,
                #then the else statment will setup all the information we would like to log.   
                if ($Message.Exception.Message) {
                    if ($LoggingPreference['ErrorActionPreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                        #Build log message                                      
                        $LogMessage = $Type + ": " + [string]$Message.Exception.Message + " Command: '" + [string]$Message.InvocationInfo.MyCommand + `
                            "' Line: '" + [string]$Message.InvocationInfo.Line.Trim() + "'"
                        $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                            "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName

                        $logline = $LineTemplate -f $LineContent
                        $global:ErrorLogStream.WriteLine($logline)
                    }#if silentlycontinue and not forced to write, don't write to the log

                    #Write back to the host if $Writebacktohost is true.
                    if ($WriteBackToHost) {
                        switch ($LoggingPreference['ErrorActionPreference']) {
                            'Stop' { $ErrorActionPreference = 'Stop'; $Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)"); Write-Error $Message -ErrorAction 'Stop'; $ErrorActionPreference = '' }
                            'Inquire' { $ErrorActionPreference = 'Inquire'; $Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)"); Write-Error $Message -ErrorAction 'Inquire'; $ErrorActionPreference = '' }
                            'Continue' { $ErrorActionPreference = 'Continue'; $Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)"); $ErrorActionPreference = '' }
                            'Suspend' { $ErrorActionPreference = 'Suspend'; $Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)"); Write-Error $Message -ErrorAction 'Suspend'; $ErrorActionPreference = '' }
                        }#switch
                    }#if writeback
                }#if standard error
                else {
                    if ($LoggingPreference['ErrorActionPreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                        #Custom error message so build out the Exception object
                        [System.Exception]$Exception = $Message
                        [String]$ErrorID = 'Custom Error'
                        [System.Management.Automation.ErrorCategory]$ErrorCategory = [Management.Automation.ErrorCategory]::WriteError
                        $ErrorRecord = New-Object Management.automation.errorrecord ($Exception, $ErrorID, $ErrorCategory, $Message)
                        $Message = $ErrorRecord

                        #Build log message                
                        $LogMessage = $Type + ": " + [string]$Message.Exception.Message
                        
                        $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                            "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName

                        $logline = $LineTemplate -f $LineContent
                        $global:ErrorLogStream.WriteLine($logline)
                    }#if silentlycontinue and not forced to write, don't write to the log
                        
                    #Write back to the host if $Writebacktohost is true.
                    if ($WriteBackToHost) {
                        switch ($LoggingPreference['ErrorActionPreference']) {
                            'Stop' { $ErrorActionPreference = 'Stop'; $Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)"); Write-Error $Message -ErrorAction 'Stop'; $ErrorActionPreference = '' }
                            'Inquire' { $ErrorActionPreference = 'Inquire'; $Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)"); Write-Error $Message -ErrorAction 'Inquire'; $ErrorActionPreference = '' }
                            'Continue' { $ErrorActionPreference = 'Continue'; $Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)"); $ErrorActionPreference = '' }
                            'Suspend' { $ErrorActionPreference = 'Suspend'; $Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)"); Write-Error $Message -ErrorAction 'Suspend'; $ErrorActionPreference = '' }
                        }#switch
                    }#if writeback
                }#else custom error
            }#Error
            4 {  
                if ($LoggingPreference['VerbosePreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                    #Build log message                
                    $LogMessage = $Type + ": " + $Message
                    $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                        "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName

                    $logline = $LineTemplate -f $LineContent
                    $global:ErrorLogStream.WriteLine($logline)
                }#if silentlycontinue and not forced to write, don't write to the log

                #Write back to the host if $Writebacktohost is true.
                if ($WriteBackToHost) {
                    switch ($LoggingPreference['VerbosePreference']) {
                        'Continue' { $VerbosePreference = 'Continue'; Write-Verbose -Message "$Message"; $VerbosePreference = '' }
                        'Inquire' { $VerbosePreference = 'Inquire'; Write-Verbose -Message "$Message"; $VerbosePreference = '' }
                        'Stop' { $VerbosePreference = 'Stop'; Write-Verbose -Message "$Message"; $VerbosePreference = '' }
                    }#switch
                }#if writeback
            }#Verbose
            5 {  
                if ($LoggingPreference['DebugPreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                    #Build log message                
                    $LogMessage = $Type + ": " + $Message
                    $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                        "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName

                    $logline = $LineTemplate -f $LineContent
                    $global:ErrorLogStream.WriteLine($logline)
                }#if silentlycontinue and not forced to write, don't write to the log
                
                #Write back to the host if $Writebacktohost is true.
                if ($WriteBackToHost) {
                    switch ($LoggingPreference['DebugPreference']) {
                        'Continue' { $DebugPreference = 'Continue'; Write-Debug -Message "$Message"; $DebugPreference = '' }
                        'Inquire' { $DebugPreference = 'Inquire'; Write-Debug -Message "$Message"; $DebugPreference = '' }
                        'Stop' { $DebugPreference = 'Stop'; Write-Debug -Message "$Message"; $DebugPreference = '' }
                    }#switch
                }#if writeback
            }#Debug
            6 {  
                if ($LoggingPreference['InformationPreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                    #Build log message                
                    $LogMessage = $Type + ": " + $Message
                    $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                        "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName

                    $logline = $LineTemplate -f $LineContent
                    $global:ErrorLogStream.WriteLine($logline)
                }#if silentlycontinue and not forced to write, don't write to the log

                #Write back to the host if $Writebacktohost is true.
                if ($WriteBackToHost) {
                    switch ($LoggingPreference['InformationPreference']) {
                        'Continue' { $InformationPreference = 'Continue'; Write-Information -Message "INFORMATION: $Message"; $InformationPreference = '' }
                        'Inquire' { $InformationPreference = 'Inquire'; Write-Information -Message "INFORMATION: $Message"; $InformationPreference = '' }
                        'Stop' { $InformationPreference = 'Stop'; Write-Information -Message "INFORMATION: $Message"; $InformationPreference = '' }
                        'Suspend' { $InformationPreference = 'Suspend'; Write-Information -Message "INFORMATION: $Message"; $InformationPreference = '' }
                    }#switch
                }#if writeback
            }#Information
        }#Switch
    }#process

    end {
        #Close log files while we are waiting
        if ($null -ne $global:ErrorLogStream) {
            $global:ErrorLogStream.Close()
            $global:ErrorLogStream.Dispose()
        }
    }#end
}#Function Out-Log
Function _ConfirmScriptRequirements {
    <#
    .SYNOPSIS
        Verifies that all necessary requirements are present for the script and return true/false
    .EXAMPLE
        $valid = _ConfirmScriptRequirements

        This would check the script requirements and set $valid to true/false based on the results
    .NOTES
        Version:
        - 5.1.2023.0726:    New function
        - 5.1.2023.1213:    Updated to use Sterling Connect-Exchange
        - 5.1.2024.0105:    Updated to allow for GCCHigh connections
        - 5.1.2024.0109:    Updated to add send-as permission collection
    #>
    [CmdletBinding()]
    Param()

    begin {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Set-Variable -Name strBaseLocation -Option AllScope -Scope Script -Value $(_GetScriptDirectory -Path)
        Set-Variable -Name ADModule -Option AllScope -Scope Script -Value "$strBaseLocation\Microsoft.ActiveDirectory.Management.dll"

        Write-Debug -Message "Starting _ConfirmScriptRequirements"
        try {
            Write-Host "Loading Sterling PowerShell module`r"

            if (Get-Module -ListAvailable Sterling -Verbose:$false) {
                Import-Module Sterling -ErrorAction Stop -Verbose:$false
                $script:boolScriptIsModulesLoaded = $true
            } else {
                Write-Warning "Missing Sterling PowerShell module`r"
                $script:boolScriptIsModulesLoaded = $false
            }#if/else
        } catch {
            Write-Error "Unable to load Sterling PowerShell module`r"
            Write-Error $_

            $script:boolScriptIsModulesLoaded = $false
        }#try/catch

        try {
            Write-Host "Loading ExchangeOnlineManagement PowerShell module`r"

            if (Get-Module -ListAvailable ExchangeOnlineManagement -Verbose:$false) {
                $global:VerbosePreference = "SilentlyContinue"

                Import-Module ExchangeOnlineManagement -ErrorAction Stop -Verbose:$false
                $script:boolScriptIsModulesLoaded = $true

                if ($htLoggingPreference['VerbosePreference'] -eq "Continue") { $global:VerbosePreference = "Continue" }#if
            } else {
                Write-Warning "Missing ExchangeOnlineManagement PowerShell module`r"
                $script:boolScriptIsModulesLoaded = $false
            }#if/else
        } catch {
            Write-Error "Unable to load ExchangeOnlineManagement PowerShell module`r"
            Write-Error $_

            $script:boolScriptIsModulesLoaded = $false
        }#try/catch

        if ($ExchangeServer -ne "Online") {
            try {
                Write-Host "Loading Active Directory PowerShell module`r"
                if (-not (Test-Path -Path $script:ADModule)) {
                    Write-Host "AD Module DLL is missing`r"
                    $script:boolScriptIsModulesLoaded = $false
                } else {
                    Import-Module $script:ADModule -ErrorAction Stop -Verbose:$false
                    $script:boolScriptIsModulesLoaded = $true
                }
            } catch {
                Write-Error "Unable to load Active Directory DLL module`r"
                Write-Error $_
    
                $script:boolScriptIsModulesLoaded = $false
            }#try/catch
        }#if

        Set-Variable -Name strBaseLocation -Option AllScope -Scope Script -Value $(_GetScriptDirectory -Path)
        Set-Variable -Name strLogFile -Option ReadOnly -Force -Scope Script -Value "$script:strBaseLocation\Logging\$script:strLogTimeStamp-$((_GetScriptDirectory -Leaf).Replace(".ps1",'')).log"

        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Debug -WriteBackToHost -Message "Starting _ConfirmScriptRequirements"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Script version $verScript starting"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: InformationPreference = $InformationPreference"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ErrorActionPreference = $ErrorActionPreference"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: VerbosePreference = $VerbosePreference"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: DebugPreference = $DebugPreference"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: GroupDomain = $GroupDomain"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: GroupFilter = $GroupFilter"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ExchangeServer = $ExchangeServer"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ExchangeForceSSL = $ExchangeForceSSL"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ExchangeAuthMethod = $ExchangeAuthMethod"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: CredentialsPath = $CredentialsPath"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: GCCHigh = $GCCHigh"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludeUnifiedGroups = $IncludeUnifiedGroups"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludeMembership = $IncludeMembership"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludePermissions = $IncludePermissions"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: DomainController = $DomainController"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ExportLocation = $ExportLocation"
    }#begin
    
    process {
        if ($script:boolScriptIsModulesLoaded) {
            try {
                $global:VerbosePreference = "SilentlyContinue"

                $ConnectSplat = @{
                    "GCCHigh" = $GCCHigh
                }

                if ($ExchangeServer -ne "Online") {
                    $ConnectSplat.Add("Server", $ExchangeServer)
                    $ConnectSplat.Add("UseSSL", $ExchangeForceSSL)
                    $ConnectSplat.Add("AuthenticationMethod", $ExchangeAuthMethod)
                }

                if ($CredentialsPath -and (Test-Path -Path $CredentialsPath)) {
                    $ConnectSplat.Add("Credential", $(Import-Clixml $CredentialsPath))
                } else {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "CredentialsPath specified but does not exist"
                    $script:boolScriptIsFilesExist = $false
                }#if/else
                
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Connecting to Exchange Online"
                Connect-Exchange @ConnectSplat
                                
                if ($htLoggingPreference['VerbosePreference'] -eq "Continue") { $global:VerbosePreference = "Continue" }#if

                if ($DomainController -eq '') {
                    $DomainController = $Env:LOGONSERVER.replace("\\", "") + "." + $Env:USERDNSDOMAIN
                }#if

                $Connection = Get-ConnectionInformation -ErrorAction SilentlyContinue
                if ($IncludePermissions -and -not $Connection) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Using $DomainController for Active Directory"
                    $Forest = ($DomainController.substring($DomainController.IndexOf(".") + 1)).Replace(".", "_")
                    New-PSDrive -Name $Forest -Scope Script -Root "" -PSProvider ActiveDirectory -Server $DomainController
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "PSDrive created"
                }#if
            } catch {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error verifying script requirements"
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
                return $false
            }#try/catch
        }#if

        #Final check
        if ($script:boolScriptIsModulesLoaded) { return $true }
        else { return $false }
    }#process

    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Debug -WriteBackToHost -Message "Finishing _ConfirmScriptRequirements"
    }#end
}#function _ConfirmScriptRequirements

Function _GetDistributionGroupInfo {
    <#
    .SYNOPSIS
        Collects the necessary distribution group cache and returns a datatable with the results
    .PARAMETER Session
        The PSSession to run the command against
    .PARAMETER Filter
        The optional parameter for a filter to be used when querying groups
   .PARAMETER GroupAttributes
        Specify the array of group attributes to return with the DataTable.
    .EXAMPLE
        $dtGroups = _GetDistributionGroupInfo -Session $session GroupAttributes $GroupAttribs
    
        This would get all groups from the PowerShell $session and return attributes $GroupAttribs to $dtGroups
    .NOTES
        Version:
            - 5.1.2023.0726:    New function
            - 5.1.2024.0110:    Updated for MVAs to output better data
    #>
 
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the filter to use for the groups.")]
        [string]$Filter = "*",

        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the Groups datatable to update with found information")]
        [System.Data.DataTable]$Groups,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the array of group attributes to return with the DataTable.")]
        [array]$GroupAttributes,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the EmailAddresses datatable to update with found information")]
        [System.Data.DataTable]$EmailAddresses,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute1 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute1,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute2 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute2,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute3 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute3,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute4 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute4,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute5 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute5,

        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ManagedBy datatable to update with found information")]
        [System.Data.DataTable]$ManagedBy,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the AcceptMessagesOnlyFrom datatable to update with found information")]
        [System.Data.DataTable]$AcceptMessagesOnlyFrom,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the AcceptMessagesOnlyFromDLMembers datatable to update with found information")]
        [System.Data.DataTable]$AcceptMessagesOnlyFromDLMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the AcceptMessagesOnlyFromSendersOrMembers datatable to update with found information")]
        [System.Data.DataTable]$AcceptMessagesOnlyFromSendersOrMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the BypassModerationFromSendersOrMembers datatable to update with found information")]
        [System.Data.DataTable]$BypassModerationFromSendersOrMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the GrantSendOnBehalfTo datatable to update with found information")]
        [System.Data.DataTable]$GrantSendOnBehalfTo,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ModeratedBy datatable to update with found information")]
        [System.Data.DataTable]$ModeratedBy,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the RejectMessagesFrom datatable to update with found information")]
        [System.Data.DataTable]$RejectMessagesFrom,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the RejectMessagesFromDLMembers datatable to update with found information")]
        [System.Data.DataTable]$RejectMessagesFromDLMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the RejectMessagesFromSendersOrMembers datatable to update with found information")]
        [System.Data.DataTable]$RejectMessagesFromSendersOrMembers,
                
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$Recipients
    )
    
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetDistributionGroupInfo"
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Building base datatable"
            $dtDLGroups = Get-DistributionGroup -ResultSize 1 -WarningAction SilentlyContinue -Verbose:$false -ErrorAction Stop | Select-Object -Property $GroupAttributes | ConvertTo-DataTable
            
            foreach ($column in $dtDLGroups.Columns) { if (-not $Groups.Columns.Contains($column.ColumnName)) { [void]$Groups.Columns.Add($column.ColumnName, $column.DataType) } }
            if (-not $Groups.Columns.Contains("GroupDomain")) { [void]$Groups.Columns.Add("GroupDomain", "string") }
            if ($Groups.Rows.Count -le 0) { $Groups.columns["Description"].Datatype = "System.String" }

            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with group information"
            Get-DistributionGroup -ResultSize Unlimited -Verbose:$false -Filter $Filter -ErrorAction Stop | Select-Object -Property $GroupAttributes | ForEach-Object {
                $drNewRow = $Groups.NewRow()
                $group = $_

                ForEach ($element in $_.PSObject.Properties) {
                    $columnName = $element.Name
                    $columnValue = $element.Value
                    
                    if ([string]::IsNullorEmpty($columnValue) -or $columnValue.ToString() -eq "Unlimited") {
                        $columnValue = [DBNull]::Value
                    } else {
                        switch ($columnName) {
                            "EmailAddresses" {
                                ForEach ($entry in $columnValue) {
                                    $drNewAddressRow = $EmailAddresses.NewRow()
                                    $drNewAddressRow["GroupDomain"] = [string]$GroupDomain
                                    $drNewAddressRow["GroupGuid"] = $group.Guid
                                    $drNewAddressRow["EmailAddresses"] = [string]$entry
                                    [void]$EmailAddresses.Rows.Add($drNewAddressRow)
                                }#foreach
                            }#EmailAddresses
                            "ExtensionCustomAttribute1" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute1Row = $ExtensionCustomAttribute1.NewRow()
                                    $drExtCustomAttribute1Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute1Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute1Row["ExtensionCustomAttribute1"] = [string]$entry
                                    [void]$ExtensionCustomAttribute1.Rows.Add($drExtCustomAttribute1Row)
                                }#foreach
                            }#ExtensionCustomAttribute1
                            "ExtensionCustomAttribute2" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute2Row = $ExtensionCustomAttribute2.NewRow()
                                    $drExtCustomAttribute2Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute2Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute2Row["ExtensionCustomAttribute2"] = [string]$entry
                                    [void]$ExtensionCustomAttribute2.Rows.Add($drExtCustomAttribute2Row)
                                }#foreach
                            }#ExtensionCustomAttribute2
                            "ExtensionCustomAttribute3" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute3Row = $ExtensionCustomAttribute3.NewRow()
                                    $drExtCustomAttribute3Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute3Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute3Row["ExtensionCustomAttribute3"] = [string]$entry
                                    [void]$ExtensionCustomAttribute3.Rows.Add($drExtCustomAttribute3Row)
                                }#foreach
                            }#ExtensionCustomAttribute3
                            "ExtensionCustomAttribute4" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute4Row = $ExtensionCustomAttribute4.NewRow()
                                    $drExtCustomAttribute4Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute4Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute4Row["ExtensionCustomAttribute4"] = [string]$entry
                                    [void]$ExtensionCustomAttribute4.Rows.Add($drExtCustomAttribute4Row)
                                }#foreach
                            }#ExtensionCustomAttribute4
                            "ExtensionCustomAttribute5" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute5Row = $ExtensionCustomAttribute5.NewRow()
                                    $drExtCustomAttribute5Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute5Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute5Row["ExtensionCustomAttribute5"] = [string]$entry
                                    [void]$ExtensionCustomAttribute5.Rows.Add($drExtCustomAttribute5Row)
                                }#foreach
                            }#ExtensionCustomAttribute5
                            "ManagedBy" {
                                ForEach ($entry in $columnValue) {
                                    $ManagedByUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $ManagedByUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]
                                    
                                    if ($ManagedByUser) {
                                        $drManagedByRow = $ManagedBy.NewRow()
                                        $drManagedByRow["ManagedByRecipientTypeDetails"] = $ManagedByUser.RecipientTypeDetails
                                        $drManagedByRow["ManagedByPrimarySmtpAddress"] = $ManagedByUser.PrimarySmtpAddress
                                        $drManagedByRow["ManagedByAlias"] = $ManagedByUser.alias
                                        $drManagedByRow["ManagedByDisplayName"] = $ManagedByUser.DisplayName
                                        $drManagedByRow["ManagedByGuid"] = $ManagedByUser.Guid
                                        $drManagedByRow["ManagedByName"] = $ManagedByUser.Name
                                        $drManagedByRow["GroupGuid"] = $group.Guid
                                        $drManagedByRow["GroupAlias"] = $group.alias
                                        $drManagedByRow["GroupDisplayName"] = $group.DisplayName
                                        $drManagedByRow["GroupName"] = $group.Name
                                        $drManagedByRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drManagedByRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drManagedByRow["GroupDomain"] = $GroupDomain
                                        [void]$ManagedBy.Rows.Add($drManagedByRow)
                                    }#if
                                }#foreach
                            }#ManagedBy
                            "AcceptMessagesOnlyFrom" {
                                ForEach ($entry in $columnValue) {
                                    $AcceptFrom = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $AcceptFrom = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($AcceptFrom) {
                                        $drAcceptMessagesOnlyFromRow = $AcceptMessagesOnlyFrom.NewRow()
                                        $drAcceptMessagesOnlyFromRow["AcceptMessagesOnlyFromRecipientTypeDetails"] = $AcceptFrom.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromRow["AcceptMessagesOnlyFromPrimarySmtpAddress"] = $AcceptFrom.PrimarySmtpAddress
                                        $drAcceptMessagesOnlyFromRow["AcceptMessagesOnlyFromAlias"] = $AcceptFrom.alias
                                        $drAcceptMessagesOnlyFromRow["AcceptMessagesOnlyFromDisplayName"] = $AcceptFrom.DisplayName
                                        $drAcceptMessagesOnlyFromRow["AcceptMessagesOnlyFromGuid"] = $AcceptFrom.Guid
                                        $drAcceptMessagesOnlyFromRow["AcceptMessagesOnlyFromName"] = $AcceptFrom.Name
                                        $drAcceptMessagesOnlyFromRow["GroupGuid"] = $group.Guid
                                        $drAcceptMessagesOnlyFromRow["GroupAlias"] = $group.alias
                                        $drAcceptMessagesOnlyFromRow["GroupDisplayName"] = $group.DisplayName
                                        $drAcceptMessagesOnlyFromRow["GroupName"] = $group.Name
                                        $drAcceptMessagesOnlyFromRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drAcceptMessagesOnlyFromRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromRow["GroupDomain"] = $GroupDomain
                                        [void]$AcceptMessagesOnlyFrom.Rows.Add($drAcceptMessagesOnlyFromRow)
                                    }
                                }#foreach
                            }#AcceptMessagesOnlyFrom
                            "AcceptMessagesOnlyFromDLMembers" {
                                ForEach ($entry in $columnValue) {
                                    $AcceptFromDLMember = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $AcceptFromDLMember = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($AcceptFromDLMember) {
                                        $drAcceptMessagesOnlyFromDLMembersRow = $AcceptMessagesOnlyFromDLMembers.NewRow()
                                        $drAcceptMessagesOnlyFromDLMembersRow["AcceptMessagesOnlyFromDLMembersRecipientTypeDetails"] = $AcceptFromDLMember.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromDLMembersRow["AcceptMessagesOnlyFromDLMembersPrimarySmtpAddress"] = $AcceptFromDLMember.PrimarySmtpAddress
                                        $drAcceptMessagesOnlyFromDLMembersRow["AcceptMessagesOnlyFromDLMembersAlias"] = $AcceptFromDLMember.alias
                                        $drAcceptMessagesOnlyFromDLMembersRow["AcceptMessagesOnlyFromDLMembersDisplayName"] = $AcceptFromDLMember.DisplayName
                                        $drAcceptMessagesOnlyFromDLMembersRow["AcceptMessagesOnlyFromDLMembersGuid"] = $AcceptFromDLMember.Guid
                                        $drAcceptMessagesOnlyFromDLMembersRow["AcceptMessagesOnlyFromDLMembersName"] = $AcceptFromDLMember.Name
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupGuid"] = $group.Guid
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupAlias"] = $group.alias
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupName"] = $group.Name
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$AcceptMessagesOnlyFromDLMembers.Rows.Add($drAcceptMessagesOnlyFromDLMembersRow)
                                    }
                                }#foreach
                            }#AcceptMessagesOnlyFromDLMembers
                            "AcceptMessagesOnlyFromSendersOrMembers" {
                                ForEach ($entry in $columnValue) {
                                    $AcceptFromSender = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $AcceptFromSender = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($AcceptFromSender) {
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow = $AcceptMessagesOnlyFromSendersOrMembers.NewRow()
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["AcceptMessagesOnlyFromSendersOrMembersRecipientTypeDetails"] = $AcceptFromSender.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["AcceptMessagesOnlyFromSendersOrMembersPrimarySmtpAddress"] = $AcceptFromSender.PrimarySmtpAddress
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["AcceptMessagesOnlyFromSendersOrMembersAlias"] = $AcceptFromSender.alias
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["AcceptMessagesOnlyFromSendersOrMembersDisplayName"] = $AcceptFromSender.DisplayName
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["AcceptMessagesOnlyFromSendersOrMembersGuid"] = $AcceptFromSender.Guid
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["AcceptMessagesOnlyFromSendersOrMembersName"] = $AcceptFromSender.Name
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupGuid"] = $group.Guid
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupAlias"] = $group.alias
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupName"] = $group.Name
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$AcceptMessagesOnlyFromSendersOrMembers.Rows.Add($drAcceptMessagesOnlyFromSendersOrMembersRow)
                                    }
                                }#foreach
                            }#AcceptMessagesOnlyFromSendersOrMembers
                            "BypassModerationFromSendersOrMembers" {
                                ForEach ($entry in $columnValue) {
                                    $BypassModerationUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $BypassModerationUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($BypassModerationUser) {
                                        $drBypassModerationFromSendersOrMembersRow = $BypassModerationFromSendersOrMembers.NewRow()
                                        $drBypassModerationFromSendersOrMembersRow["BypassModerationFromSendersOrMembersRecipientTypeDetails"] = $BypassModerationUser.RecipientTypeDetails
                                        $drBypassModerationFromSendersOrMembersRow["BypassModerationFromSendersOrMembersPrimarySmtpAddress"] = $BypassModerationUser.PrimarySmtpAddress
                                        $drBypassModerationFromSendersOrMembersRow["BypassModerationFromSendersOrMembersAlias"] = $BypassModerationUser.alias
                                        $drBypassModerationFromSendersOrMembersRow["BypassModerationFromSendersOrMembersDisplayName"] = $BypassModerationUser.DisplayName
                                        $drBypassModerationFromSendersOrMembersRow["BypassModerationFromSendersOrMembersGuid"] = $BypassModerationUser.Guid
                                        $drBypassModerationFromSendersOrMembersRow["BypassModerationFromSendersOrMembersName"] = $BypassModerationUser.Name
                                        $drBypassModerationFromSendersOrMembersRow["GroupGuid"] = $group.Guid
                                        $drBypassModerationFromSendersOrMembersRow["GroupAlias"] = $group.alias
                                        $drBypassModerationFromSendersOrMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drBypassModerationFromSendersOrMembersRow["GroupName"] = $group.Name
                                        $drBypassModerationFromSendersOrMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drBypassModerationFromSendersOrMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drBypassModerationFromSendersOrMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$BypassModerationFromSendersOrMembers.Rows.Add($drBypassModerationFromSendersOrMembersRow)
                                    }
                                }#foreach
                            }#BypassModerationFromSendersOrMembers
                            "GrantSendOnBehalfTo" {
                                ForEach ($entry in $columnValue) {
                                    $SendOnBehalfUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $SendOnBehalfUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($SendOnBehalfUser) {
                                        $drGrantSendOnBehalfToRow = $GrantSendOnBehalfTo.NewRow()
                                        $drGrantSendOnBehalfToRow["GrantSendOnBehalfToRecipientTypeDetails"] = $SendOnBehalfUser.RecipientTypeDetails
                                        $drGrantSendOnBehalfToRow["GrantSendOnBehalfToPrimarySmtpAddress"] = $SendOnBehalfUser.PrimarySmtpAddress
                                        $drGrantSendOnBehalfToRow["GrantSendOnBehalfToAlias"] = $SendOnBehalfUser.alias
                                        $drGrantSendOnBehalfToRow["GrantSendOnBehalfToDisplayName"] = $SendOnBehalfUser.DisplayName
                                        $drGrantSendOnBehalfToRow["GrantSendOnBehalfToGuid"] = $SendOnBehalfUser.Guid
                                        $drGrantSendOnBehalfToRow["GrantSendOnBehalfToName"] = $SendOnBehalfUser.Name
                                        $drGrantSendOnBehalfToRow["GroupGuid"] = $group.Guid
                                        $drGrantSendOnBehalfToRow["GroupAlias"] = $group.alias
                                        $drGrantSendOnBehalfToRow["GroupDisplayName"] = $group.DisplayName
                                        $drGrantSendOnBehalfToRow["GroupName"] = $group.Name
                                        $drGrantSendOnBehalfToRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drGrantSendOnBehalfToRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drGrantSendOnBehalfToRow["GroupDomain"] = $GroupDomain
                                        [void]$GrantSendOnBehalfTo.Rows.Add($drGrantSendOnBehalfToRow)
                                    }
                                }#foreach
                            }#GrantSendOnBehalfTo
                            "ModeratedBy" {
                                ForEach ($entry in $columnValue) {
                                    $ModeratedByUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $ModeratedByUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($ModeratedByUser) {
                                        $drModeratedByRow = $ModeratedBy.NewRow()
                                        $drModeratedByRow["ModeratedByRecipientTypeDetails"] = $ModeratedByUser.RecipientTypeDetails
                                        $drModeratedByRow["ModeratedByPrimarySmtpAddress"] = $ModeratedByUser.PrimarySmtpAddress
                                        $drModeratedByRow["ModeratedByAlias"] = $ModeratedByUser.alias
                                        $drModeratedByRow["ModeratedByDisplayName"] = $ModeratedByUser.DisplayName
                                        $drModeratedByRow["ModeratedByGuid"] = $ModeratedByUser.Guid
                                        $drModeratedByRow["ModeratedByName"] = $ModeratedByUser.Name
                                        $drModeratedByRow["GroupGuid"] = $group.Guid
                                        $drModeratedByRow["GroupAlias"] = $group.alias
                                        $drModeratedByRow["GroupDisplayName"] = $group.DisplayName
                                        $drModeratedByRow["GroupName"] = $group.Name
                                        $drModeratedByRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drModeratedByRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drModeratedByRow["GroupDomain"] = $GroupDomain
                                        [void]$ModeratedBy.Rows.Add($drModeratedByRow)
                                    }
                                }#foreach
                            }#ModeratedBy
                            "RejectMessagesFrom" {
                                ForEach ($entry in $columnValue) {
                                    $RejectFromUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $RejectFromUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($RejectFromUser) {
                                        $drRejectMessagesFromRow = $RejectMessagesFrom.NewRow()
                                        $drRejectMessagesFromRow["ManagedByRecipientTypeDetails"] = $RejectFromUser.RecipientTypeDetails
                                        $drRejectMessagesFromRow["ManagedByPrimarySmtpAddress"] = $RejectFromUser.PrimarySmtpAddress
                                        $drRejectMessagesFromRow["ManagedByAlias"] = $RejectFromUser.alias
                                        $drRejectMessagesFromRow["ManagedByDisplayName"] = $RejectFromUser.DisplayName
                                        $drRejectMessagesFromRow["ManagedByGuid"] = $RejectFromUser.Guid
                                        $drRejectMessagesFromRow["ManagedByName"] = $RejectFromUser.Name
                                        $drRejectMessagesFromRow["GroupGuid"] = $group.Guid
                                        $drRejectMessagesFromRow["GroupAlias"] = $group.alias
                                        $drRejectMessagesFromRow["GroupDisplayName"] = $group.DisplayName
                                        $drRejectMessagesFromRow["GroupName"] = $group.Name
                                        $drRejectMessagesFromRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drRejectMessagesFromRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drRejectMessagesFromRow["GroupDomain"] = $GroupDomain
                                        [void]$RejectMessagesFrom.Rows.Add($drRejectMessagesFromRow)
                                    }
                                }#foreach
                            }#RejectMessagesFrom
                            "RejectMessagesFromDLMembers" {
                                ForEach ($entry in $columnValue) {
                                    $RejectFromDL = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $RejectFromDL = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($RejectFromDL) {
                                        $drRejectMessagesFromDLMembersRow = $RejectMessagesFromDLMembers.NewRow()
                                        $drRejectMessagesFromDLMembersRow["RejectMessagesFromDLMembersRecipientTypeDetails"] = $RejectFromDL.RecipientTypeDetails
                                        $drRejectMessagesFromDLMembersRow["RejectMessagesFromDLMembersPrimarySmtpAddress"] = $RejectFromDL.PrimarySmtpAddress
                                        $drRejectMessagesFromDLMembersRow["RejectMessagesFromDLMembersAlias"] = $RejectFromDL.alias
                                        $drRejectMessagesFromDLMembersRow["RejectMessagesFromDLMembersDisplayName"] = $RejectFromDL.DisplayName
                                        $drRejectMessagesFromDLMembersRow["RejectMessagesFromDLMembersGuid"] = $RejectFromDL.Guid
                                        $drRejectMessagesFromDLMembersRow["RejectMessagesFromDLMembersName"] = $RejectFromDL.Name
                                        $drRejectMessagesFromDLMembersRow["GroupGuid"] = $group.Guid
                                        $drRejectMessagesFromDLMembersRow["GroupAlias"] = $group.alias
                                        $drRejectMessagesFromDLMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drRejectMessagesFromDLMembersRow["GroupName"] = $group.Name
                                        $drRejectMessagesFromDLMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drRejectMessagesFromDLMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drRejectMessagesFromDLMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$RejectMessagesFromDLMembers.Rows.Add($drRejectMessagesFromDLMembersRow)
                                    }
                                }#foreach
                            }#RejectMessagesFromDLMembers
                            "RejectMessagesFromSendersOrMembers" {
                                ForEach ($entry in $columnValue) {
                                    $RejectFromSender = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $RejectFromSender = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($RejectFromSender) {
                                        $drRejectMessagesFromSendersOrMembersRow = $RejectMessagesFromSendersOrMembers.NewRow()
                                        $drRejectMessagesFromSendersOrMembersRow["RejectMessagesFromSendersOrMembersRecipientTypeDetails"] = $RejectFromSender.RecipientTypeDetails
                                        $drRejectMessagesFromSendersOrMembersRow["RejectMessagesFromSendersOrMembersPrimarySmtpAddress"] = $RejectFromSender.PrimarySmtpAddress
                                        $drRejectMessagesFromSendersOrMembersRow["RejectMessagesFromSendersOrMembersAlias"] = $RejectFromSender.alias
                                        $drRejectMessagesFromSendersOrMembersRow["RejectMessagesFromSendersOrMembersDisplayName"] = $RejectFromSender.DisplayName
                                        $drRejectMessagesFromSendersOrMembersRow["RejectMessagesFromSendersOrMembersGuid"] = $RejectFromSender.Guid
                                        $drRejectMessagesFromSendersOrMembersRow["RejectMessagesFromSendersOrMembersName"] = $RejectFromSender.Name
                                        $drRejectMessagesFromSendersOrMembersRow["GroupGuid"] = $group.Guid
                                        $drRejectMessagesFromSendersOrMembersRow["GroupAlias"] = $group.alias
                                        $drRejectMessagesFromSendersOrMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drRejectMessagesFromSendersOrMembersRow["GroupName"] = $group.Name
                                        $drRejectMessagesFromSendersOrMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drRejectMessagesFromSendersOrMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drRejectMessagesFromSendersOrMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$RejectMessagesFromSendersOrMembers.Rows.Add($drRejectMessagesFromSendersOrMembersRow)
                                    }
                                }#foreach
                            }#RejectMessagesFromSendersOrMembers
                            "MaxSendSize" { $drNewRow["MaxSendSize"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "MaxReceiveSize" { $drNewRow["MaxReceiveSize"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            default {
                                if ($columnValue.gettype().Name -eq "ArrayList") {
                                    $drNewRow["$columnName"] = [string]$columnValue.Clone()
                                } else {
                                    $drNewRow["$columnName"] = $columnValue
                                }#if/else
                            }#default
                        }#switch
                    }#if/else
                }#loop through each property
                $drNewRow["GroupDomain"] = [string]$GroupDomain

                [void]$Groups.Rows.Add($drNewRow)
            }#invoke/foreach
            
            return @(, $Groups)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather group information"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
    
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetDistributionGroupInfo"
    }#end
}#function _GetDistributionGroupInfo

Function _GetDynamicGroupInfo {
    <#
    .SYNOPSIS
        Collects the necessary dynamic distribution group cache and returns a datatable with the results
    .PARAMETER Session
        The PSSession to run the command against
    .PARAMETER Filter
        The optional parameter for a filter to be used when querying groups
   .PARAMETER GroupAttributes
        Specify the array of group attributes to return with the DataTable.
    .EXAMPLE
        $dtGroups = _GetDynamicGroupInfo -Session $session GroupAttributes $GroupAttribs
    
        This would get all groups from the PowerShell $session and return attributes $GroupAttribs to $dtGroups
    .NOTES
        Version:
            - 5.0.2023.0726:    New function
            - 5.1.2024.0110:    Updated for MVAs to output better data
    #>
 
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the filter to use for the groups.")]
        [string]$Filter = "*",
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the Groups datatable to update with found information")]
        [System.Data.DataTable]$Groups,

        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the array of group attributes to return with the DataTable.")]
        [array]$GroupAttributes,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the EmailAddresses datatable to update with found information")]
        [System.Data.DataTable]$EmailAddresses,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute1 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute1,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute2 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute2,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute3 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute3,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute4 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute4,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute5 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute5,

        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ManagedBy datatable to update with found information")]
        [System.Data.DataTable]$ManagedBy,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the AcceptMessagesOnlyFrom datatable to update with found information")]
        [System.Data.DataTable]$AcceptMessagesOnlyFrom,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the AcceptMessagesOnlyFromDLMembers datatable to update with found information")]
        [System.Data.DataTable]$AcceptMessagesOnlyFromDLMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the AcceptMessagesOnlyFromSendersOrMembers datatable to update with found information")]
        [System.Data.DataTable]$AcceptMessagesOnlyFromSendersOrMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the BypassModerationFromSendersOrMembers datatable to update with found information")]
        [System.Data.DataTable]$BypassModerationFromSendersOrMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the GrantSendOnBehalfTo datatable to update with found information")]
        [System.Data.DataTable]$GrantSendOnBehalfTo,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ModeratedBy datatable to update with found information")]
        [System.Data.DataTable]$ModeratedBy,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the RejectMessagesFrom datatable to update with found information")]
        [System.Data.DataTable]$RejectMessagesFrom,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the RejectMessagesFromDLMembers datatable to update with found information")]
        [System.Data.DataTable]$RejectMessagesFromDLMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the RejectMessagesFromSendersOrMembers datatable to update with found information")]
        [System.Data.DataTable]$RejectMessagesFromSendersOrMembers,

        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$Recipients
    )
    
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetDynamicGroupInfo"
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Building base datatable"
            $dtDDLGroups = Get-DynamicDistributionGroup -ResultSize 1 -WarningAction SilentlyContinue -Verbose:$false -ErrorAction Stop | Select-Object -Property $GroupAttributes | ConvertTo-DataTable
            
            foreach ($column in $dtDDLGroups.Columns) { if (-not $Groups.Columns.Contains($column.ColumnName)) { $Groups.Columns.Add($column.ColumnName, $column.DataType) } }
            if (-not $Groups.Columns.Contains("GroupDomain")) { [void]$Groups.Columns.Add("GroupDomain", "string") }

            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with dynamic group information"
            
            Get-DynamicDistributionGroup -ResultSize Unlimited -Verbose:$false -Filter $Filter -ErrorAction Stop | Select-Object -Property $GroupAttributes | ForEach-Object {
                $drNewRow = $Groups.NewRow()
                $group = $_

                ForEach ($element in $_.PSObject.Properties) {
                    $columnName = $element.Name
                    $columnValue = $element.Value
                    
                    if ([string]::IsNullorEmpty($columnValue) -or $columnValue.ToString() -eq "Unlimited") {
                        $columnValue = [DBNull]::Value
                    } else {
                        switch ($columnName) {
                            "EmailAddresses" {
                                ForEach ($entry in $columnValue) {
                                    $drNewAddressRow = $EmailAddresses.NewRow()
                                    $drNewAddressRow["GroupDomain"] = [string]$GroupDomain
                                    $drNewAddressRow["GroupGuid"] = $group.Guid
                                    $drNewAddressRow["EmailAddresses"] = [string]$entry
                                    [void]$EmailAddresses.Rows.Add($drNewAddressRow)
                                }#foreach
                            }#EmailAddresses
                            "ExtensionCustomAttribute1" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute1Row = $ExtensionCustomAttribute1.NewRow()
                                    $drExtCustomAttribute1Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute1Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute1Row["ExtensionCustomAttribute1"] = [string]$entry
                                    [void]$ExtensionCustomAttribute1.Rows.Add($drExtCustomAttribute1Row)
                                }#foreach
                            }#ExtensionCustomAttribute1
                            "ExtensionCustomAttribute2" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute2Row = $ExtensionCustomAttribute2.NewRow()
                                    $drExtCustomAttribute2Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute2Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute2Row["ExtensionCustomAttribute2"] = [string]$entry
                                    [void]$ExtensionCustomAttribute2.Rows.Add($drExtCustomAttribute2Row)
                                }#foreach
                            }#ExtensionCustomAttribute2
                            "ExtensionCustomAttribute3" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute3Row = $ExtensionCustomAttribute3.NewRow()
                                    $drExtCustomAttribute3Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute3Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute3Row["ExtensionCustomAttribute3"] = [string]$entry
                                    [void]$ExtensionCustomAttribute3.Rows.Add($drExtCustomAttribute3Row)
                                }#foreach
                            }#ExtensionCustomAttribute3
                            "ExtensionCustomAttribute4" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute4Row = $ExtensionCustomAttribute4.NewRow()
                                    $drExtCustomAttribute4Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute4Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute4Row["ExtensionCustomAttribute4"] = [string]$entry
                                    [void]$ExtensionCustomAttribute4.Rows.Add($drExtCustomAttribute4Row)
                                }#foreach
                            }#ExtensionCustomAttribute4
                            "ExtensionCustomAttribute5" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute5Row = $ExtensionCustomAttribute5.NewRow()
                                    $drExtCustomAttribute5Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute5Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute5Row["ExtensionCustomAttribute5"] = [string]$entry
                                    [void]$ExtensionCustomAttribute5.Rows.Add($drExtCustomAttribute5Row)
                                }#foreach
                            }#ExtensionCustomAttribute5
                            "ManagedBy" {
                                ForEach ($entry in $columnValue) {
                                    $ManagedByUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $ManagedByUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]
                                    
                                    if ($ManagedByUser) {
                                        $drManagedByRow = $ManagedBy.NewRow()
                                        $drManagedByRow["ManagedByRecipientTypeDetails"] = $ManagedByUser.RecipientTypeDetails
                                        $drManagedByRow["ManagedByPrimarySmtpAddress"] = $ManagedByUser.PrimarySmtpAddress
                                        $drManagedByRow["ManagedByAlias"] = $ManagedByUser.alias
                                        $drManagedByRow["ManagedByDisplayName"] = $ManagedByUser.DisplayName
                                        $drManagedByRow["ManagedByGuid"] = $ManagedByUser.Guid
                                        $drManagedByRow["ManagedByName"] = $ManagedByUser.Name
                                        $drManagedByRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drManagedByRow["GroupAlias"] = $group.alias
                                        $drManagedByRow["GroupDisplayName"] = $group.DisplayName
                                        $drManagedByRow["GroupName"] = $group.Name
                                        $drManagedByRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drManagedByRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drManagedByRow["GroupDomain"] = $GroupDomain
                                        [void]$ManagedBy.Rows.Add($drManagedByRow)
                                    }#if
                                }#foreach
                            }#ManagedBy
                            "AcceptMessagesOnlyFrom" {
                                ForEach ($entry in $columnValue) {
                                    $AcceptFrom = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $AcceptFrom = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($AcceptFrom) {
                                        $drAcceptMessagesOnlyFromRow = $AcceptMessagesOnlyFrom.NewRow()
                                        $drAcceptMessagesOnlyFromRow["ManagedByRecipientTypeDetails"] = $AcceptFrom.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromRow["ManagedByPrimarySmtpAddress"] = $AcceptFrom.PrimarySmtpAddress
                                        $drAcceptMessagesOnlyFromRow["ManagedByAlias"] = $AcceptFrom.alias
                                        $drAcceptMessagesOnlyFromRow["ManagedByDisplayName"] = $AcceptFrom.DisplayName
                                        $drAcceptMessagesOnlyFromRow["ManagedByGuid"] = $AcceptFrom.Guid
                                        $drAcceptMessagesOnlyFromRow["ManagedByName"] = $AcceptFrom.Name
                                        $drAcceptMessagesOnlyFromRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drAcceptMessagesOnlyFromRow["GroupAlias"] = $group.alias
                                        $drAcceptMessagesOnlyFromRow["GroupDisplayName"] = $group.DisplayName
                                        $drAcceptMessagesOnlyFromRow["GroupName"] = $group.Name
                                        $drAcceptMessagesOnlyFromRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drAcceptMessagesOnlyFromRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromRow["GroupDomain"] = $GroupDomain
                                        [void]$AcceptMessagesOnlyFrom.Rows.Add($drAcceptMessagesOnlyFromRow)
                                    }
                                }#foreach
                            }#AcceptMessagesOnlyFrom
                            "AcceptMessagesOnlyFromDLMembers" {
                                ForEach ($entry in $columnValue) {
                                    $AcceptFromDLMember = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $AcceptFromDLMember = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($AcceptFromDLMember) {
                                        $drAcceptMessagesOnlyFromDLMembersRow = $AcceptMessagesOnlyFromDLMembers.NewRow()
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByRecipientTypeDetails"] = $AcceptFromDLMember.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByPrimarySmtpAddress"] = $AcceptFromDLMember.PrimarySmtpAddress
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByAlias"] = $AcceptFromDLMember.alias
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByDisplayName"] = $AcceptFromDLMember.DisplayName
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByGuid"] = $AcceptFromDLMember.Guid
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByName"] = $AcceptFromDLMember.Name
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupAlias"] = $group.alias
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupName"] = $group.Name
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$AcceptMessagesOnlyFromDLMembers.Rows.Add($drAcceptMessagesOnlyFromDLMembersRow)
                                    }
                                }#foreach
                            }#AcceptMessagesOnlyFromDLMembers
                            "AcceptMessagesOnlyFromSendersOrMembers" {
                                ForEach ($entry in $columnValue) {
                                    $AcceptFromSender = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $AcceptFromSender = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($AcceptFromSender) {
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow = $AcceptMessagesOnlyFromSendersOrMembers.NewRow()
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByRecipientTypeDetails"] = $AcceptFromSender.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByPrimarySmtpAddress"] = $AcceptFromSender.PrimarySmtpAddress
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByAlias"] = $AcceptFromSender.alias
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByDisplayName"] = $AcceptFromSender.DisplayName
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByGuid"] = $AcceptFromSender.Guid
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByName"] = $AcceptFromSender.Name
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupAlias"] = $group.alias
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupName"] = $group.Name
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$AcceptMessagesOnlyFromSendersOrMembers.Rows.Add($drAcceptMessagesOnlyFromSendersOrMembersRow)
                                    }
                                }#foreach
                            }#AcceptMessagesOnlyFromSendersOrMembers
                            "BypassModerationFromSendersOrMembers" {
                                ForEach ($entry in $columnValue) {
                                    $BypassModerationUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $BypassModerationUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($BypassModerationUser) {
                                        $drBypassModerationFromSendersOrMembersRow = $BypassModerationFromSendersOrMembers.NewRow()
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByRecipientTypeDetails"] = $BypassModerationUser.RecipientTypeDetails
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByPrimarySmtpAddress"] = $BypassModerationUser.PrimarySmtpAddress
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByAlias"] = $BypassModerationUser.alias
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByDisplayName"] = $BypassModerationUser.DisplayName
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByGuid"] = $BypassModerationUser.Guid
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByName"] = $BypassModerationUser.Name
                                        $drBypassModerationFromSendersOrMembersRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drBypassModerationFromSendersOrMembersRow["GroupAlias"] = $group.alias
                                        $drBypassModerationFromSendersOrMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drBypassModerationFromSendersOrMembersRow["GroupName"] = $group.Name
                                        $drBypassModerationFromSendersOrMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drBypassModerationFromSendersOrMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drBypassModerationFromSendersOrMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$BypassModerationFromSendersOrMembers.Rows.Add($drBypassModerationFromSendersOrMembersRow)
                                    }
                                }#foreach
                            }#BypassModerationFromSendersOrMembers
                            "GrantSendOnBehalfTo" {
                                ForEach ($entry in $columnValue) {
                                    $SendOnBehalfUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $SendOnBehalfUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($SendOnBehalfUser) {
                                        $drGrantSendOnBehalfToRow = $GrantSendOnBehalfTo.NewRow()
                                        $drGrantSendOnBehalfToRow["ManagedByRecipientTypeDetails"] = $SendOnBehalfUser.RecipientTypeDetails
                                        $drGrantSendOnBehalfToRow["ManagedByPrimarySmtpAddress"] = $SendOnBehalfUser.PrimarySmtpAddress
                                        $drGrantSendOnBehalfToRow["ManagedByAlias"] = $SendOnBehalfUser.alias
                                        $drGrantSendOnBehalfToRow["ManagedByDisplayName"] = $SendOnBehalfUser.DisplayName
                                        $drGrantSendOnBehalfToRow["ManagedByGuid"] = $SendOnBehalfUser.Guid
                                        $drGrantSendOnBehalfToRow["ManagedByName"] = $SendOnBehalfUser.Name
                                        $drGrantSendOnBehalfToRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drGrantSendOnBehalfToRow["GroupAlias"] = $group.alias
                                        $drGrantSendOnBehalfToRow["GroupDisplayName"] = $group.DisplayName
                                        $drGrantSendOnBehalfToRow["GroupName"] = $group.Name
                                        $drGrantSendOnBehalfToRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drGrantSendOnBehalfToRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drGrantSendOnBehalfToRow["GroupDomain"] = $GroupDomain
                                        [void]$GrantSendOnBehalfTo.Rows.Add($drGrantSendOnBehalfToRow)
                                    }
                                }#foreach
                            }#GrantSendOnBehalfTo
                            "ModeratedBy" {
                                ForEach ($entry in $columnValue) {
                                    $ModeratedByUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $ModeratedByUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($ModeratedByUser) {
                                        $drModeratedByRow = $ModeratedBy.NewRow()
                                        $drModeratedByRow["ManagedByRecipientTypeDetails"] = $ModeratedByUser.RecipientTypeDetails
                                        $drModeratedByRow["ManagedByPrimarySmtpAddress"] = $ModeratedByUser.PrimarySmtpAddress
                                        $drModeratedByRow["ManagedByAlias"] = $ModeratedByUser.alias
                                        $drModeratedByRow["ManagedByDisplayName"] = $ModeratedByUser.DisplayName
                                        $drModeratedByRow["ManagedByGuid"] = $ModeratedByUser.Guid
                                        $drModeratedByRow["ManagedByName"] = $ModeratedByUser.Name
                                        $drModeratedByRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drModeratedByRow["GroupAlias"] = $group.alias
                                        $drModeratedByRow["GroupDisplayName"] = $group.DisplayName
                                        $drModeratedByRow["GroupName"] = $group.Name
                                        $drModeratedByRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drModeratedByRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drModeratedByRow["GroupDomain"] = $GroupDomain
                                        [void]$ModeratedBy.Rows.Add($drModeratedByRow)
                                    }
                                }#foreach
                            }#ModeratedBy
                            "RejectMessagesFrom" {
                                ForEach ($entry in $columnValue) {
                                    $RejectFromUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $RejectFromUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($RejectFromUser) {
                                        $drRejectMessagesFromRow = $RejectMessagesFrom.NewRow()
                                        $drRejectMessagesFromRow["ManagedByRecipientTypeDetails"] = $RejectFromUser.RecipientTypeDetails
                                        $drRejectMessagesFromRow["ManagedByPrimarySmtpAddress"] = $RejectFromUser.PrimarySmtpAddress
                                        $drRejectMessagesFromRow["ManagedByAlias"] = $RejectFromUser.alias
                                        $drRejectMessagesFromRow["ManagedByDisplayName"] = $RejectFromUser.DisplayName
                                        $drRejectMessagesFromRow["ManagedByGuid"] = $RejectFromUser.Guid
                                        $drRejectMessagesFromRow["ManagedByName"] = $RejectFromUser.Name
                                        $drRejectMessagesFromRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drRejectMessagesFromRow["GroupAlias"] = $group.alias
                                        $drRejectMessagesFromRow["GroupDisplayName"] = $group.DisplayName
                                        $drRejectMessagesFromRow["GroupName"] = $group.Name
                                        $drRejectMessagesFromRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drRejectMessagesFromRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drRejectMessagesFromRow["GroupDomain"] = $GroupDomain
                                        [void]$RejectMessagesFrom.Rows.Add($drRejectMessagesFromRow)
                                    }
                                }#foreach
                            }#RejectMessagesFrom
                            "RejectMessagesFromDLMembers" {
                                ForEach ($entry in $columnValue) {
                                    $RejectFromDL = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $RejectFromDL = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($RejectFromDL) {
                                        $drRejectMessagesFromDLMembersRow = $RejectMessagesFromDLMembers.NewRow()
                                        $drRejectMessagesFromDLMembersRow["ManagedByRecipientTypeDetails"] = $RejectFromDL.RecipientTypeDetails
                                        $drRejectMessagesFromDLMembersRow["ManagedByPrimarySmtpAddress"] = $RejectFromDL.PrimarySmtpAddress
                                        $drRejectMessagesFromDLMembersRow["ManagedByAlias"] = $RejectFromDL.alias
                                        $drRejectMessagesFromDLMembersRow["ManagedByDisplayName"] = $RejectFromDL.DisplayName
                                        $drRejectMessagesFromDLMembersRow["ManagedByGuid"] = $RejectFromDL.Guid
                                        $drRejectMessagesFromDLMembersRow["ManagedByName"] = $RejectFromDL.Name
                                        $drRejectMessagesFromDLMembersRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drRejectMessagesFromDLMembersRow["GroupAlias"] = $group.alias
                                        $drRejectMessagesFromDLMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drRejectMessagesFromDLMembersRow["GroupName"] = $group.Name
                                        $drRejectMessagesFromDLMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drRejectMessagesFromDLMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drRejectMessagesFromDLMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$RejectMessagesFromDLMembers.Rows.Add($drRejectMessagesFromDLMembersRow)
                                    }
                                }#foreach
                            }#RejectMessagesFromDLMembers
                            "RejectMessagesFromSendersOrMembers" {
                                ForEach ($entry in $columnValue) {
                                    $RejectFromSender = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $RejectFromSender = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($RejectFromSender) {
                                        $drRejectMessagesFromSendersOrMembersRow = $RejectMessagesFromSendersOrMembers.NewRow()
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByRecipientTypeDetails"] = $RejectFromSender.RecipientTypeDetails
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByPrimarySmtpAddress"] = $RejectFromSender.PrimarySmtpAddress
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByAlias"] = $RejectFromSender.alias
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByDisplayName"] = $RejectFromSender.DisplayName
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByGuid"] = $RejectFromSender.Guid
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByName"] = $RejectFromSender.Name
                                        $drRejectMessagesFromSendersOrMembersRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drRejectMessagesFromSendersOrMembersRow["GroupAlias"] = $group.alias
                                        $drRejectMessagesFromSendersOrMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drRejectMessagesFromSendersOrMembersRow["GroupName"] = $group.Name
                                        $drRejectMessagesFromSendersOrMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drRejectMessagesFromSendersOrMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drRejectMessagesFromSendersOrMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$RejectMessagesFromSendersOrMembers.Rows.Add($drRejectMessagesFromSendersOrMembersRow)
                                    }
                                }#foreach
                            }#RejectMessagesFromSendersOrMembers
                            "MaxSendSize" { $drNewRow["MaxSendSize"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "MaxReceiveSize" { $drNewRow["MaxReceiveSize"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
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
                $drNewRow["GroupType"] = "Dynamic"
                $drNewRow["GroupDomain"] = [string]$GroupDomain

                [void]$Groups.Rows.Add($drNewRow)
            }#invoke/foreach
            
            return @(, $Groups)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather dynamic group information"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
    
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetDynamicGroupInfo"
    }#end
}#function _GetDynamicGroupInfo

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
            - 5.1.2024.0111:    New function
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
            $col1 = New-Object System.Data.DataColumn GroupDomain, ([string])
            $col2 = New-Object System.Data.DataColumn GroupGuid, ([Guid])
            $col3 = New-Object System.Data.DataColumn $Attribute, ([string])
            $dtMVA.Columns.Add($col1)
            $dtMVA.Columns.Add($col2)
            $dtMVA.Columns.Add($col3)
            [System.Data.DataColumn[]]$KeyColumn = ($dtMVA.Columns["GroupDomain"], $dtMVA.Columns["GroupGuid"], $dtMVA.Columns[$Attribute])
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

Function _CreateExpandedMVATable {
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
            - 5.1.2023.0726:    New function
            - 5.1.2024.0111:    Updated for better MVA output
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
            $col1 = New-Object System.Data.DataColumn GroupDomain, ([string])
            $col2 = New-Object System.Data.DataColumn GroupGuid, ([Guid])
            $col3 = New-Object System.Data.DataColumn GroupAlias, ([string])
            $col4 = New-Object System.Data.DataColumn GroupDisplayName, ([string])
            $col5 = New-Object System.Data.DataColumn GroupName, ([string])
            $col6 = New-Object System.Data.DataColumn GroupPrimarySmtpAddress, ([string])
            $col7 = New-Object System.Data.DataColumn GroupRecipientTypeDetails, ([string])
            $col8 = New-Object System.Data.DataColumn $Attribute"RecipientTypeDetails", ([string])
            $col9 = New-Object System.Data.DataColumn $Attribute"PrimarySmtpAddress", ([string])
            $col10 = New-Object System.Data.DataColumn $Attribute"Alias", ([string])
            $col11 = New-Object System.Data.DataColumn $Attribute"DisplayName", ([string])
            $col12 = New-Object System.Data.DataColumn $Attribute"Guid", ([string])
            $col13 = New-Object System.Data.DataColumn $Attribute"Name", ([string])
            $dtMVA.Columns.Add($col1)
            $dtMVA.Columns.Add($col2)
            $dtMVA.Columns.Add($col3)
            $dtMVA.Columns.Add($col4)
            $dtMVA.Columns.Add($col5)
            $dtMVA.Columns.Add($col6)
            $dtMVA.Columns.Add($col7)
            $dtMVA.Columns.Add($col8)
            $dtMVA.Columns.Add($col9)
            $dtMVA.Columns.Add($col10)
            $dtMVA.Columns.Add($col11)
            $dtMVA.Columns.Add($col12)
            $dtMVA.Columns.Add($col13)
            [System.Data.DataColumn[]]$KeyColumn = ($dtMVA.Columns["GroupDomain"], $dtMVA.Columns["GroupGuid"], $dtMVA.Columns[$Attribute + "Guid"])
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
}#function _CreateExpandedMVATable

Function _GetUnifiedGroupInfo {
    <#
    .SYNOPSIS
        Collects the necessary distribution group cache and returns a datatable with the results
    .PARAMETER Session
        The PSSession to run the command against
    .PARAMETER Filter
        The optional parameter for a filter to be used when querying groups
   .PARAMETER GroupAttributes
        Specify the array of group attributes to return with the DataTable.
    .EXAMPLE
        $dtGroups = _GetDistributionGroupInfo -Session $session GroupAttributes $GroupAttribs
    
        This would get all groups from the PowerShell $session and return attributes $GroupAttribs to $dtGroups
    .NOTES
        Version:
            - 5.1.2023.0726:    New function
            - 5.1.2024.0110:    Updated for MVAs to output better data
    #>
 
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the filter to use for the groups.")]
        [string]$Filter = "*",
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the Groups datatable to update with found information")]
        [System.Data.DataTable]$Groups,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the array of group attributes to return with the DataTable.")]
        [array]$GroupAttributes,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the EmailAddresses datatable to update with found information")]
        [System.Data.DataTable]$EmailAddresses,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute1 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute1,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute2 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute2,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute3 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute3,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute4 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute4,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ExtensionCustomAttribute5 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute5,

        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ManagedBy datatable to update with found information")]
        [System.Data.DataTable]$ManagedBy,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the AcceptMessagesOnlyFrom datatable to update with found information")]
        [System.Data.DataTable]$AcceptMessagesOnlyFrom,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the AcceptMessagesOnlyFromDLMembers datatable to update with found information")]
        [System.Data.DataTable]$AcceptMessagesOnlyFromDLMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the AcceptMessagesOnlyFromSendersOrMembers datatable to update with found information")]
        [System.Data.DataTable]$AcceptMessagesOnlyFromSendersOrMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the BypassModerationFromSendersOrMembers datatable to update with found information")]
        [System.Data.DataTable]$BypassModerationFromSendersOrMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the GrantSendOnBehalfTo datatable to update with found information")]
        [System.Data.DataTable]$GrantSendOnBehalfTo,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the ModeratedBy datatable to update with found information")]
        [System.Data.DataTable]$ModeratedBy,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the RejectMessagesFrom datatable to update with found information")]
        [System.Data.DataTable]$RejectMessagesFrom,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the RejectMessagesFromDLMembers datatable to update with found information")]
        [System.Data.DataTable]$RejectMessagesFromDLMembers,
                
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the RejectMessagesFromSendersOrMembers datatable to update with found information")]
        [System.Data.DataTable]$RejectMessagesFromSendersOrMembers,
        
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$Recipients
    )
    
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetUnifiedGroupInfo"
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Building base datatable"
            $dtUnifiedGroups = Get-UnifiedGroup -ResultSize 1 -Filter $Filter -WarningAction SilentlyContinue -Verbose:$false -ErrorAction Stop | Select-Object -Property $GroupAttributes | ConvertTo-DataTable
            
            foreach ($column in $dtUnifiedGroups.Columns) { if (-not $Groups.Columns.Contains($column.ColumnName)) { $Groups.Columns.Add($column.ColumnName, $column.DataType) } }
            if (-not $Groups.Columns.Contains("GroupDomain")) { [void]$Groups.Columns.Add("GroupDomain", "string") }
            if ($Groups.Rows.Count -le 0) { $Groups.columns["Description"].Datatype = "System.String" }

            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with group information"
            Get-UnifiedGroup -ResultSize Unlimited -Verbose:$false -Filter $Filter -ErrorAction Stop | Select-Object -Property $GroupAttributes | ForEach-Object {
                $drNewRow = $Groups.NewRow()
                $group = $_
                
                ForEach ($element in $_.PSObject.Properties) {
                    $columnName = $element.Name
                    $columnValue = $element.Value
                    
                    if ([string]::IsNullorEmpty($columnValue) -or $columnValue.ToString() -eq "Unlimited") {
                        $columnValue = [DBNull]::Value
                    } else {
                        switch ($columnName) {
                            "EmailAddresses" {
                                ForEach ($entry in $columnValue) {
                                    $drNewAddressRow = $EmailAddresses.NewRow()
                                    $drNewAddressRow["GroupDomain"] = [string]$GroupDomain
                                    $drNewAddressRow["GroupGuid"] = $group.Guid
                                    $drNewAddressRow["EmailAddresses"] = [string]$entry
                                    [void]$EmailAddresses.Rows.Add($drNewAddressRow)
                                }#foreach
                            }#EmailAddresses
                            "ExtensionCustomAttribute1" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute1Row = $ExtensionCustomAttribute1.NewRow()
                                    $drExtCustomAttribute1Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute1Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute1Row["ExtensionCustomAttribute1"] = [string]$entry
                                    [void]$ExtensionCustomAttribute1.Rows.Add($drExtCustomAttribute1Row)
                                }#foreach
                            }#ExtensionCustomAttribute1
                            "ExtensionCustomAttribute2" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute2Row = $ExtensionCustomAttribute2.NewRow()
                                    $drExtCustomAttribute2Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute2Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute2Row["ExtensionCustomAttribute2"] = [string]$entry
                                    [void]$ExtensionCustomAttribute2.Rows.Add($drExtCustomAttribute2Row)
                                }#foreach
                            }#ExtensionCustomAttribute2
                            "ExtensionCustomAttribute3" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute3Row = $ExtensionCustomAttribute3.NewRow()
                                    $drExtCustomAttribute3Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute3Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute3Row["ExtensionCustomAttribute3"] = [string]$entry
                                    [void]$ExtensionCustomAttribute3.Rows.Add($drExtCustomAttribute3Row)
                                }#foreach
                            }#ExtensionCustomAttribute3
                            "ExtensionCustomAttribute4" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute4Row = $ExtensionCustomAttribute4.NewRow()
                                    $drExtCustomAttribute4Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute4Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute4Row["ExtensionCustomAttribute4"] = [string]$entry
                                    [void]$ExtensionCustomAttribute4.Rows.Add($drExtCustomAttribute4Row)
                                }#foreach
                            }#ExtensionCustomAttribute4
                            "ExtensionCustomAttribute5" {
                                ForEach ($entry in $columnValue) {
                                    $drExtCustomAttribute5Row = $ExtensionCustomAttribute5.NewRow()
                                    $drExtCustomAttribute5Row["GroupDomain"] = [string]$GroupDomain
                                    $drExtCustomAttribute5Row["GroupGuid"] = $group.Guid
                                    $drExtCustomAttribute5Row["ExtensionCustomAttribute5"] = [string]$entry
                                    [void]$ExtensionCustomAttribute5.Rows.Add($drExtCustomAttribute5Row)
                                }#foreach
                            }#ExtensionCustomAttribute5
                            "ManagedBy" {
                                ForEach ($entry in $columnValue) {
                                    $ManagedByUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $ManagedByUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]
                                    
                                    if ($ManagedByUser) {
                                        $drManagedByRow = $ManagedBy.NewRow()
                                        $drManagedByRow["ManagedByRecipientTypeDetails"] = $ManagedByUser.RecipientTypeDetails
                                        $drManagedByRow["ManagedByPrimarySmtpAddress"] = $ManagedByUser.PrimarySmtpAddress
                                        $drManagedByRow["ManagedByAlias"] = $ManagedByUser.alias
                                        $drManagedByRow["ManagedByDisplayName"] = $ManagedByUser.DisplayName
                                        $drManagedByRow["ManagedByGuid"] = $ManagedByUser.Guid
                                        $drManagedByRow["ManagedByName"] = $ManagedByUser.Name
                                        $drManagedByRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drManagedByRow["GroupAlias"] = $group.alias
                                        $drManagedByRow["GroupDisplayName"] = $group.DisplayName
                                        $drManagedByRow["GroupName"] = $group.Name
                                        $drManagedByRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drManagedByRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drManagedByRow["GroupDomain"] = $GroupDomain
                                        [void]$ManagedBy.Rows.Add($drManagedByRow)
                                    }#if
                                }#foreach
                            }#ManagedBy
                            "AcceptMessagesOnlyFrom" {
                                ForEach ($entry in $columnValue) {
                                    $AcceptFrom = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $AcceptFrom = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($AcceptFrom) {
                                        $drAcceptMessagesOnlyFromRow = $AcceptMessagesOnlyFrom.NewRow()
                                        $drAcceptMessagesOnlyFromRow["ManagedByRecipientTypeDetails"] = $AcceptFrom.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromRow["ManagedByPrimarySmtpAddress"] = $AcceptFrom.PrimarySmtpAddress
                                        $drAcceptMessagesOnlyFromRow["ManagedByAlias"] = $AcceptFrom.alias
                                        $drAcceptMessagesOnlyFromRow["ManagedByDisplayName"] = $AcceptFrom.DisplayName
                                        $drAcceptMessagesOnlyFromRow["ManagedByGuid"] = $AcceptFrom.Guid
                                        $drAcceptMessagesOnlyFromRow["ManagedByName"] = $AcceptFrom.Name
                                        $drAcceptMessagesOnlyFromRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drAcceptMessagesOnlyFromRow["GroupAlias"] = $group.alias
                                        $drAcceptMessagesOnlyFromRow["GroupDisplayName"] = $group.DisplayName
                                        $drAcceptMessagesOnlyFromRow["GroupName"] = $group.Name
                                        $drAcceptMessagesOnlyFromRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drAcceptMessagesOnlyFromRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromRow["GroupDomain"] = $GroupDomain
                                        [void]$AcceptMessagesOnlyFrom.Rows.Add($drAcceptMessagesOnlyFromRow)
                                    }
                                }#foreach
                            }#AcceptMessagesOnlyFrom
                            "AcceptMessagesOnlyFromDLMembers" {
                                ForEach ($entry in $columnValue) {
                                    $AcceptFromDLMember = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $AcceptFromDLMember = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($AcceptFromDLMember) {
                                        $drAcceptMessagesOnlyFromDLMembersRow = $AcceptMessagesOnlyFromDLMembers.NewRow()
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByRecipientTypeDetails"] = $AcceptFromDLMember.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByPrimarySmtpAddress"] = $AcceptFromDLMember.PrimarySmtpAddress
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByAlias"] = $AcceptFromDLMember.alias
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByDisplayName"] = $AcceptFromDLMember.DisplayName
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByGuid"] = $AcceptFromDLMember.Guid
                                        $drAcceptMessagesOnlyFromDLMembersRow["ManagedByName"] = $AcceptFromDLMember.Name
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupAlias"] = $group.alias
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupName"] = $group.Name
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromDLMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$AcceptMessagesOnlyFromDLMembers.Rows.Add($drAcceptMessagesOnlyFromDLMembersRow)
                                    }
                                }#foreach
                            }#AcceptMessagesOnlyFromDLMembers
                            "AcceptMessagesOnlyFromSendersOrMembers" {
                                ForEach ($entry in $columnValue) {
                                    $AcceptFromSender = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $AcceptFromSender = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($AcceptFromSender) {
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow = $AcceptMessagesOnlyFromSendersOrMembers.NewRow()
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByRecipientTypeDetails"] = $AcceptFromSender.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByPrimarySmtpAddress"] = $AcceptFromSender.PrimarySmtpAddress
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByAlias"] = $AcceptFromSender.alias
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByDisplayName"] = $AcceptFromSender.DisplayName
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByGuid"] = $AcceptFromSender.Guid
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["ManagedByName"] = $AcceptFromSender.Name
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupAlias"] = $group.alias
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupName"] = $group.Name
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drAcceptMessagesOnlyFromSendersOrMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$AcceptMessagesOnlyFromSendersOrMembers.Rows.Add($drAcceptMessagesOnlyFromSendersOrMembersRow)
                                    }
                                }#foreach
                            }#AcceptMessagesOnlyFromSendersOrMembers
                            "BypassModerationFromSendersOrMembers" {
                                ForEach ($entry in $columnValue) {
                                    $BypassModerationUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $BypassModerationUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($BypassModerationUser) {
                                        $drBypassModerationFromSendersOrMembersRow = $BypassModerationFromSendersOrMembers.NewRow()
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByRecipientTypeDetails"] = $BypassModerationUser.RecipientTypeDetails
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByPrimarySmtpAddress"] = $BypassModerationUser.PrimarySmtpAddress
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByAlias"] = $BypassModerationUser.alias
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByDisplayName"] = $BypassModerationUser.DisplayName
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByGuid"] = $BypassModerationUser.Guid
                                        $drBypassModerationFromSendersOrMembersRow["ManagedByName"] = $BypassModerationUser.Name
                                        $drBypassModerationFromSendersOrMembersRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drBypassModerationFromSendersOrMembersRow["GroupAlias"] = $group.alias
                                        $drBypassModerationFromSendersOrMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drBypassModerationFromSendersOrMembersRow["GroupName"] = $group.Name
                                        $drBypassModerationFromSendersOrMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drBypassModerationFromSendersOrMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drBypassModerationFromSendersOrMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$BypassModerationFromSendersOrMembers.Rows.Add($drBypassModerationFromSendersOrMembersRow)
                                    }
                                }#foreach
                            }#BypassModerationFromSendersOrMembers
                            "GrantSendOnBehalfTo" {
                                ForEach ($entry in $columnValue) {
                                    $SendOnBehalfUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $SendOnBehalfUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($SendOnBehalfUser) {
                                        $drGrantSendOnBehalfToRow = $GrantSendOnBehalfTo.NewRow()
                                        $drGrantSendOnBehalfToRow["ManagedByRecipientTypeDetails"] = $SendOnBehalfUser.RecipientTypeDetails
                                        $drGrantSendOnBehalfToRow["ManagedByPrimarySmtpAddress"] = $SendOnBehalfUser.PrimarySmtpAddress
                                        $drGrantSendOnBehalfToRow["ManagedByAlias"] = $SendOnBehalfUser.alias
                                        $drGrantSendOnBehalfToRow["ManagedByDisplayName"] = $SendOnBehalfUser.DisplayName
                                        $drGrantSendOnBehalfToRow["ManagedByGuid"] = $SendOnBehalfUser.Guid
                                        $drGrantSendOnBehalfToRow["ManagedByName"] = $SendOnBehalfUser.Name
                                        $drGrantSendOnBehalfToRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drGrantSendOnBehalfToRow["GroupAlias"] = $group.alias
                                        $drGrantSendOnBehalfToRow["GroupDisplayName"] = $group.DisplayName
                                        $drGrantSendOnBehalfToRow["GroupName"] = $group.Name
                                        $drGrantSendOnBehalfToRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drGrantSendOnBehalfToRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drGrantSendOnBehalfToRow["GroupDomain"] = $GroupDomain
                                        [void]$GrantSendOnBehalfTo.Rows.Add($drGrantSendOnBehalfToRow)
                                    }
                                }#foreach
                            }#GrantSendOnBehalfTo
                            "ModeratedBy" {
                                ForEach ($entry in $columnValue) {
                                    $ModeratedByUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $ModeratedByUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($ModeratedByUser) {
                                        $drModeratedByRow = $ModeratedBy.NewRow()
                                        $drModeratedByRow["ManagedByRecipientTypeDetails"] = $ModeratedByUser.RecipientTypeDetails
                                        $drModeratedByRow["ManagedByPrimarySmtpAddress"] = $ModeratedByUser.PrimarySmtpAddress
                                        $drModeratedByRow["ManagedByAlias"] = $ModeratedByUser.alias
                                        $drModeratedByRow["ManagedByDisplayName"] = $ModeratedByUser.DisplayName
                                        $drModeratedByRow["ManagedByGuid"] = $ModeratedByUser.Guid
                                        $drModeratedByRow["ManagedByName"] = $ModeratedByUser.Name
                                        $drModeratedByRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drModeratedByRow["GroupAlias"] = $group.alias
                                        $drModeratedByRow["GroupDisplayName"] = $group.DisplayName
                                        $drModeratedByRow["GroupName"] = $group.Name
                                        $drModeratedByRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drModeratedByRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drModeratedByRow["GroupDomain"] = $GroupDomain
                                        [void]$ModeratedBy.Rows.Add($drModeratedByRow)
                                    }
                                }#foreach
                            }#ModeratedBy
                            "RejectMessagesFrom" {
                                ForEach ($entry in $columnValue) {
                                    $RejectFromUser = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $RejectFromUser = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($RejectFromUser) {
                                        $drRejectMessagesFromRow = $RejectMessagesFrom.NewRow()
                                        $drRejectMessagesFromRow["ManagedByRecipientTypeDetails"] = $RejectFromUser.RecipientTypeDetails
                                        $drRejectMessagesFromRow["ManagedByPrimarySmtpAddress"] = $RejectFromUser.PrimarySmtpAddress
                                        $drRejectMessagesFromRow["ManagedByAlias"] = $RejectFromUser.alias
                                        $drRejectMessagesFromRow["ManagedByDisplayName"] = $RejectFromUser.DisplayName
                                        $drRejectMessagesFromRow["ManagedByGuid"] = $RejectFromUser.Guid
                                        $drRejectMessagesFromRow["ManagedByName"] = $RejectFromUser.Name
                                        $drRejectMessagesFromRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drRejectMessagesFromRow["GroupAlias"] = $group.alias
                                        $drRejectMessagesFromRow["GroupDisplayName"] = $group.DisplayName
                                        $drRejectMessagesFromRow["GroupName"] = $group.Name
                                        $drRejectMessagesFromRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drRejectMessagesFromRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drRejectMessagesFromRow["GroupDomain"] = $GroupDomain
                                        [void]$RejectMessagesFrom.Rows.Add($drRejectMessagesFromRow)
                                    }
                                }#foreach
                            }#RejectMessagesFrom
                            "RejectMessagesFromDLMembers" {
                                ForEach ($entry in $columnValue) {
                                    $RejectFromDL = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $RejectFromDL = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($RejectFromDL) {
                                        $drRejectMessagesFromDLMembersRow = $RejectMessagesFromDLMembers.NewRow()
                                        $drRejectMessagesFromDLMembersRow["ManagedByRecipientTypeDetails"] = $RejectFromDL.RecipientTypeDetails
                                        $drRejectMessagesFromDLMembersRow["ManagedByPrimarySmtpAddress"] = $RejectFromDL.PrimarySmtpAddress
                                        $drRejectMessagesFromDLMembersRow["ManagedByAlias"] = $RejectFromDL.alias
                                        $drRejectMessagesFromDLMembersRow["ManagedByDisplayName"] = $RejectFromDL.DisplayName
                                        $drRejectMessagesFromDLMembersRow["ManagedByGuid"] = $RejectFromDL.Guid
                                        $drRejectMessagesFromDLMembersRow["ManagedByName"] = $RejectFromDL.Name
                                        $drRejectMessagesFromDLMembersRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drRejectMessagesFromDLMembersRow["GroupAlias"] = $group.alias
                                        $drRejectMessagesFromDLMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drRejectMessagesFromDLMembersRow["GroupName"] = $group.Name
                                        $drRejectMessagesFromDLMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drRejectMessagesFromDLMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drRejectMessagesFromDLMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$RejectMessagesFromDLMembers.Rows.Add($drRejectMessagesFromDLMembersRow)
                                    }
                                }#foreach
                            }#RejectMessagesFromDLMembers
                            "RejectMessagesFromSendersOrMembers" {
                                ForEach ($entry in $columnValue) {
                                    $RejectFromSender = ""
                                    $CheckUser = $entry -replace "'", "''"
                                    $RejectFromSender = ($Recipients.Select("Identity = '$CheckUser'"))[0]

                                    if ($RejectFromSender) {
                                        $drRejectMessagesFromSendersOrMembersRow = $RejectMessagesFromSendersOrMembers.NewRow()
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByRecipientTypeDetails"] = $RejectFromSender.RecipientTypeDetails
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByPrimarySmtpAddress"] = $RejectFromSender.PrimarySmtpAddress
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByAlias"] = $RejectFromSender.alias
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByDisplayName"] = $RejectFromSender.DisplayName
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByGuid"] = $RejectFromSender.Guid
                                        $drRejectMessagesFromSendersOrMembersRow["ManagedByName"] = $RejectFromSender.Name
                                        $drRejectMessagesFromSendersOrMembersRow["GroupGuid"] = [guid]($drNewRow["Guid"]).Guid
                                        $drRejectMessagesFromSendersOrMembersRow["GroupAlias"] = $group.alias
                                        $drRejectMessagesFromSendersOrMembersRow["GroupDisplayName"] = $group.DisplayName
                                        $drRejectMessagesFromSendersOrMembersRow["GroupName"] = $group.Name
                                        $drRejectMessagesFromSendersOrMembersRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                                        $drRejectMessagesFromSendersOrMembersRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                                        $drRejectMessagesFromSendersOrMembersRow["GroupDomain"] = $GroupDomain
                                        [void]$RejectMessagesFromSendersOrMembers.Rows.Add($drRejectMessagesFromSendersOrMembersRow)
                                    }
                                }#foreach
                            }#RejectMessagesFromSendersOrMembers
                            "MaxSendSize" { $drNewRow["MaxSendSize"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            "MaxReceiveSize" { $drNewRow["MaxReceiveSize"] = [math]::Round(($columnValue.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) }
                            default {
                                if ($columnValue.gettype().Name -eq "ArrayList") {
                                    $drNewRow["$columnName"] = [string]$columnValue.Clone()
                                } else {
                                    $drNewRow["$columnName"] = $columnValue
                                }#if/else
                            }#default
                        }#switch
                    }#if/else
                }#loop through each property
                $drNewRow["GroupDomain"] = [string]$GroupDomain

                [void]$Groups.Rows.Add($drNewRow)
            }#invoke/foreach

            return @(, $Groups)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather group information"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
    
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetUnifiedGroupInfo"
    }#end
}#function _GetUnifiedGroupInfo

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

Function _CreateMembershipTable {
    <#
    .SYNOPSIS
        Create a blank membership datatable
    .EXAMPLE
        $dtGroupMembers = _CreateMembershipTable
    
        This would create a blank membership datatable
    .NOTES
        Version:
            - 5.1.2023.0727:    New function
    #>
 
    [CmdletBinding()]
    Param()
    
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _CreateMembershipTable"
    }#begin
    
    process	{
        try	{
            $dtMembers = New-Object System.Data.DataTable
            $col1 = New-Object System.Data.DataColumn GroupDomain, ([string])
            $col2 = New-Object System.Data.DataColumn GroupGuid, ([Guid])
            $col3 = New-Object System.Data.DataColumn GroupAlias, ([string])
            $col4 = New-Object System.Data.DataColumn GroupDisplayName, ([string])
            $col5 = New-Object System.Data.DataColumn GroupName, ([string])
            $col6 = New-Object System.Data.DataColumn GroupPrimarySmtpAddress, ([string])
            $col7 = New-Object System.Data.DataColumn GroupRecipientTypeDetails, ([string])
            $col8 = New-Object System.Data.DataColumn MemberRecipientTypeDetails, ([string])
            $col9 = New-Object System.Data.DataColumn MemberPrimarySmtpAddress, ([string])
            $col10 = New-Object System.Data.DataColumn MemberAlias, ([string])
            $col11 = New-Object System.Data.DataColumn MemberDisplayName, ([string])
            $col12 = New-Object System.Data.DataColumn MemberGuid, ([string])
            $col13 = New-Object System.Data.DataColumn MemberName, ([string])
            $dtMembers.Columns.Add($col1)
            $dtMembers.Columns.Add($col2)
            $dtMembers.Columns.Add($col3)
            $dtMembers.Columns.Add($col4)
            $dtMembers.Columns.Add($col5)
            $dtMembers.Columns.Add($col6)
            $dtMembers.Columns.Add($col7)
            $dtMembers.Columns.Add($col8)
            $dtMembers.Columns.Add($col9)
            $dtMembers.Columns.Add($col10)
            $dtMembers.Columns.Add($col11)
            $dtMembers.Columns.Add($col12)
            $dtMembers.Columns.Add($col13)
            [System.Data.DataColumn[]]$KeyColumn = ($dtMembers.Columns["GroupDomain"], $dtMembers.Columns["GroupGuid"], $dtMembers.Columns["MemberGUID"])
            $dtMembers.PrimaryKey = $KeyColumn
            return @(, $dtMembers)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to create membership datatable"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $false
        }#try/catch
    }#process
    
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _CreateMembershipTable"
    }#end
}#function _CreateMembershipTable

Function _GetDistributionGroupMembership {
    <#
    .SYNOPSIS
        Collects the necessary distribution group cache and returns a datatable with the results
    .PARAMETER Session
        The PSSession to run the command against
    .PARAMETER Filter
        The optional parameter for a filter to be used when querying groups
   .PARAMETER GroupAttributes
        Specify the array of group attributes to return with the DataTable.
    .EXAMPLE
        $dtGroups = _GetDistributionGroupInfo -Session $session GroupAttributes $GroupAttribs
    
        This would get all groups from the PowerShell $session and return attributes $GroupAttribs to $dtGroups
    .NOTES
        Version:
            - 5.1.2023.0727:    New function
    #>
 
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the Groups datatable to update with found information")]
        [System.Data.DataTable]$Groups,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the Members datatable to update with found information")]
        [System.Data.DataTable]$Members
    )
    
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetDistributionGroupMembership"
        $MemberAttributes = 'RecipientTypeDetails', 'DisplayName', 'alias', 'Guid', 'Name', 'PrimarySMTPAddress'
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with group membership information"

            foreach ($group in $dtGroups.Select("RecipientTypeDetails = 'MailUniversalDistributionGroup' OR RecipientTypeDetails = 'MailUniversalSecurityGroup'")) {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Processing $($group.DisplayName)"
                Get-DistributionGroupMember -Identity ($group.Guid) -Verbose:$false -ErrorAction Stop | Select-Object -Property $MemberAttributes | ForEach-Object {
                    $drNewRow = $Members.NewRow()
                    
                    $drNewRow["MemberRecipientTypeDetails"] = $_.RecipientTypeDetails
                    $drNewRow["MemberPrimarySmtpAddress"] = $_.PrimarySmtpAddress
                    $drNewRow["MemberAlias"] = $_.alias
                    $drNewRow["MemberDisplayName"] = $_.DisplayName
                    $drNewRow["MemberGuid"] = $_.Guid
                    $drNewRow["MemberName"] = $_.Name
                    $drNewRow["GroupGuid"] = $group.Guid
                    $drNewRow["GroupAlias"] = $group.alias
                    $drNewRow["GroupDisplayName"] = $group.DisplayName
                    $drNewRow["GroupName"] = $group.Name
                    $drNewRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                    $drNewRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                    $drNewRow["GroupDomain"] = $group.GroupDomain
                    [void]$Members.Rows.Add($drNewRow)
                }#foreach
            }#foreach
            
            return @(, $Members)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather group membership information"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
    
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetDistributionGroupMembership"
    }#end
}#function _GetDistributionGroupMembership

Function _GetUnifiedGroupMembership {
    <#
    .SYNOPSIS
        Collects the necessary distribution group cache and returns a datatable with the results
    .PARAMETER Session
        The PSSession to run the command against
    .PARAMETER Filter
        The optional parameter for a filter to be used when querying groups
   .PARAMETER GroupAttributes
        Specify the array of group attributes to return with the DataTable.
    .EXAMPLE
        $dtGroups = _GetDistributionGroupInfo -Session $session GroupAttributes $GroupAttribs
    
        This would get all groups from the PowerShell $session and return attributes $GroupAttribs to $dtGroups
    .NOTES
        Version:
            - 5.1.2023.0727:    New function
    #>
 
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the Groups datatable to update with found information")]
        [System.Data.DataTable]$Groups,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the Members datatable to update with found information")]
        [System.Data.DataTable]$Members
    )
    
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetUnifiedGroupMembership"
        $MemberAttributes = 'RecipientTypeDetails', 'DisplayName', 'alias', 'Guid', 'Name', 'PrimarySMTPAddress'
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with unified group membership information"

            foreach ($group in $dtGroups.Select("RecipientTypeDetails = 'GroupMailbox'")) {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Processing $($group.DisplayName)"
                Get-UnifiedGroupLinks -Identity ($group.Guid) -LinkType Members -Verbose:$false -ErrorAction Stop | Select-Object -Property $MemberAttributes | ForEach-Object {
                    $drNewRow = $Members.NewRow()
                    
                    $drNewRow["MemberRecipientTypeDetails"] = $_.RecipientTypeDetails
                    $drNewRow["MemberPrimarySmtpAddress"] = $_.PrimarySmtpAddress
                    $drNewRow["MemberAlias"] = $_.alias
                    $drNewRow["MemberDisplayName"] = $_.DisplayName
                    $drNewRow["MemberGuid"] = $_.Guid
                    $drNewRow["MemberName"] = $_.Name
                    $drNewRow["GroupGuid"] = $group.Guid
                    $drNewRow["GroupAlias"] = $group.alias
                    $drNewRow["GroupDisplayName"] = $group.DisplayName
                    $drNewRow["GroupName"] = $group.Name
                    $drNewRow["GroupPrimarySmtpAddress"] = $group.PrimarySMTPAddress
                    $drNewRow["GroupRecipientTypeDetails"] = $group.RecipientTypeDetails
                    $drNewRow["GroupDomain"] = $group.GroupDomain
                    [void]$Members.Rows.Add($drNewRow)
                }#foreach
            }#foreach
            
            return @(, $Members)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather group membership information"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
    
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetUnifiedGroupMembership"
    }#end
}#function _GetUnifiedGroupMembership
#endregion

#region Active Development
Function _GetCustomSendAsPermissions {
    <#
    .SYNOPSIS
        Processes the group to get Send-As level permissions
    .PARAMETER Mailbox
        Datarow that is the group to check
    .PARAMETER Permissions
        Arraylist for the collected permissions
    .PARAMETER Exceptions
        Arraylist for the collected permissions exceptions
    .EXAMPLE
        _GetCustomSendAsPermissionsID -Group <datarow> -Permissions $arrPermissions -Exceptions $arrExceptions

        This would get Send-As permissions for <datarow>, save the results to $arrPermissions, and note any exceptions found in $arrExceptions
    .NOTES
        Version:
        - 5.1.2024.0109:    New function
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [System.Data.Datarow]$Group,

        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$Recipients,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.ArrayList]$Permissions,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.ArrayList]$Exceptions
    )
    
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetCustomSendAsPermissions"

        $GroupDisplayName = $Group.DisplayName.ToString()
        $GroupDN = $Group.DistinguishedName.ToString()
        $GroupSAM = $Group.samAccountName.ToString()
    }#begin
    
    process	{
        try {
            if ($ExchangeServer -eq "Online") {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Gathering <$GroupDisplayName> Send-As permissions from Exchange Online"
        
                $groupSendAs = Get-RecipientPermission -Identity $GroupDN -AccessRights SendAs -ResultSize Unlimited -ErrorAction Stop | `
                        Where-Object { $_.IsInherited -eq $False -and $_.Trustee -notlike "NT AUTHORITY\SELF" -and $_.Trustee -notlike "S-1-5-21*" } | Select-Object Identity, Trustee, AccessControlType
            } else {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Gathering <$GroupDisplayName> Send-As permissions from Active Directory"
                
                $Forest = ($GroupDN.substring($GroupDN.IndexOf("DC=") + 3)).Replace(",DC=", "_")
                
                if (-not (Get-PSDrive $Forest -ErrorAction SilentlyContinue)) {
                    $AltDC = ($GroupDN.substring($GroupDN.IndexOf("DC=") + 3)).Replace(",DC=", ".")
                    $AltForest = $AltDC.Replace(".", "_")
                    New-PSDrive -Name $AltForest -Scope Script -Root "" -PSProvider ActiveDirectory -Server $AltDC
                }
                
                $groupSendAs = (Get-Acl "$($Forest):$($GroupDN)" -ErrorAction Stop).access | Where-Object { $_.IsInherited -eq $false -and ($_.objecttype -eq $script:SendAsGUID) `
                        -and $_.IdentityReference -notlike "NT AUTHORITY\SELF" -and $_.AccessControlType -eq "Allow" -and $_.IdentityReference -notlike "*$GroupSAM" } | Select-Object -Expand IdentityReference
                $groupSendAs = $groupSendAs | Select-Object -Expand Value | Where-Object { $_ -notlike "S-1-5-21*" }
            }#if/else

            if ($groupSendAs) {
                foreach ($entry in $groupSendAs) {
                    $User = ""
                    $AccessAccount = ""
                    
                    if ($ExchangeServer -eq "Online") {
                        $AccessAccount = $entry.Trustee
                        
                        if ($AccessAccount -like "*@*") {
                            $User = ($Recipients.Select("UserPrincipalName = '$AccessAccount'"))[0]
                        } else {
                            #If the user is actually a group, it won't have the '@' in the name
                            $User = ($Recipients.Select("Identity = '$AccessAccount'"))[0]
                        }#if/else
                    } else {
                        $AccessAccount = $entry.ToString().Split('\')[1]
                        $User = ($Recipients.Select("sAMAccountName = '$AccessAccount'"))[0]
                    }#if/else

                    if ($User) {
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Processing user account <$AccessAccount> found"
                        $objNew = New-Object -TypeName PSCustomObject -Property @{
                            "GroupDisplayName"          = $GroupDisplayName
                            "GroupsAMAccountName"       = $GroupSAM
                            "GroupPrimarySmtpAddress"   = $Group.PrimarySMTPAddress.ToString()
                            "GroupGuid"                 = $Group.Guid.ToString()
                            "GroupOU"                   = $Group.OrganizationalUnit.ToString()
                            "GroupRecipientTypeDetails" = $Group.RecipientTypeDetails
                            "GroupAlias"                = $Group.Alias
                            "GroupName"                 = $Group.Name
                            "UserExchangeGuid"          = $User.ExchangeGuid
                            "UserGuid"                  = $User.Guid
                            "UserUserPrincipalName"     = $User.UserPrincipalName
                            "UsersamAccountName"        = $User.samAccountName
                            "UserPrimarySMTPAddress"    = $User.PrimarySmtpAddress.ToString()
                            "UserDisplayName"           = $User.DisplayName
                            "UserOU"                    = $User.OrganizationalUnit
                            "UserRecipientTypeDetails"  = $User.RecipientTypeDetails
                            "UserAlias"                 = $User.Alias
                            "UserName"                  = $User.Name
                            "AccessLevel"               = "Group"
                            "AccessRight"               = "Send-As"
                            "GroupDomain"               = $GroupDomain
                        }
                        [void]$Permissions.Add($objNew)
                    } else {
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "User account <$AccessAccount> NOT found"
                    
                        $WhenModified = (Get-Date).ToUniversalTime()
                        $objNew = New-Object -TypeName PSCustomObject -Property @{
                            "GroupDisplayName"        = $GroupDisplayName
                            "GroupsAMAccountName"     = $GroupSAM
                            "GroupPrimarySmtpAddress" = $GroupPrimarySMTPAddress
                            "UsersAMAccountName"      = $AccessAccount
                            "AccessLevel"             = "Group"
                            "AccessRight"             = "Send-As"
                            "WhenModifiedUTC"         = $WhenModified
                            "GroupDomain"             = $GroupDomain
                        }
                        [void]$Exceptions.Add($objNew)
                    }#if/else
                }#foreach permissions entry
            } else {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "No custom Send-As permissions for <$GroupDisplayName>"
            }#if/else
            
            return $true
        }#try
        catch {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to process <$GroupDisplayName> Send-As permissions"
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_

            return $null
        }#catch
    }#process
    
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetCustomSendAsPermissions"
    }#end
}#function _GetCustomSendAsPermissions

Function _GetRecipientCache {
    <#
    .SYNOPSIS
        Collects the necessary recipient cache and returns a datatable with the results
    .PARAMETER Filter
        The optional parameter for a filter to be used when querying recipient
    .EXAMPLE
        $dtRecipients = _GetRecipientsCache
    
        This would get all recipients from the PowerShell $session and return it to $dtRecipients
    
    .NOTES
        Version:
            - 5.1.2024.0110:    New function
    #>
        
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $false,
            HelpMessage = "Specify the filter to use for the recipients.")]
        [string]$Filter = "*"
    )
    
    begin {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetRecipientCache"

        $htFullAttribs = @{"RecipientType" = "System.String"; "RecipientTypeDetails" = "System.String"; "ExchangeGuid" = "System.Guid"; "UserPrincipalName" = "System.String"; `
                "samAccountName" = "System.String"; "PrimarySmtpAddress" = "System.String"; "DisplayName" = "System.String"; "OrganizationalUnit" = "System.String"; `
                "WindowsLiveID" = "System.String"; "DistinguishedName" = "System.String"; "Guid" = "System.Guid"; "Identity" = "System.String"; `
                "GrantSendOnBehalfTo" = "System.Collections.ArrayList"; "alias" = "System.String"; "Name" = "System.String"
        }

        $RecipientAttribs = @('RecipientType', 'RecipientTypeDetails', 'samAccountName', 'PrimarySmtpAddress', 'DisplayName', `
                'OrganizationalUnit', 'DistinguishedName', 'Guid', 'Identity', 'ExchangeGuid', 'alias', 'Name')
        $MailboxAttribs = @('ExchangeGuid', 'UserPrincipalName', 'WindowsLiveID', 'Guid', 'GrantSendOnBehalfTo')
        #$GroupAttribs = @('Guid')

        Out-Log -LoggingPreference $htLoggingPreference -Type Verbose -WriteBackToHost -Message "Building base datatable"
        $dtRecipients = New-Object System.Data.DataTable
        foreach ($attrib in $htFullAttribs.GetEnumerator()) {
            $Column = New-Object System.Data.DataColumn
            $Column.ColumnName = $attrib.Name
            $Column.DataType = [System.Type]::GetType($attrib.Value)
            $Column.AllowDBNull = $true
            [void]$dtRecipients.Columns.Add($Column)
            #Option: $dtRecipients.Columns.Add(Name, type)
        }
        $dtRecipients.PrimaryKey = $dtRecipients.Columns["Guid"]
    }#begin
    
    process	{
        try	{
            Out-Log -LoggingPreference $htLoggingPreference -Type Verbose -WriteBackToHost -Message "Populating datatable with recipient information"
            
            Get-Recipient -ResultSize Unlimited -Verbose:$false -Filter $Filter | Select-Object -Property $RecipientAttribs | ForEach-Object {
                $drNewRow = $dtRecipients.NewRow()
                foreach ($element in $_.PSObject.Properties) {
                    $columnName = $element.Name
                    $columnValue = $element.Value
                    
                    if ([string]::IsNullorEmpty($columnValue)) { $columnValue = [DBNull]::Value }

                    if ($columnValue.gettype().Name -eq "ArrayList") {
                        $drNewRow["$columnName"] = $columnValue.Clone()
                    } else {
                        $drNewRow["$columnName"] = $columnValue
                    }#if/else
                }#loop through each property

                [void]$dtRecipients.Rows.Add($drNewRow)
            }#invoke/foreach

            Out-Log -LoggingPreference $htLoggingPreference -Type Verbose -WriteBackToHost -Message "Populating datatable with related mailbox information"

            Get-Mailbox -ResultSize Unlimited -Verbose:$false -Filter $Filter | Select-Object -Property $MailboxAttribs | ForEach-Object {
                $row = $dtRecipients.Rows.Find($_.Guid)
                
                foreach ($element in $_.PSObject.Properties) {
                    $columnName = $element.Name
                    $columnValue = $element.Value
                    
                    if ($columnValue.gettype().Name -eq "ArrayList") {
                        $row["$columnName"] = $columnValue.Clone()
                    } else {
                        $row["$columnName"] = $columnValue
                    }
                }#loop through each property
            }#invoke/foreach
            
            Out-Log -LoggingPreference $htLoggingPreference -Type Verbose -WriteBackToHost -Message "Populating datatable with related group mailbox information"
            if ($script:boolScriptIsEMSOnline) {
                Get-Mailbox -ResultSize Unlimited -GroupMailbox -Verbose:$false -Filter $Filter | Select-Object -Property $MailboxAttribs | ForEach-Object {
                    $row = $dtRecipients.Rows.Find($_.Guid)
                    
                    foreach ($element in $_.PSObject.Properties) {
                        $columnName = $element.Name
                        $columnValue = $element.Value
                        
                        if ($columnValue.gettype().Name -eq "ArrayList") {
                            $row["$columnName"] = $columnValue.Clone()
                        } else {
                            $row["$columnName"] = $columnValue
                        }
                    }#loop through each property
                }#invoke/foreach
            }#if

            return @(, $dtRecipients)
        } catch {
            Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message "Error while trying to gather recipients`r"
            Out-Log -LoggingPreference $htLoggingPreference -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
    
    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetRecipientCache"
    }#end
}#function _GetRecipientCache
#endregion

#region Main Program
Write-Host "`r"
Write-Host "Script Written by Sterling Consulting`r"
Write-Host "All rights reserved. Proprietary and Confidential Material`r"
Write-Host "Exchange Distribution Group Inventory Script`r"
Write-Host "`r"

Write-Host "Script starting`r"

if (_ConfirmScriptRequirements) {
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Script requirements met"

    #Get recipients from system based on supplied filter
    # if ($GroupFilter -ne "") {
    #     $RecipientFilter = $RecipientFilter + " -and (" + $GroupFilter + ")"
    # }#if
    # Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Recipient filter being used: $RecipientFilter"

    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Retrieving recipient information. Be patient!"
    $dtRecipients = _GetRecipientCache -Filter $RecipientFilter
    
    if (-not $dtRecipients) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "No recipient information found. Unable to continue without recipient information. Exiting script"
        Exit $ExitCode
    }#if
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished retrieving recipient information"
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "$(($dtRecipients | Measure-Object).Count) Recipient entries collected"

    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Creating MVA datatables"
    $dtEmailAddresses = _CreateMVATable -Attribute "EmailAddresses"
    $dtExtensionCustomAttribute1 = _CreateMVATable -Attribute "ExtensionCustomAttribute1"
    $dtExtensionCustomAttribute2 = _CreateMVATable -Attribute "ExtensionCustomAttribute2"
    $dtExtensionCustomAttribute3 = _CreateMVATable -Attribute "ExtensionCustomAttribute3"
    $dtExtensionCustomAttribute4 = _CreateMVATable -Attribute "ExtensionCustomAttribute4"
    $dtExtensionCustomAttribute5 = _CreateMVATable -Attribute "ExtensionCustomAttribute5"
    $dtManagedBy = _CreateExpandedMVATable -Attribute "ManagedBy"
    $dtAcceptMessagesOnlyFrom = _CreateExpandedMVATable -Attribute "AcceptMessagesOnlyFrom"
    $dtAcceptMessagesOnlyFromDLMembers = _CreateExpandedMVATable -Attribute "AcceptMessagesOnlyFromDLMembers"
    $dtAcceptMessagesOnlyFromSendersOrMembers = _CreateExpandedMVATable -Attribute "AcceptMessagesOnlyFromSendersOrMembers"
    $dtBypassModerationFromSendersOrMembers = _CreateExpandedMVATable -Attribute "BypassModerationFromSendersOrMembers"
    $dtGrantSendOnBehalfTo = _CreateExpandedMVATable -Attribute "GrantSendOnBehalfTo"
    $dtModeratedBy = _CreateExpandedMVATable -Attribute "ModeratedBy"
    $dtRejectMessagesFrom = _CreateExpandedMVATable -Attribute "RejectMessagesFrom"
    $dtRejectMessagesFromDLMembers = _CreateExpandedMVATable -Attribute "RejectMessagesFromDLMembers"
    $dtRejectMessagesFromSendersOrMembers = _CreateExpandedMVATable -Attribute "RejectMessagesFromSendersOrMembers"
    if ($IncludeMembership) { $dtMembers = _CreateMembershipTable } #Can probably switch this to standard _CreateMVATable
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Finished creating MVA datatables"

    $GroupInfoParams = @{
        "Filter"                                 = $GroupFilter
        "Groups"                                 = $dtGroups
        "GroupAttributes"                        = $arrGroupAttribs
        "EmailAddresses"                         = $dtEmailAddresses
        "ExtensionCustomAttribute1"              = $dtExtensionCustomAttribute1
        "ExtensionCustomAttribute2"              = $dtExtensionCustomAttribute2
        "ExtensionCustomAttribute3"              = $dtExtensionCustomAttribute3
        "ExtensionCustomAttribute4"              = $dtExtensionCustomAttribute4
        "ExtensionCustomAttribute5"              = $dtExtensionCustomAttribute5
        "ManagedBy"                              = $dtManagedBy
        "AcceptMessagesOnlyFrom"                 = $dtAcceptMessagesOnlyFrom
        "AcceptMessagesOnlyFromDLMembers"        = $dtAcceptMessagesOnlyFromDLMembers
        "AcceptMessagesOnlyFromSendersOrMembers" = $dtAcceptMessagesOnlyFromSendersOrMembers
        "BypassModerationFromSendersOrMembers"   = $dtBypassModerationFromSendersOrMembers
        "GrantSendOnBehalfTo"                    = $dtGrantSendOnBehalfTo
        "ModeratedBy"                            = $dtModeratedBy
        "RejectMessagesFrom"                     = $dtRejectMessagesFrom
        "RejectMessagesFromDLMembers"            = $dtRejectMessagesFromDLMembers
        "RejectMessagesFromSendersOrMembers"     = $dtRejectMessagesFromSendersOrMembers
        "Recipients"                             = $dtRecipients
    }

    #Get Unified Group from Exchange Online
    if ($IncludeUnifiedGroups) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Retrieving unified group information"
        _GetUnifiedGroupInfo @GroupInfoParams | Out-Null
        if ($IncludeMembership) { _GetUnifiedGroupMembership -Groups $dtGroups -Members $dtMembers | Out-Null }
    }

    #Get Distribution Groups from Exchange
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Retrieving distribution group information"
    _GetDistributionGroupInfo @GroupInfoParams | Out-Null
    if ($IncludeMembership) { _GetDistributionGroupMembership -Groups $dtGroups -Members $dtMembers | Out-Null }

    #Get Dynamic Distribution Groups from Exchange
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Retrieving dynamic group information"
    _GetDynamicGroupInfo @GroupInfoParams | Out-Null

    if ($IncludePermissions) {
        foreach ($group in $dtGroups) {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting check for Send-As level permissions"
            _GetCustomSendAsPermissions -Group $group -Recipients $dtRecipients -Permissions $arrPermissions -Exceptions $arrPermissionsException | Out-Null
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing check for Send-As level permissions"
        }#foreach
    }#if

    if ($dtGroups.DefaultView.Count -le 0) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "No group information found. Unable to continue without group information. Exiting script"
        Exit $ExitCode
    }#if
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished retrieving group information"
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "$($dtGroups.Rows.Count) Group entries collected"

    if ($ExportLocation -eq "") {
        $ExportLocation = $script:strBaseLocation + "\Exchange"
    }
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting CSV to $ExportLocation with , delimiter"
        
    #Check for path/folder
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Checking for $ExportLocation"
    if (-not (Test-Path -Path $ExportLocation)) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Creating folder structure for $ExportLocation"
        New-Item -Path $ExportLocation -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    }
    
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting export of CSV"
    if ($dtGroups.DefaultView.Count -ge 1) {
        $dtGroups | Export-Csv -Path "$ExportLocation\GroupInfo_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
        
        if ($dtEmailAddresses.DefaultView.Count -ge 1) { $dtEmailAddresses | Export-Csv -Path "$ExportLocation\GroupInfo_EmailAddresses_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtExtensionCustomAttribute1.DefaultView.Count -ge 1) { $dtExtensionCustomAttribute1 | Export-Csv -Path "$ExportLocation\GroupInfo_ExtensionCustomAttribute1_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtExtensionCustomAttribute2.DefaultView.Count -ge 1) { $dtExtensionCustomAttribute2 | Export-Csv -Path "$ExportLocation\GroupInfo_ExtensionCustomAttribute2_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtExtensionCustomAttribute3.DefaultView.Count -ge 1) { $dtExtensionCustomAttribute3 | Export-Csv -Path "$ExportLocation\GroupInfo_ExtensionCustomAttribute3_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtExtensionCustomAttribute4.DefaultView.Count -ge 1) { $dtExtensionCustomAttribute4 | Export-Csv -Path "$ExportLocation\GroupInfo_ExtensionCustomAttribute4_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtExtensionCustomAttribute5.DefaultView.Count -ge 1) { $dtExtensionCustomAttribute5 | Export-Csv -Path "$ExportLocation\GroupInfo_ExtensionCustomAttribute5_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtManagedBy.DefaultView.Count -ge 1) { $dtManagedBy | Export-Csv -Path "$ExportLocation\GroupInfo_ManagedBy_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtAcceptMessagesOnlyFrom.DefaultView.Count -ge 1) { $dtAcceptMessagesOnlyFrom | Export-Csv -Path "$ExportLocation\GroupInfo_AcceptMessagesOnlyFrom_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtAcceptMessagesOnlyFromDLMembers.DefaultView.Count -ge 1) { $dtAcceptMessagesOnlyFromDLMembers | Export-Csv -Path "$ExportLocation\GroupInfo_AcceptMessagesOnlyFromDLMembers_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtAcceptMessagesOnlyFromSendersOrMembers.DefaultView.Count -ge 1) { $dtAcceptMessagesOnlyFromSendersOrMembers | Export-Csv -Path "$ExportLocation\GroupInfo_AcceptMessagesOnlyFromSendersOrMembers_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtBypassModerationFromSendersOrMembers.DefaultView.Count -ge 1) { $dtBypassModerationFromSendersOrMembers | Export-Csv -Path "$ExportLocation\GroupInfo_BypassModerationFromSendersOrMembers_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtGrantSendOnBehalfTo.DefaultView.Count -ge 1) { $dtGrantSendOnBehalfTo | Export-Csv -Path "$ExportLocation\GroupInfo_GrantSendOnBehalfTo_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtModeratedBy.DefaultView.Count -ge 1) { $dtModeratedBy | Export-Csv -Path "$ExportLocation\GroupInfo_ModeratedBy_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtRejectMessagesFrom.DefaultView.Count -ge 1) { $dtRejectMessagesFrom | Export-Csv -Path "$ExportLocation\GroupInfo_RejectMessagesFrom_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtRejectMessagesFromDLMembers.DefaultView.Count -ge 1) { $dtRejectMessagesFromDLMembers | Export-Csv -Path "$ExportLocation\GroupInfo_RejectMessagesFromDLMembers_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtRejectMessagesFromSendersOrMembers.DefaultView.Count -ge 1) { $dtRejectMessagesFromSendersOrMembers | Export-Csv -Path "$ExportLocation\GroupInfo_RejectMessagesFromSendersOrMembers_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($dtMembers.DefaultView.Count -ge 1) { $dtMembers | Export-Csv -Path "$ExportLocation\GroupInfo_Members_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($arrPermissions) { $arrPermissions | Export-Csv -Path "$ExportLocation\GroupInfo_SendAsPermissions_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        if ($arrPermissionsException) { $arrPermissionsException | Export-Csv -Path "$ExportLocation\GroupInfo_SendAsPermissionsExceptions_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null }
        
        $ExitCode = 0
    }
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished export of CSV"

    $RunTime = ((Get-Date).ToUniversalTime() - $dateStartTimeStamp)
    $RunTime = '{0:00}:{1:00}:{2:00}:{3:00}.{4:00}' -f $RunTime.Days, $RunTime.Hours, $RunTime.Minutes, $RunTime.Seconds, $RunTime.Milliseconds
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Run time was $RunTime"
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exit code is $ExitCode"
} else {
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Script requirements not met:"
    
    if (-not $script:boolScriptIsModulesLoaded) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Missing required PowerShell module(s) or could not load modules"
    }#if
}#if/else

Get-ConnectionInformation -ErrorAction SilentlyContinue -Verbose:$fasle | Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
Exit $ExitCode
#endregion