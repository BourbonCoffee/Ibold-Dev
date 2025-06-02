#Requires -version 5.0

#region Parameters
[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the group member CSV file")]
    [string]$MemberFilePath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the credential XML file to be used")]
    [string]$CredentialsPath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify to connect to GCC High tenants")]
    [switch]$GCCHigh,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify to create Unified/M365 groups")]
    [switch]$IncludeUnifiedGroups,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the desired prefix of the group names")]
    [string]$Prefix,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the desired suffix of the group names")]
    [string]$Suffix
)
#endregion

#region User Variables
#Very little reason to change these
$InformationPreference = "Continue"

if ($DebugPreference -eq "Confirm" -or $DebugPreference -eq "Inquire") { $DebugPreference = "Continue" }
#endregion

#region Static Variables
#Don't change these
Set-Variable -Name strBaseLocation -WhatIf:$false -Option AllScope -Scope Script
Set-Variable -Name dateStartTimeStamp -WhatIf:$false -Option AllScope -Scope Script -Value (Get-Date).ToUniversalTime()
Set-Variable -Name strLogTimeStamp -WhatIf:$false -Option AllScope -Scope Script -Value $dateStartTimeStamp.ToString("MMddyyyy_HHmmss")
Set-Variable -Name strLogFile -WhatIf:$false -Option ReadOnly -Scope Script
Set-Variable -Name htLoggingPreference -WhatIf:$false -Option AllScope -Scope Script -Value @{"InformationPreference" = $InformationPreference; `
        "WarningPreference" = $WarningPreference; "ErrorActionPreference" = $ErrorActionPreference; "VerbosePreference" = $VerbosePreference; "DebugPreference" = $DebugPreference
}
Set-Variable -Name verScript -WhatIf:$false -Option AllScope -Scope Script -Value "5.1.2024.0422"

Set-Variable -Name boolScriptIsModulesLoaded -WhatIf:$false -Option AllScope -Scope Script -Value $false
Set-Variable -Name ExitCode -WhatIf:$false -Option AllScope -Scope Script -Value 1

New-Object System.Data.DataTable | Set-Variable dtMembers -WhatIf:$false -Option AllScope -Scope Script
New-Object System.Collections.ArrayList | Set-Variable arrExceptions -WhatIf:$false -Option AllScope -Scope Script
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
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
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
        - 5.1.2023.0727:    New function
        - 5.1.2024.0105:    Updated to allow for GCCHigh connections
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
                Import-Module ExchangeOnlineManagement -ErrorAction Stop -Verbose:$false
                $script:boolScriptIsModulesLoaded = $true
            } else {
                Write-Warning "Missing ExchangeOnlineManagement PowerShell module`r"
                $script:boolScriptIsModulesLoaded = $false
            }#if/else
        } catch {
            Write-Error "Unable to load ExchangeOnlineManagement PowerShell module`r"
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
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: MemberFilePath = $MemberFilePath"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: CredentialsPath = $CredentialsPath"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: GCCHigh = $GCCHigh"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludeUnifiedGroups = $IncludeUnifiedGroups"
    }#begin
    
    process {
        if ($script:boolScriptIsModulesLoaded) {
            try {
                $global:VerbosePreference = "SilentlyContinue"

                $ConnectSplat = @{
                    "GCCHigh" = $GCCHigh
                }

                if ($CredentialsPath) {
                    $ConnectSplat.Add("Credential", $(Import-Clixml $CredentialsPath))
                }

                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Connecting to Exchange Online"
                Connect-Exchange @ConnectSplat
                
                if ($htLoggingPreference['VerbosePreference'] -eq "Continue") { $global:VerbosePreference = "Continue" }#if                
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

#endregion

#region Main Program
Write-Host "`r"
Write-Host "Script Written by Sterling Consulting`r"
Write-Host "All rights reserved. Proprietary and Confidential Material`r"
Write-Host "Exchange Distribution Group Membership Script`r"
Write-Host "`r"

Write-Host "Script starting`r"

$WhatIfPreference = $false
if (_ConfirmScriptRequirements) {
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Script requirements met"

    if ($IncludeUnifiedGroups) {
        $dtMembers = Import-Csv $MemberFilePath | ConvertTo-DataTable
    } else {
        $dtMembers = Import-Csv $MemberFilePath | Where-Object { $_.GroupRecipientTypeDetails -notmatch "GroupMailbox" } | ConvertTo-DataTable
    }
    
    if ($dtMembers.Rows.Count -ge 1) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "$($dtMembers.Rows.Count) Members found"

        foreach ($member in $dtMembers.Rows) {
            try {
                $GroupMemberSplat = @{
                    "Identity" = $Prefix + $member.GroupAlias + $Suffix
                    "Member"   = $member.targetEmailAddress
                }

                if ($member.GroupRecipientTypeDetails -match "MailUniversal") {
                    if ($htLoggingPreference['WhatIfPreference']) { $WhatIfPreference = $true }#if
                    if ($PSCmdlet.ShouldProcess("Add-DistributionGroupMember with parameters: $($GroupMemberSplat | Out-String)", "", "")) {
                        $newMember = Add-DistributionGroupMember @GroupMemberSplat -Verbose:$false -ErrorAction Stop
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Added $($member.MemberDisplayName) to $($member.GroupDisplayName)"
                    }

                    $WhatIfPreference = $false
                } elseif ($IncludeUnifiedGroups) {
                    if ($htLoggingPreference['WhatIfPreference']) { $WhatIfPreference = $true }#if
                    if ($PSCmdlet.ShouldProcess("Add-UnifiedGroupLinks with parameters: $($GroupMemberSplat | Out-String)", "", "")) {
                        $newMember = Add-UnifiedGroupLinks @GroupMemberSplat -LinkType Members -Verbose:$false -ErrorAction Stop
                        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Added $($member.MemberDisplayName) to $($member.GroupDisplayName)"
                    }

                    $WhatIfPreference = $false
                }#if/else

                $ExitCode = 0
            } catch {
                $ErrorMessage = $_.Exception.Message

                if ($ErrorMessage -notmatch "already a member of the group") {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Failed to add $($member.MemberDisplayName) to $($member.GroupDisplayName)"
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $ErrorMessage
                    
                    [void]$arrExceptions.Add($member)

                    $ExitCode = 1
                }
            }#try/catch
        }#foreach
    }#if

    
    if ($arrExceptions.Count -ge 1) {
        $ExportLocation = $script:strBaseLocation + "\Exchange"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting exception CSV to $ExportLocation"
        
        $arrExceptions | Export-Csv -Path "$ExportLocation\GroupMembership_Exceptions_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
    }#if

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