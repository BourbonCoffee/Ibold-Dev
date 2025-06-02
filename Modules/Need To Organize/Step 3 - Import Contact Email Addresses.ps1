#Requires -version 5.0
#Requires -Modules Sterling

#region Parameters
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the contact email addresses CSV file")]
    [string]$AddressesFilePath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the credential XML file to be used")]
    [string]$CredentialsPath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the attribute to use to store the source contact guid. It will default to customAttribute1 if not specified.")]
    [string]$MappingAttribute = "customAttribute1"
)
#endregion

#region User Variables
#Very little reason to change these
$InformationPreference = "Continue"

if ($DebugPreference -eq "Confirm" -or $DebugPreference -eq "Inquire") {$DebugPreference = "Continue"}
#endregion

#region Static Variables
#Don't change these
Set-Variable -Name strBaseLocation -Option AllScope -Scope Script
Set-Variable -Name dateStartTimeStamp -Option AllScope -Scope Script -Value (Get-Date).ToUniversalTime()
Set-Variable -Name strLogTimeStamp -Option AllScope -Scope Script -Value $dateStartTimeStamp.ToString("MMddyyyy_HHmmss")
Set-Variable -Name strLogFile -Option ReadOnly -Scope Script
Set-Variable -Name htLoggingPreference -Option AllScope -Scope Script -Value @{"InformationPreference"=$InformationPreference; `
    "WarningPreference"=$WarningPreference;"ErrorActionPreference"=$ErrorActionPreference;"VerbosePreference"=$VerbosePreference;"DebugPreference"=$DebugPreference}
Set-Variable -Name verScript -Option AllScope -Scope Script -Value "5.1.2023.0808"

Set-Variable -Name boolScriptIsModulesLoaded -Option AllScope -Scope Script -Value $false
Set-Variable -Name ExitCode -Option AllScope -Scope Script -Value 1

New-Object System.Data.DataTable | Set-Variable dtContacts -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtAddresses -Option AllScope -Scope Script

New-Object System.Collections.ArrayList | Set-Variable arrExceptions -Option AllScope -Scope Script
#endregion

#region Complete Functions
Function _ConfirmScriptRequirements
{
    <#
    .SYNOPSIS
        Verifies that all necessary requirements are present for the script and return true/false
    .EXAMPLE
        $valid = _ConfirmScriptRequirements

        This would check the script requirements and set $valid to true/false based on the results
    .NOTES
        Version:
        - 5.1.2023.0808:    New function
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

        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Debug -WriteBackToHost -Message "Starting _ConfirmScriptRequirements"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Script version $verScript starting"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: InformationPreference = $InformationPreference"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ErrorActionPreference = $ErrorActionPreference"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: VerbosePreference = $VerbosePreference"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: DebugPreference = $DebugPreference"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: AddressesFilePath = $AddressesFilePath"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: CredentialsPath = $CredentialsPath"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: MappingAtrribute = $MappingAttribute"
    }#begin
    
    process{
        if ($script:boolScriptIsModulesLoaded) {
            try{
                $global:VerbosePreference = "SilentlyContinue"

                $ConnectSplat = @{
                    "ShowBanner" = $False
                }

                if ($CredentialsPath) {
                    $ConnectSplat.Add("Credential", $(Import-Clixml $CredentialsPath))
                }

                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Connecting to Exchange Online"
                Connect-ExchangeOnline @ConnectSplat
                
                if($htLoggingPreference['VerbosePreference'] -eq "Continue"){$global:VerbosePreference = "Continue"}#if                
            } catch {
                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error verifying script requirements"
                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
                return $false
            }#try/catch
        }#if

        #Final check
        if ($script:boolScriptIsModulesLoaded){return $true}
        else {return $false}
    }#process

    end {
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Debug -WriteBackToHost -Message "Finishing _ConfirmScriptRequirements"
    }#end
}#function _ConfirmScriptRequirements

function _GetScriptDirectory
{
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
        if($Leaf) {
            Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf
        } elseif($LeafBase) {
            (Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf).Split(".")[0]
        } elseif($Path) {
            Split-Path $hostinvocation.MyCommand.path
        } else {
            Split-Path $hostinvocation.MyCommand.path
        }#if/else
    } elseif ($null -ne $script:MyInvocation.MyCommand.Path) {
        if($Leaf) {
            Split-Path $script:MyInvocation.MyCommand.Path -Leaf
        } elseif($LeafBase) {
            (Split-Path $script:MyInvocation.MyCommand.Path -Leaf).Split(".")[0]
        } elseif($Path) {
            Split-Path $script:MyInvocation.MyCommand.Path
        } else {
            (Get-Location).Path + "\" + (Split-Path $script:MyInvocation.MyCommand.Path -Leaf)
        }#if/else
    } else {
        if($Leaf) {
            Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf
        } elseif($LeafBase) {
            (Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf).Split(".")[0]
        } elseif($Path) {
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
Write-Host "Exchange Contact Email Addresses Script`r"
Write-Host "`r"

Write-Host "Script starting`r"

if (_ConfirmScriptRequirements) {
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Script requirements met"

    $dtAddresses = Import-Csv $AddressesFilePath | ConvertTo-DataTable
    
    if ($dtAddresses.Rows.Count -ge 1) {
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "$($dtAddresses.Rows.Count) Addresses found"

        foreach($address in $dtAddresses.Rows) {
            try {
                $addressToAdd = $address.EmailAddresses -ireplace "smtp:", ""
                $newAddress = Get-MailContact -ResultSize Unlimited -Verbose:$false | Where-Object {$_.$MappingAttribute -eq $address.Guid} | Set-MailContact -EmailAddresses @{Add = $addressToAdd} -Verbose:$false -ErrorAction Stop
                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Added $($address.EmailAddresses) to $($address.Guid)"

                $ExitCode = 0
            } catch {
                $ErrorMessage = $_.Exception.Message
                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Failed to add $($address.EmailAddresses) to $($address.Guid)"
                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $ErrorMessage
                
                $objNew = New-Object -TypeName PSCustomObject -Property @{
                    "ContactDomain" = $address["ContactDomain"]
                    "EmailAddress" = $address["EmailAddresses"]
                    "SourceGuid" = $address["Guid"]
                }
                [void]$arrExceptions.Add($objNew)

                $ExitCode = 1
            }#try/catch
        }#foreach
    }#if

    
    if($arrExceptions.Count -ge 1){
        $ExportLocation = $script:strBaseLocation + "\Exchange"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting exception CSV to $ExportLocation"
        
        $arrExceptions | Export-Csv -Path "$ExportLocation\GroupMembership_Exceptions_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
    }#if

    $RunTime = ((get-date).ToUniversalTime() - $dateStartTimeStamp)
    $RunTime = '{0:00}:{1:00}:{2:00}:{3:00}.{4:00}' -f $RunTime.Days,$RunTime.Hours,$RunTime.Minutes,$RunTime.Seconds,$RunTime.Milliseconds
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Run time was $RunTime"
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exit code is $ExitCode"
} else {
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Script requirements not met:"

    if (-not $script:boolScriptIsModulesLoaded) {
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Missing required PowerShell module(s) or could not load modules"
    }#if
}#if/else

Get-ConnectionInformation -ErrorAction SilentlyContinue -Verbose:$fasle | Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
Exit $ExitCode
#endregion