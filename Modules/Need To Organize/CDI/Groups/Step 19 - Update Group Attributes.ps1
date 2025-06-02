#Requires -version 5.0
#Requires -Modules Sterling

#region Parameters
[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the group CSV file")]
    [string]$GroupFilePath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the credential XML file to be used")]
    [string]$CredentialsPath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify to update the primary and proxy addresses")]
    [switch]$IncludeEmailAddresses,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the group email addresses CSV file")]
    [string]$EmailAddressesFilePath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify to update Unified/M365 groups")]
    [switch]$IncludeUnifiedGroups,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the attribute to use to match the source group guid. It will default to customAttribute1 if not specified.")]
    [string]$MappingAttribute = "customAttribute1",
    
    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify to connect to GCC High tenants")]
    [switch]$GCCHigh
)
#endregion

#region User Variables
#Very little reason to change these
$InformationPreference = "Continue"

if ($DebugPreference -eq "Confirm" -or $DebugPreference -eq "Inquire") {$DebugPreference = "Continue"}
#endregion

#region Static Variables
#Don't change these
Set-Variable -Name strBaseLocation -WhatIf:$false -Option AllScope -Scope Script
Set-Variable -Name dateStartTimeStamp -WhatIf:$false -Option AllScope -Scope Script -Value (Get-Date).ToUniversalTime()
Set-Variable -Name strLogTimeStamp -WhatIf:$false -Option AllScope -Scope Script -Value $dateStartTimeStamp.ToString("MMddyyyy_HHmmss")
Set-Variable -Name strLogFile -WhatIf:$false -Option ReadOnly -Scope Script
Set-Variable -Name htLoggingPreference -WhatIf:$false -Option AllScope -Scope Script -Value @{"InformationPreference"=$InformationPreference; `
    "WarningPreference"=$WarningPreference;"ErrorActionPreference"=$ErrorActionPreference;"VerbosePreference"=$VerbosePreference; `
    "DebugPreference"=$DebugPreference;"WhatIfPreference"=$WhatIfPreference}
Set-Variable -Name verScript -WhatIf:$false -Option AllScope -Scope Script -Value "5.1.2024.0118"

Set-Variable -Name boolScriptIsModulesLoaded -WhatIf:$false -Option AllScope -Scope Script -Value $false
Set-Variable -Name ExitCode -WhatIf:$false -Option AllScope -Scope Script -Value 1

New-Object System.Data.DataTable | Set-Variable -Name dtGroups -WhatIf:$false -Option AllScope -Scope Script
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
        - 5.1.2023.1213:    New function
        - 5.1.2024.0118:    Added GCCHigh parameter
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
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: GroupFilePath = $GroupFilePath"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: CredentialsPath = $CredentialsPath"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludeEmailAddresses = $IncludeEmailAddresses"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: IncludeUnifiedGroups = $IncludeUnifiedGroups"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: GCCHigh = $GCCHigh"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: MappingAttribute = $MappingAttribute"
    }#begin
    
    process{
        if ($script:boolScriptIsModulesLoaded) {
            try{
                $global:VerbosePreference = "SilentlyContinue"

                $ConnectSplat = @{
                    "GCCHigh" = $GCCHigh
                }

                if ($CredentialsPath) {
                    $ConnectSplat.Add("Credential", $(Import-Clixml $CredentialsPath))
                }

                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Connecting to Exchange Online"
                Connect-Exchange @ConnectSplat
                
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
Write-Host "Exchange Distribution Group Update Script`r"
Write-Host "`r"

Write-Host "Script starting`r"

$WhatIfPreference = $false
if (_ConfirmScriptRequirements) {
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Script requirements met"

    if ($IncludeUnifiedGroups) {
        $dtGroups = Import-Csv $GroupFilePath | Where {$_.GroupType -notmatch "Dynamic"} | ConvertTo-DataTable
    } else {
        $dtGroups = Import-Csv $GroupFilePath | Where {$_.GroupType -notmatch "Dynamic" -and $_.RecipientTypeDetails -ne "GroupMailbox"} | ConvertTo-DataTable
    }

    if ($IncludeEmailAddresses) {
        $dtEmailAddresses = Import-Csv $EmailAddressesFilePath | Where {$_.EmailAddresses -cnotmatch "SMTP:"} | ConvertTo-DataTable
    }
    
    if ($dtGroups.Rows.Count -ge 1) {
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "$($dtGroups.Rows.Count) Groups imported"

        foreach($group in $dtGroups.Rows) {
            #Update group
            try {
                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Working on group: $($group.DisplayName)"
                
                #Update to use matching attribute
                $foundGroup = Get-Recipient -Filter "$MappingAttribute -like '$($group.Guid)'" -Verbose:$false -ErrorAction SilentlyContinue
                
                if (-not $foundGroup) {
                    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Group $($group.alias) not found, unable to update"
                } else {
                    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Found existing group"

                    $GroupSettingsSplat = @{
                        "DisplayName" = $group.DisplayName
                        "CustomAttribute1" = $group.CustomAttribute1
                        "CustomAttribute10" = $group.CustomAttribute10
                        "CustomAttribute11" = $group.CustomAttribute11
                        "CustomAttribute12" = $group.CustomAttribute12
                        "CustomAttribute13" = $group.CustomAttribute13
                        "CustomAttribute14" = $group.CustomAttribute14
                        "CustomAttribute15" = $group.CustomAttribute15
                        "CustomAttribute2" = $group.CustomAttribute2
                        "CustomAttribute3" = $group.CustomAttribute3
                        "CustomAttribute4" = $group.CustomAttribute4
                        "CustomAttribute5" = $group.CustomAttribute5
                        "CustomAttribute6" = $group.CustomAttribute6
                        "CustomAttribute7" = $group.CustomAttribute7
                        "CustomAttribute8" = $group.CustomAttribute8
                        "CustomAttribute9" = $group.CustomAttribute9
                        "HiddenFromAddressListsEnabled" = [System.Convert]::ToBoolean($group.HiddenFromAddressListsEnabled)
                    }

                    if (-not [string]::IsNullorEmpty($group.MailTip)) {$GroupSettingsSplat.Add("MailTip", $group.MailTip)}#if
                    if (-not [string]::IsNullorEmpty($group.MaxReceiveSize)) {$GroupSettingsSplat.Add("MaxReceiveSize", $group.MaxReceiveSize)}#if
                    if (-not [string]::IsNullorEmpty($group.MaxSendSize)) {$GroupSettingsSplat.Add("MaxSendSize", $group.MaxSendSize)}#if

                    if (-not [string]::IsNullorEmpty($group.Description) -and $group.RecipientTypeDetails -notmatch "GroupMailbox") {$GroupSettingsSplat.Add("Description", $group.Description)}#if
                    if (-not [string]::IsNullorEmpty($group.ReportToManagerEnabled) -and $group.RecipientTypeDetails -notmatch "GroupMailbox") {$GroupSettingsSplat.Add("ReportToManagerEnabled", [System.Convert]::ToBoolean($group.ReportToManagerEnabled))}#if
                    if (-not [string]::IsNullorEmpty($group.ReportToOriginatorEnabled) -and $group.RecipientTypeDetails -notmatch "GroupMailbox") {$GroupSettingsSplat.Add("ReportToOriginatorEnabled", [System.Convert]::ToBoolean($group.ReportToOriginatorEnabled))}#if
                    if (-not [string]::IsNullorEmpty($group.SendOofMessageToOriginatorEnabled) -and $group.RecipientTypeDetails -notmatch "GroupMailbox") {$GroupSettingsSplat.Add("SendOofMessageToOriginatorEnabled", [System.Convert]::ToBoolean($group.SendOofMessageToOriginatorEnabled))}#if

                    if (-not [string]::IsNullorEmpty($group.ExtensionCustomAttribute1)) {$GroupSettingsSplat.Add("ExtensionCustomAttribute1", $group.ExtensionCustomAttribute1)}#if
                    if (-not [string]::IsNullorEmpty($group.ExtensionCustomAttribute2)) {$GroupSettingsSplat.Add("ExtensionCustomAttribute2", $group.ExtensionCustomAttribute2)}#if
                    if (-not [string]::IsNullorEmpty($group.ExtensionCustomAttribute3)) {$GroupSettingsSplat.Add("ExtensionCustomAttribute3", $group.ExtensionCustomAttribute3)}#if
                    if (-not [string]::IsNullorEmpty($group.ExtensionCustomAttribute4)) {$GroupSettingsSplat.Add("ExtensionCustomAttribute4", $group.ExtensionCustomAttribute4)}#if
                    if (-not [string]::IsNullorEmpty($group.ExtensionCustomAttribute5)) {$GroupSettingsSplat.Add("ExtensionCustomAttribute5", $group.ExtensionCustomAttribute5)}#if

                    if(-not [string]::IsNullorEmpty($group.BccBlocked) -and $group.RecipientTypeDetails -notmatch "GroupMailbox"){$GroupSettingsSplat.Add("BccBlocked", [System.Convert]::ToBoolean($group.BccBlocked))}
                    if(-not [string]::IsNullorEmpty($group.BypassNestedModerationEnabled) -and $group.RecipientTypeDetails -notmatch "GroupMailbox"){$GroupSettingsSplat.Add("BypassNestedModerationEnabled", [System.Convert]::ToBoolean($group.BypassNestedModerationEnabled))}
                    if(-not [string]::IsNullorEmpty($group.ModerationEnabled) -and $group.RecipientTypeDetails -notmatch "GroupMailbox"){$GroupSettingsSplat.Add("ModerationEnabled", [System.Convert]::ToBoolean($group.ModerationEnabled))}
                    if(-not [string]::IsNullorEmpty($group.RequireSenderAuthenticationEnabled)){$GroupSettingsSplat.Add("RequireSenderAuthenticationEnabled", [System.Convert]::ToBoolean($group.RequireSenderAuthenticationEnabled))}
                    if(-not [string]::IsNullorEmpty($group.SendModerationNotifications) -and $group.RecipientTypeDetails -notmatch "GroupMailbox"){$GroupSettingsSplat.Add("SendModerationNotifications", $group.SendModerationNotifications)}
                    if(-not [string]::IsNullorEmpty($group.MemberDepartRestriction) -and $group.RecipientTypeDetails -notmatch "GroupMailbox"){$GroupSettingsSplat.Add("MemberDepartRestriction", $group.MemberDepartRestriction)}
                    if(-not [string]::IsNullorEmpty($group.MemberJoinRestriction) -and $group.RecipientTypeDetails -notmatch "GroupMailbox"){$GroupSettingsSplat.Add("MemberJoinRestriction", $group.MemberJoinRestriction)}
                    
                    #EmailAddresses
                    if ($IncludeEmailAddresses) {
                        $GroupAddresses = ($dtEmailAddresses.Select("GroupGuid = '$($group.Guid)'")).EmailAddresses
                    }#if
                    
                    #Remove the mapping attribute so we don't accidentally update it
                    $GroupSettingsSplat.Remove($MappingAttribute)

                    if ($group.RecipientTypeDetails -match "MailUniversal") {
                        $GroupSettingsSplat.Add("IgnoreNamingPolicy", $true)
                        $GroupSettingsSplat.Add("Name", $group.Name)
                        if(-not [string]::IsNullorEmpty($group.HiddenGroupMembershipEnabled)){$GroupSettingsSplat.Add("HiddenGroupMembershipEnabled", [System.Convert]::ToBoolean($group.HiddenGroupMembershipEnabled))}

                        if($htLoggingPreference['WhatIfPreference']){$WhatIfPreference = $true}#if
                        if ($PSCmdlet.ShouldProcess("Set-DistributionGroup with parameters: $($GroupSettingsSplat | Out-String)", "", "")) {
                            Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Updating group settings: $($GroupSettingsSplat | Out-String)"

                            Set-DistributionGroup -Identity $foundGroup.Identity @GroupSettingsSplat -Verbose:$false -ErrorAction Stop | Out-Null

                            if ($IncludeEmailAddresses) {
                                foreach ($address in $GroupAddresses) {Set-DistributionGroup -Identity $foundGroup.Identity -EmailAddresses @{Add=$address} -Verbose:$false -ErrorAction Stop | Out-Null}
                            }#if
                            Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Group updated"
                        }

                        $WhatIfPreference = $false
                    } elseif ($IncludeUnifiedGroups -and $group.RecipientTypeDetails -match "GroupMailbox") {
                        $GroupSettingsSplat.Add("Notes", $group.Notes)

                        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Updated M365 group: $($GroupSettingsSplat | Out-String)"

                        #For a really frustrating default with O365 where ErrorAction isn't an allowed parameter
                        #https://stackoverflow.com/questions/61473366/new-unifiedgroup-doesnt-work-with-erroraction#:~:text=The%20%22ErrorAction%22%20parameter%20can',to%20you%2C%20and%20try%20again.
                        $global:ErrorActionPreference = "Stop"
                        if($htLoggingPreference['WhatIfPreference']){$WhatIfPreference = $true}#if
                        if ($PSCmdlet.ShouldProcess("Set-UnifiedGroup with parameters: $($GroupSettingsSplat | Out-String)", "", "")) {
                            Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Updating M365 group settings: $($GroupSettingsSplat | Out-String)"

                            Set-UnifiedGroup -Identity $foundGroup.Identity @GroupSettingsSplat -Verbose:$false -ErrorAction Stop | Out-Null

                            if ($IncludeEmailAddresses) {
                                foreach ($address in $GroupAddresses) {Set-UnifiedGroup -Identity $foundGroup.Identity -EmailAddresses @{Add=$address} -Verbose:$false -ErrorAction Stop | Out-Null}
                            }#if
                            Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "M365 Group updated"
                        }
                        
                        if($htLoggingPreference['ErrorActionPreference'] -ne "Stop"){$global:ErrorActionPreference = $htLoggingPreference['ErrorActionPreference']}#if
                        $WhatIfPreference = $false
                    }#if/elseif
                }#if/else

                $ExitCode = 0
            } catch {
                $ErrorMessage = $_.Exception.Message
                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Failed to update group"
                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $ErrorMessage

                $ExitCode = 1
            }#try/catch
        }#foreach
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