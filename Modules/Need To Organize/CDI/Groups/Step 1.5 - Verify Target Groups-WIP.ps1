#Requires -version 5.0
#Requires -Modules Sterling
#Requires -Modules AzureADPreview

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
        HelpMessage = "Specify to connect to GCC High tenants")]
    [switch]$GCCHigh,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify to verify Unified/M365 groups")]
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
Set-Variable -Name verScript -WhatIf:$false -Option AllScope -Scope Script -Value "5.1.2024.0122"

Set-Variable -Name boolScriptIsModulesLoaded -WhatIf:$false -Option AllScope -Scope Script -Value $false
Set-Variable -Name ExitCode -WhatIf:$false -Option AllScope -Scope Script -Value 1

New-Object System.Data.DataTable | Set-Variable -Name dtGroups -WhatIf:$false -Option AllScope -Scope Script
New-Object System.Collections.ArrayList | Set-Variable arrExceptions -WhatIf:$false -Option AllScope -Scope Script
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
        - 5.1.2024.0122:    New function
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
            Write-Host "Loading Azure AD Preview PowerShell module`r"

            if (Get-Module -ListAvailable AzureADPreview -Verbose:$false) {
                Import-Module AzureADPreview -ErrorAction Stop -Verbose:$false
                $script:boolScriptIsModulesLoaded = $true
            } else {
                Write-Warning "Missing Azure AD Preview PowerShell module`r"
                $script:boolScriptIsModulesLoaded = $false
            }#if/else
        } catch {
            Write-Error "Unable to load Azure AD Preview PowerShell module`r"
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
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: GroupFilePath = $GroupFilePath"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: CredentialsPath = $CredentialsPath"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: Prefix = $Prefix"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: Suffix = $Suffix"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: GCCHigh = $GCCHigh"
    }#begin
    
    process{
        if ($script:boolScriptIsModulesLoaded) {
            try{
                $global:VerbosePreference = "SilentlyContinue"

                $ConnectSplat = @{
                    "AzureEnvironmentName" = "AzureCloud"
                }

                if($GCCHigh) {
                    $ConnectSplat["AzureEnvironmentName"] = "AzureUSGovernment"
                }

                if ($CredentialsPath) {
                    $ConnectSplat.Add("Credential", $(Import-Clixml $CredentialsPath))
                }

                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Connecting to Entra ID"
                Connect-AzureAD @ConnectSplat
                
                if($htLoggingPreference['VerbosePreference'] -eq "Continue"){$global:VerbosePreference = "Continue"}#if
            } catch {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error verifying script requirements"
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
                return $false
            }#try/catch
        }#if

        #Final check
        if ($script:boolScriptIsModulesLoaded){return $true}
        else {return $false}
    }#process

    end {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Debug -WriteBackToHost -Message "Finishing _ConfirmScriptRequirements"
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
    #>
    [CmdletBinding()]
    [OutputType([System.Void])]
    Param( 
        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the file path for the log file. It will default to %temp%\Sterling.log if not specified.")]
        [string]$Logfile = "$env:TEMP\Sterling.log",
        
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the object or string for the log")]
        $Message,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the severity of message to log.")]
        [ValidateSet('Warning','Error','Verbose','Debug', 'Information')] 
        [string]$Type = "Information",

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify whether or not to write the message back to the console. It will default to false if not specified.")]
        [switch]$WriteBackToHost,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify a hashtable of the preference variables so that they can be honored. It will use the default PowerShell values if not specified.")]
        [hashtable]$LoggingPreference = @{"InformationPreference"=$InformationPreference; `
            "WarningPreference"=$WarningPreference;"ErrorActionPreference"=$ErrorActionPreference;"VerbosePreference"=$VerbosePreference;"DebugPreference"=$DebugPreference},

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify whether or not to force the message to be written to the log. It will default to false if not specified.")]
        [switch]$ForceWriteToLog
    )#Param

    begin {
        #Update variables based on parameters
        $Type = $Type.ToUpper()
        
        #Set the order 
        switch($Type){
            'Warning' {$severity = 2}#Warning
            'Error' {$severity = 3}#Error
            'Verbose' {$severity = 4}#Verbose
            'Debug' {$severity = 5}#Debug
            'Information' {$severity = 6}#Information
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
        if($null -eq $global:ErrorLogStream -or -not $global:ErrorLogStream.BaseStream) {
            If(-not (Test-Path -Path (Split-Path $LogFile))) {New-Item (Split-Path $LogFile) -ItemType Directory | Out-Null}
            
            $global:ErrorLogStream = New-Object System.IO.StreamWriter $Logfile, $true, ([System.Text.Encoding]::UTF8)
            $global:ErrorLogStream.AutoFlush = $true
        }

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
    }

    process {
        #Switch statement to write out to the log and/or back to the host.
        switch ($severity){
            2{
                if ($LoggingPreference['WarningPreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                    #Build log message
                    $LogMessage = $Type + ": " + $Message
                    $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                        "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
                    #$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                    #    "$(_GetScriptDirectory -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $(_GetScriptDirectory)

                    $logline = $LineTemplate -f $LineContent
                    $global:ErrorLogStream.WriteLine($logline)
                }#if silentlycontinue and not forced to write, don't write to the log
                
                #Write back to the host if $Writebacktohost is true.
                if($WriteBackToHost){
                    switch($LoggingPreference['WarningPreference']){
                        'Continue' {$WarningPreference = 'Continue';Write-Warning -Message "$Message";$WarningPreference=''}
                        'Stop' {$WarningPreference = 'Stop';Write-Warning -Message "$Message";$WarningPreference=''}
                        'Inquire' {$WarningPreference ='Inquire';Write-Warning -Message "$Message";$WarningPreference=''}
                    }#switch
                }#if writeback
            }#Warning
            3{  
                #This if statement is to catch the two different types of errors that may come through. 
                #A normal terminating exception will have all the information that is needed, if it's a user generated error by using Write-Error,
                #then the else statment will setup all the information we would like to log.   
                if ($Message.Exception.Message){
                    if ($LoggingPreference['ErrorActionPreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                        #Build log message                                      
                        $LogMessage = $Type + ": " + [string]$Message.Exception.Message + " Command: '" + [string]$Message.InvocationInfo.MyCommand +`
                            "' Line: '" + [string]$Message.InvocationInfo.Line.Trim() + "'"
                        $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                            "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
                        #$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                        #    "$(_GetScriptDirectory -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $(_GetScriptDirectory)

                        $logline = $LineTemplate -f $LineContent
                        $global:ErrorLogStream.WriteLine($logline)
                    }#if silentlycontinue and not forced to write, don't write to the log

                    #Write back to the host if $Writebacktohost is true.
                    if($WriteBackToHost){
                        switch($LoggingPreference['ErrorActionPreference']){
                            'Stop'{$ErrorActionPreference = 'Stop';$Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)");Write-Error $Message -ErrorAction 'Stop';$ErrorActionPreference=''}
                            'Inquire'{$ErrorActionPreference = 'Inquire';$Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)");Write-Error $Message -ErrorAction 'Inquire';$ErrorActionPreference=''}
                            'Continue'{$ErrorActionPreference = 'Continue';$Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)");$ErrorActionPreference=''}
                            'Suspend'{$ErrorActionPreference = 'Suspend';$Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)");Write-Error $Message -ErrorAction 'Suspend';$ErrorActionPreference=''}
                        }#switch
                    }#if writeback
                }#if standard error
                else{
                    if ($LoggingPreference['ErrorActionPreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                        #Custom error message so build out the Exception object
                        [System.Exception]$Exception = $Message
                        [String]$ErrorID = 'Custom Error'
                        [System.Management.Automation.ErrorCategory]$ErrorCategory = [Management.Automation.ErrorCategory]::WriteError
                        $ErrorRecord = New-Object Management.automation.errorrecord ($Exception,$ErrorID,$ErrorCategory,$Message)
                        $Message = $ErrorRecord

                        #Build log message                
                        $LogMessage = $Type + ": " + [string]$Message.Exception.Message
                        $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                            "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
                        #$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                        #    "$(_GetScriptDirectory -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $(_GetScriptDirectory)

                        $logline = $LineTemplate -f $LineContent
                        $global:ErrorLogStream.WriteLine($logline)
                    }#if silentlycontinue and not forced to write, don't write to the log
                        
                    #Write back to the host if $Writebacktohost is true.
                    if($WriteBackToHost){
                        switch($LoggingPreference['ErrorActionPreference']){
                            'Stop'{$ErrorActionPreference = 'Stop';$Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)");Write-Error $Message -ErrorAction 'Stop';$ErrorActionPreference=''}
                            'Inquire'{$ErrorActionPreference = 'Inquire';$Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)");Write-Error $Message -ErrorAction 'Inquire';$ErrorActionPreference=''}
                            'Continue'{$ErrorActionPreference = 'Continue';$Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)");$ErrorActionPreference=''}
                            'Suspend'{$ErrorActionPreference = 'Suspend';$Host.Ui.WriteErrorLine("ERROR: $([String]$Message.Exception.Message)");Write-Error $Message -ErrorAction 'Suspend';$ErrorActionPreference=''}
                        }#switch
                    }#if writeback
                }#else custom error
            }#Error
            4{  
                if ($LoggingPreference['VerbosePreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                    #Build log message                
                    $LogMessage = $Type + ": " + $Message
                    $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                        "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
                    #$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                    #    "$(_GetScriptDirectory -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $(_GetScriptDirectory)
                    
                    $logline = $LineTemplate -f $LineContent
                    $global:ErrorLogStream.WriteLine($logline)
                }#if silentlycontinue and not forced to write, don't write to the log

                #Write back to the host if $Writebacktohost is true.
                if($WriteBackToHost){
                    switch($LoggingPreference['VerbosePreference']){
                        'Continue' {$VerbosePreference = 'Continue'; Write-Verbose -Message "$Message";$VerbosePreference = ''}
                        'Inquire' {$VerbosePreference = 'Inquire'; Write-Verbose -Message "$Message";$VerbosePreference = ''}
                        'Stop' {$VerbosePreference = 'Stop'; Write-Verbose -Message "$Message";$VerbosePreference = ''}
                    }#switch
                }#if writeback
            }#Verbose
            5{  
                if ($LoggingPreference['DebugPreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                    #Build log message                
                    $LogMessage = $Type + ": " + $Message
                    $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                        "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
                    #$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                    #    "$(_GetScriptDirectory -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $(_GetScriptDirectory)
                    
                        $logline = $LineTemplate -f $LineContent
                    $global:ErrorLogStream.WriteLine($logline)
                }#if silentlycontinue and not forced to write, don't write to the log
                
                #Write back to the host if $Writebacktohost is true.
                if($WriteBackToHost){
                    switch($LoggingPreference['DebugPreference']){
                        'Continue' {$DebugPreference = 'Continue'; Write-Debug -Message "$Message";$DebugPreference = ''}
                        'Inquire' {$DebugPreference = 'Inquire'; Write-Debug -Message "$Message";$DebugPreference = ''}
                        'Stop' {$DebugPreference = 'Stop'; Write-Debug -Message "$Message";$DebugPreference = ''}
                    }#switch
                }#if writeback
            }#Debug
            6{  
                if ($LoggingPreference['InformationPreference'] -ne 'SilentlyContinue' -or $ForceWriteToLog) {
                    #Build log message                
                    $LogMessage = $Type + ": " + $Message
                    $LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                        "$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
                    #$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
                    #    "$(_GetScriptDirectory -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $(_GetScriptDirectory)
                    
                    $logline = $LineTemplate -f $LineContent
                    $global:ErrorLogStream.WriteLine($logline)
                }#if silentlycontinue and not forced to write, don't write to the log

                #Write back to the host if $Writebacktohost is true.
                if($WriteBackToHost){
                    switch($LoggingPreference['InformationPreference']){
                        'Continue' {$InformationPreference = 'Continue'; Write-Information -Message "INFORMATION: $Message";$InformationPreference = ''}
                        'Inquire' {$InformationPreference = 'Inquire'; Write-Information -Message "INFORMATION: $Message";$InformationPreference = ''}
                        'Stop' {$InformationPreference = 'Stop'; Write-Information -Message "INFORMATION: $Message";$InformationPreference = ''}
                        'Suspend' {$InformationPreference = 'Suspend';Write-Information -Message "INFORMATION: $Message";$InformationPreference = ''}
                    }#switch
                }#if writeback
            }#Information
        }#Switch
    }#process

    end{
        #Close log files while we are waiting
        if($null -ne $global:ErrorLogStream) {
            $global:ErrorLogStream.Close()
            $global:ErrorLogStream.Dispose()
        }
    }#end
}#Function Out-Log

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
        Function _gettype 
        {
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
            
            If ($types -contains $type) {return $type}
            ElseIf ($type -match "System.Collections.Generic.List"){return "System.Array"}
            ElseIf ($type -match "System.Collections.ArrayList"){return "System.Collections.ArrayList"}
            ElseIf ($type -match "MultiValuedProperty"){return "System.Collections.ArrayList"}
            Else {return "System.String"}
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
                    $Property.Value | %{$Value += $_}
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

                if ($null -eq $Value) {$Value = [DBNull]::Value}
                
                if ($Value.getType().ToString() -eq "System.Collections.ArrayList") {
                    $NewDataTableRow.Item($Name) = [System.Collections.ArrayList]$Value
                } else {$NewDataTableRow.Item($Name) = $Value}
            }#foreach property
            
            [void]$NewDatatable.Rows.Add($NewDataTableRow)
            
            $First = $false
        }#foreach row
    }#process

    end {
        # Because PowerShell handles returning objects stupidly
        return @(,$NewDatatable)
    }#end
}#Function ConvertTo-DataTable
#endregion

#region Main Program
Write-Host "`r"
Write-Host "Script Written by Sterling Consulting`r"
Write-Host "All rights reserved. Proprietary and Confidential Material`r"
Write-Host "Exchange Distribution Group Target Verification Script`r"
Write-Host "`r"

Write-Host "Script starting`r"

$WhatIfPreference = $false
if (_ConfirmScriptRequirements) {
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Script requirements met"

    if ($IncludeUnifiedGroups) {
        $dtGroups = Import-Csv $GroupFilePath | Where {$_.GroupType -notmatch "Dynamic"} | ConvertTo-DataTable
    } else {
        $dtGroups = Import-Csv $GroupFilePath | Where {$_.GroupType -notmatch "Dynamic" -and $_.RecipientTypeDetails -ne "GroupMailbox"} | ConvertTo-DataTable
    }
    
    if ($dtGroups.Rows.Count -ge 1) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "$($dtGroups.Rows.Count) Groups imported"

        foreach($group in $dtGroups.Rows) {
            $verifyGroup = $null
            
            #Create group
            try {
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Working on group: $($group.DisplayName)"
                
                $VerifyGroupSplat = @{
                    "SearchString" = $Prefix+$group.DisplayName+$Suffix
                    #Could be used for on-premises checking
                    #"Filter" = "DisplayName -eq '$Prefix$($group.DisplayName)$Suffix' -or Name -eq '$Prefix$($group.DisplayName)$Suffix'"
                }

                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Looking for: $($VerifyGroupSplat.SearchString)"
                $verifyGroup = Get-AzureADGroup @VerifyGroupSplat -Verbose:$false -ErrorAction SilentlyContinue
                #Could be used for on-premises checking
                #$verifyGroup = Get-ADGroup @VerifyGroupSplat -Verbose:$false -ErrorAction SilentlyContinue
                if ($verifyGroup) {
                    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Warning -WriteBackToHost -Message "Group $($group.DisplayName) found"
                    [void]$arrExceptions.Add($group)
                }#if

                $ExitCode = 0
            } catch {
                $ErrorMessage = $_.Exception.Message
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Failed to find group"
                Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $ErrorMessage
                [void]$arrExceptions.Add($group)

                $ExitCode = 1
            }#try/catch

            
        }#foreach
    }#if

    if($arrExceptions.Count -ge 1){
        $ExportLocation = $script:strBaseLocation + "\Exchange"
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting exception CSV to $ExportLocation"

        #Check for path/folder
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Checking for $ExportLocation"
        if (-not (Test-Path -Path $ExportLocation)) {
            Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Creating folder structure for $ExportLocation"
            New-Item -Path $ExportLocation -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        }
        
        $arrExceptions | Export-Csv -Path "$ExportLocation\GroupVerifyTarget_Exceptions_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
    }#if

    $RunTime = ((get-date).ToUniversalTime() - $dateStartTimeStamp)
    $RunTime = '{0:00}:{1:00}:{2:00}:{3:00}.{4:00}' -f $RunTime.Days,$RunTime.Hours,$RunTime.Minutes,$RunTime.Seconds,$RunTime.Milliseconds
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Run time was $RunTime"
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exit code is $ExitCode"
} else {
    Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Script requirements not met:"

    if (-not $script:boolScriptIsModulesLoaded) {
        Out-Log -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Missing required PowerShell module(s) or could not load modules"
    }#if
}#if/else

Disconnect-AzureAD -Confirm:$false -Verbose:$false
Exit $ExitCode
#endregion