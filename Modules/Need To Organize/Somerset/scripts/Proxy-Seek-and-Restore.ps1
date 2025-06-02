<#
.SYNOPSIS
    Taking the output of the Proxy-Seek-and-Destroy script and restoring those email addresses back onto different objects
.DESCRIPTION
    Taking the output of the Proxy-Seek-and-Destroy script and restoring those email addresses back onto different objects

.PARAMETER StartDirectory
This parameter defines the default location when asking for the input csv file, which is the output from Proxy-Seek-and-Destroy.ps1
Example value could be "C:\Temp" or "E:\Temp Directory"

.PARAMETER RestoreToLocation
This parameter has two options "On-Premises" and "ExchangeOnline". This parameter is used to decide where the object types are located.

.PARAMETER ObjectType
This parameter has three options "Mailboxes", "Groups" and "All". This parameter is used to decide what object types are in scope. 
If Type equals "On-Premises" Unified Groups are automatically skipped, since that is not an option.
Default Value: "All"

.PARAMETER IgnoreSIP
This parameter is a switch meaning if added or set to true if takes effect. If added we will not restore SIP addresses

.PARAMETER IgnoreModernGroups
This parameter is a switch meaning if added or set to true if takes effect. If added we will not restore addresses onto Modern Groups

.PARAMETER IgnoreGroupMailboxes
This parameter is a switch meaning if added or set to true if takes effect. If added we will not restore addresses onto GroupMailboxes

.PARAMETER ConvertLegDNtoX500
This parameter is a switch meaning if added or set to true if takes effect. If added we will convert the LegacyExchangeDN to a x500 and apply the addresses

.PARAMETER OrgMgmtAdmin
This parameter defines the location and xml file name to the encrypted credentials used for Exchange on-premises access, if none are provided Kerberos is tried.
The credentials must be encrypted using CliXML
Example value: "C:\Temp\orgmgmt.xml" or "E:\Temp Directory\orgmgmt.xml"

.PARAMETER ExchangeServer
This parameter defines the fully qualified server name of the on-premises Exchange server.
Example value: server1.contoso.com

.PARAMETER O365Admin
This parameter defines the location and xml file name to the encrypted credentials required for Office 365 access.
The credentials must be encrypted using CliXML
Example value: "C:\Temp\o365admin.xml" or "E:\Temp Directory\o365admin.xml"

.PARAMETER DisableReportMode
This parameter is a switch meaning if added or set to true if takes effect. If added the script will follow all logic but make zero changes

.EXAMPLE
Proxy-Seek-and-Restore.ps1 -StartDirectory C:\temp -Type On-Premises -ExchangeServer server.domain.com -OrgMgmtAdmin C:\temp\creds.xml

.EXAMPLE
Proxy-Seek-and-Restore.ps1 -StartDirectory C:\temp -Type On-Premises -ObjectType Mailboxes -ExchangeServer server.domain.com -OrgMgmtAdmin C:\temp\creds.xml

.EXAMPLE
Proxy-Seek-and-Restore.ps1 -StartDirectory C:\temp -Type On-Premises -ObjectType Groups -ExchangeServer server.domain.com -OrgMgmtAdmin C:\temp\creds.xml

.EXAMPLE
Proxy-Seek-and-Restore.ps1 -StartDirectory C:\temp -Type On-Premises -ObjectType Groups -ExchangeServer server.domain.com -OrgMgmtAdmin C:\temp\creds.xml -DisableReportMode

.EXAMPLE
Proxy-Seek-and-Restore.ps1 -StartDirectory C:\temp -Type Office365 -O365Admin C:\temp\O365Creds.xml

.EXAMPLE
Proxy-Seek-and-Restore.ps1 -StartDirectory C:\temp -Type Office365 -ObjectType Mailboxes -O365Admin C:\temp\O365Creds.xml

.EXAMPLE
Proxy-Seek-and-Restore.ps1 -StartDirectory C:\temp -Type Office365 -ObjectType Groups -O365Admin C:\temp\O365Creds.xml

.EXAMPLE
Proxy-Seek-and-Restore.ps1 -StartDirectory C:\temp -Type Office365 -O365Admin C:\temp\O365Creds.xml -DisableReportMode

#>

#region Parameters
Param(
	[Parameter(Position = 1, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Specify the starting location for picking a CSV (ex: C:\Temp)")]
	[ValidateNotNullOrEmpty()]
	[String]$StartDirectory,

	[Parameter(Position = 2, Mandatory = $true,
		HelpMessage = "Specify the location to restore to On-Premises or ExchangeOnline")]
	[ValidateSet ("On-Premises", "ExchangeOnline")]
	[String]$RestoreToLocation,

	[Parameter(Position = 3, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Select the type of objects you want to restore proxy addresses")]
	[ValidateSet ("Mailboxes", "Groups", "All")]
	[String]$ObjectType = "All",

	[Parameter(Position = 4, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Do you want to ignore the SIP proxy address")]
	[Switch]$IgnoreSIP,

	[Parameter(Position = 5, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Do you want to ignore Group Mailboxes")]
	[Switch]$IgnoreGroupMailboxes,

	[Parameter(Position = 6, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Do you want to convert the LegacyExchangeDN to an x500: proxy address")]
	[Switch]$ConvertLegDNtoX500,

	[Parameter(Position = 7, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Do you want to convert all addresses to proxy address")]
	[Switch]$ConvertToAlias,

	[Parameter(Position = 8, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Specify the path to the encrypted creds (ex: C:\Temp\OrgAdmin.xml)")]
	[ValidateNotNullOrEmpty()]
	[String]$OrgMgmtAdmin,

	[Parameter(Position = 9, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Specify the FQDN of the Exchange Server on-premises (ex: server1.contoso.com)")]
	[ValidateNotNullOrEmpty()]
	[String]$ExchangeServer,

	[Parameter(Position = 10, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Specify the path to the encrypted creds (ex: C:\Temp\O365Admin.xml)")]
	[ValidateNotNullOrEmpty()]
	[String]$O365Admin = ($env:userprofile + "\Documents\WindowsPowershell\pswd\dbiga.xml"),

	[Parameter(Position = 11, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Disable Report Mode")]
	[Switch]$DisableReportMode
)
#endregion

#region Setup Logging
$InformationPreference = "Continue"
If ($DebugPreference -eq "Confirm" -or $DebugPreference -eq "Inquire") { $DebugPreference = "Continue" }
Set-Variable -Name strBaseLocation -Option AllScope -Scope Script -Value (Split-Path $script:MyInvocation.MyCommand.Path)
Set-Variable -Name dateStartTimeStamp -Option AllScope -Scope Script -Value (Get-Date).ToUniversalTime()
Set-Variable -Name strLogTimeStamp -Option AllScope -Scope Script -Value $dateStartTimeStamp.ToString("MMddyyyy_HHmmss")
Set-Variable -Name strTranscriptFile -Option AllScope -Scope Script -Value "$strBaseLocation\Logging\$strLogTimeStamp-$((Split-Path $script:MyInvocation.MyCommand.Path -Leaf).Replace(".ps1",''))-Transcript.log"
Set-Variable -Name logFileName -Option AllScope -Scope Script -Value "$strBaseLocation\Logging\$strLogTimeStamp-$((Split-Path $script:MyInvocation.MyCommand.Path -Leaf).Replace(".ps1",'')).log"
Set-Variable -Name htLoggingPreference -Option AllScope -Scope Script -Value @{"InformationPreference" = $InformationPreference; `
		"WarningPreference" = $WarningPreference; "ErrorActionPreference" = $ErrorActionPreference; "VerbosePreference" = $VerbosePreference; "DebugPreference" = $DebugPreference
}
#endregion

#region Functions
Function Get-FileName($StartDirectory) {
	# Prompts the user for the input file starting in $StartDirectory
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$openFileDialog.initialDirectory = $StartDirectory
	$openFileDialog.filter = "All files (*.*)| *.*"
	$openFileDialog.ShowDialog() | Out-Null
	$openFileDialog.filename
}
Function Connect-Exchange {
	Param(
		[Parameter(Position = 1, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[String] $ExchangeServer
	)
	$SessionOptions = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -OpenTimeout 20000
	
	# Attempt #1 https
	Try {
		Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Trying to connect to https://" + $ExchangeServer + "/PowerShell" + " using provided credentials") -Type Information -WriteBackToHost 
		$session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("https://" + $ExchangeServer + "/PowerShell") -Credential $ExchangeCred -AllowRedirection -SessionOption $SessionOptions -ErrorAction Stop
		Import-PSSession -Session $session -AllowClobber
		Set-ADServerSettings -ViewEntireForest:$true
		Return $session
	} Catch {
		# Attempt #2 http
		Try {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Trying to connect to http://" + $ExchangeServer + "/PowerShell" + " using provided credentials") -Type Information -WriteBackToHost 
			$session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell") -Credential $ExchangeCred -AllowRedirection -SessionOption $SessionOptions -ErrorAction Stop
			Import-PSSession -Session $session -AllowClobber
			Set-ADServerSettings -ViewEntireForest:$true
			Return $session
		} Catch {
			# Attempt #3 https and current logged on user
			Try {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Trying to connect to https://" + $ExchangeServer + "/PowerShell" + " using current logged on creds") -Type Information -WriteBackToHost 
				$session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("https://" + $ExchangeServer + "/PowerShell") -Authentication Kerberos -AllowRedirection -SessionOption $SessionOptions -ErrorAction Stop
				Import-PSSession -Session $session -AllowClobber
				Set-ADServerSettings -ViewEntireForest:$true
				Return $session
			} Catch {
				# Attempt #4 http and current logged on user
				Try {
					Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Trying to connect to http://" + $ExchangeServer + "/PowerShell" + " using current logged on creds") -Type Information -WriteBackToHost 
					$session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell") -Authentication Kerberos -AllowRedirection -SessionOption $SessionOptions -ErrorAction Stop
					Import-PSSession -Session $session -AllowClobber
					Set-ADServerSettings -ViewEntireForest:$true
					Return $session
				} Catch {
					Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("We have tried four different ways to connect to Exchnge and all have failed. Please ensure Exchange server is correct and PSRemoting is enabled") -Type Error -WriteBackToHost
				}
			}
		}
	}
}
Function Connect-ExOnline {
	Param(
		[Parameter(Position = 1, Mandatory = $false)]
		[Switch] $MFAEnabled
	)

	If ($MFAEnabled) {
		Import-Module ExchangeOnlineManagement
		Connect-ExchangeOnline -UserPrincipalName $ExchangeCred.UserName -ShowBanner:$false
	} Else {
		Import-Module ExchangeOnlineManagement
		Connect-ExchangeOnline -Credential $ExchangeCred -ShowBanner:$false -Verbose:$false
	}
}

Function ConvertTo-X500 {
	[CmdletBinding()]
	Param(
		[Parameter(Position = 1, Mandatory = $True,
			HelpMessage = "Specify Legacy DN wich needs to converted to X500.")]
		[Alias("DN", "LD", "Legacy")]
		[String] $LegacyDN
	)
    
	ForEach ($X500 in $LegacyDN) {
		<# Replace any underscore character (_) with a slash character (/).
	   Replace "+20" with a blank space.
	   Replace "+28" with an opening parenthesis character.
	   Replace "+29" with a closing parenthesis character.
	   Delete the "IMCEAEX-" string.
	   Delete the "@mgd.domain.com" string.
	   Add "X500:" at the beginning.
	   After you make these changes, the proxy address for the example in the "Symptoms" section resembles the following:
	   X500:/O=MMS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=User-addd-4b03-95f5-b9c9a421957358d
	#>
		$X500 = $X500.Replace("_", "/")
		$X500 = $X500.Replace("+20", " ")
		$X500 = $X500.Replace("IMCEAEX-", "")
		$X500 = $X500.Replace("+28", "(")
		$X500 = $X500.Replace("+29", ")")
		$X500 = $X500.Replace("2E", ".")
		$X500 = $X500.Replace("5F", "_")
		$x500 = $x500.Split("@")[0]
	}
	# Return value
	Return $x500 = "X500:" + $x500   
}

Function Out-CMTraceLog {
	<# 
	.SYNOPSIS 
		Write to a log file in a format that takes advantage of the CMTrace.exe log viewer that comes with SCCM.
		
	.DESCRIPTION 
		Output strings to a log file that is formatted for use with CMTRace.exe and also writes back to the host.
		
		The severity of the logged line can be set as: 
		
				2-Error
				3-Warning
				4-Verbose
				5-Debug
				6-Information

		Warnings will be highlighted in yellow. Errors are highlighted in red. 
		
		The tools to view the log: 
		SMS Trace - http://www.microsoft.com/en-us/download/details.aspx?id=18153 
		CM Trace - https://www.microsoft.com/en-us/download/details.aspx?id=50012 or the Installation directory on Configuration Manager 2012 Site Server - <Install Directory>\tools\ 
		
	.EXAMPLE 
		Try {
			Get-Process -Name DoesnotExist -ea stop
		}
		Catch {
			Out-CMTraceLog -Logfile "C:\output\logfile.log -Message  $_ -Type Error
		}
		
		This will write a line to the logfile.log file in c:\output\logfile.log. It will state the errordetails in the log file 
		and highlight the line in Red. It will also write back to the host in a friendlier red on black message than
		the normal error record.
		
	.EXAMPLE
		$VerbosePreference = Continue
		Out-CMTraceLog -Message  "This is a verbose message." -Type Verbose -VerbosePreference $VerbosePreference

		This example will write a verbose entry into the log file and also write back to the host. The Out-CMTraceLog will obey
		the preference variables.

	.EXAMPLE
		Out-CMTraceLog -Message  "This is an informational message" -WritebacktoHost:$false

		This example will write the informational message to the log but not back to the host.

	.EXAMPLE
		Function Test{
			[CmdletBinding()]
			Param()
			Out-CMTraceLog -VerbosePreference $VerbosePreference -Message  "This is a verbose message" -Type Verbose
		}
		Test -Verbose

		This example shows how to use Out-CMTraceLog inside a function and then call the function with the -verbose switch.
		The Out-CMTraceLog function will then print the verbose message.

	.NOTES
		Version:
			- 5.0.2020.0417:	Initial version. Adopted from
									https://wolffhaven.gitlab.io/wolffhaven_icarus_test/powershell/write-cmtracelog-dropping-logs-like-a-boss/
									https://adamtheautomator.com/building-logs-for-cmtrace-powershell/
			- 5.0.2020.0422:    Updated parameters to default more of them
	#> 

	#Define and validate parameters 
	[CmdletBinding()] 
	Param( 
	
		#Path to the log file 
		[Parameter(Mandatory = $false)]      
		[string]$Logfile = "C:\temp\CMTrace.log",
				
		#The information to log 
		[Parameter(Mandatory = $true)] 
		$message,
		
		#The severity (Error, Warning, Verbose, Debug, Information)
		[Parameter(Mandatory = $false)]
		[ValidateSet('Warning', 'Error', 'Verbose', 'Debug', 'Information')] 
		[string]$Type = "Information",
	
		#Write back to the console or just to the log file. By default it will write back to the host.
		[Parameter(Mandatory = $false)]
		[switch] $WriteBackToHost = $true,
	
		[Parameter(Mandatory = $false)]
		[hashtable]$LoggingPreference = @{"InformationPreference" = $InformationPreference; `
				"WarningPreference" = $WarningPreference; "ErrorActionPreference" = $ErrorActionPreference; "VerbosePreference" = $VerbosePreference; "DebugPreference" = $DebugPreference
		}
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
	
		#open log stream
		if ($null -eq $global:ErrorLogStream -or -not $global:ErrorLogStream.BaseStream) {
			If (-not (Test-Path -Path (Split-Path $LogFile))) { New-Item (Split-Path $LogFile) -ItemType Directory | Out-Null }
				
			$global:ErrorLogStream = New-Object System.IO.StreamWriter $Logfile, $true, ([System.Text.Encoding]::UTF8)
			$global:ErrorLogStream.AutoFlush = $true
		}
	
		#Need the callstack information to get the details about the calling script
		$CallStack = Get-PSCallStack | Select-Object -Property *
		if (($null -ne $CallStack.Count) -or (($CallStack.Command -ne '<ScriptBlock>') -and ($CallStack.Location -ne '<No file>') -and ($null -ne $CallStack.ScriptName))) {
			if ($CallStack.Count -eq 1) {
				$CallingInfo = $CallStack[0]
			} else {
				$CallingInfo = $CallStack[($CallStack.Count - 2)]
			}#need only or the second to the last one if multiple returned
		} else {
			Write-Error -Message 'No callstack detected' -Category 'InvalidData'
		}#if callstack info found
	}
	
	process {
		#Switch statement to write out to the log and/or back to the host.
		switch ($severity) {
			2 {     
				#Build log message
				$LogMessage = $Type + ": " + $message
				$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
					"$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
	
				$logline = $LineTemplate -f $LineContent
				$global:ErrorLogStream.WriteLine($logline)
					
				#Write back to the host if $Writebacktohost is true.
				if ($WriteBackToHost) {
					#switch($PSCmdlet.GetVariableValue('WarningPreference')){
					switch ($LoggingPreference['WarningPreference']) {
						'Continue' { $WarningPreference = 'Continue'; Write-Warning -Message "$message"; $WarningPreference = '' }
						'Stop' { $WarningPreference = 'Stop'; Write-Warning -Message "$message"; $WarningPreference = '' }
						'Inquire' { $WarningPreference = 'Inquire'; Write-Warning -Message "$message"; $WarningPreference = '' }
					}#switch
				}#if writeback
			}#Warning
			3 {  
				#This if statement is to catch the two different types of errors that may come through. 
				#A normal terminating exception will have all the information that is needed, if it's a user generated error by using Write-Error,
				#then the else statment will setup all the information we would like to log.   
				If ($message.Exception.Message) {
					#Build log message                                      
					$LogMessage = $Type + ": " + [string]$message.Exception.Message + " Command: '" + [string]$message.InvocationInfo.MyCommand + `
						"' Line: '" + [string]$message.InvocationInfo.Line.Trim() + "'"
					$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
						"$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
	
					$logline = $LineTemplate -f $LineContent
					$global:ErrorLogStream.WriteLine($logline)
						
					#Write back to the host if $Writebacktohost is true.
					if ($WriteBackToHost) {
						switch ($LoggingPreference['ErrorActionPreference']) {
							'Stop' { $ErrorActionPreference = 'Stop'; $Host.Ui.WriteErrorLine("ERROR: $([String]$message.Exception.Message)"); Write-Error $message -ErrorAction 'Stop'; $ErrorActionPreference = '' }
							'Inquire' { $ErrorActionPreference = 'Inquire'; $Host.Ui.WriteErrorLine("ERROR: $([String]$message.Exception.Message)"); Write-Error $message -ErrorAction 'Inquire'; $ErrorActionPreference = '' }
							'Continue' { $ErrorActionPreference = 'Continue'; $Host.Ui.WriteErrorLine("ERROR: $([String]$message.Exception.Message)"); $ErrorActionPreference = '' }
							'Suspend' { $ErrorActionPreference = 'Suspend'; $Host.Ui.WriteErrorLine("ERROR: $([String]$message.Exception.Message)"); Write-Error $message -ErrorAction 'Suspend'; $ErrorActionPreference = '' }
						}#switch
					}#if writeback
				}#if standard error
				else {
					#Custom error message so build out the Exception object
					[System.Exception]$Exception = $message
					[String]$ErrorID = 'Custom Error'
					[System.Management.Automation.ErrorCategory]$ErrorCategory = [Management.Automation.ErrorCategory]::WriteError
					$ErrorRecord = New-Object Management.automation.errorrecord ($Exception, $ErrorID, $ErrorCategory, $message)
					$message = $ErrorRecord
	
					#Build log message                
					$LogMessage = $Type + ": " + [string]$message.Exception.Message
					$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
						"$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
	
					$logline = $LineTemplate -f $LineContent
					$global:ErrorLogStream.WriteLine($logline)
							
					#Write back to the host if $Writebacktohost is true.
					if ($WriteBackToHost) {
						switch ($LoggingPreference['ErrorActionPreference']) {
							'Stop' { $ErrorActionPreference = 'Stop'; $Host.Ui.WriteErrorLine("ERROR: $([String]$message.Exception.Message)"); Write-Error $message -ErrorAction 'Stop'; $ErrorActionPreference = '' }
							'Inquire' { $ErrorActionPreference = 'Inquire'; $Host.Ui.WriteErrorLine("ERROR: $([String]$message.Exception.Message)"); Write-Error $message -ErrorAction 'Inquire'; $ErrorActionPreference = '' }
							'Continue' { $ErrorActionPreference = 'Continue'; $Host.Ui.WriteErrorLine("ERROR: $([String]$message.Exception.Message)"); $ErrorActionPreference = '' }
							'Suspend' { $ErrorActionPreference = 'Suspend'; $Host.Ui.WriteErrorLine("ERROR: $([String]$message.Exception.Message)"); Write-Error $message -ErrorAction 'Suspend'; $ErrorActionPreference = '' }
						}#switch
					}#if writeback
				}#else custom error
			}#Error
			4 {  
				#Build log message                
				$LogMessage = $Type + ": " + $message
				$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
					"$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
					
				$logline = $LineTemplate -f $LineContent
				$global:ErrorLogStream.WriteLine($logline)
	
				#Write back to the host if $Writebacktohost is true.
				if ($WriteBackToHost) {
					switch ($LoggingPreference['VerbosePreference']) {
						'Continue' { $VerbosePreference = 'Continue'; Write-Verbose -Message "$message"; $VerbosePreference = '' }
						'Inquire' { $VerbosePreference = 'Inquire'; Write-Verbose -Message "$message"; $VerbosePreference = '' }
						'Stop' { $VerbosePreference = 'Stop'; Write-Verbose -Message "$message"; $VerbosePreference = '' }
					}#switch
				}#if writeback
			}#Verbose
			5 {  
				#Build log message                
				$LogMessage = $Type + ": " + $message
				$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
					"$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
					
				$logline = $LineTemplate -f $LineContent
				$global:ErrorLogStream.WriteLine($logline)
					
				#Write back to the host if $Writebacktohost is true.
				if ($WriteBackToHost) {
					switch ($LoggingPreference['DebugPreference']) {
						'Continue' { $DebugPreference = 'Continue'; Write-Debug -Message "$message"; $DebugPreference = '' }
						'Inquire' { $DebugPreference = 'Inquire'; Write-Debug -Message "$message"; $DebugPreference = '' }
						'Stop' { $DebugPreference = 'Stop'; Write-Debug -Message "$message"; $DebugPreference = '' }
					}#switch
				}#if writeback
			}#Debug
			6 {  
				#Build log message                
				$LogMessage = $Type + ": " + $message
				$LineContent = $LogMessage, $TimeGenerated.ToString("HH:mm:ss.fff+000"), $TimeGenerated.ToString("MM-dd-yyyy"), `
					"$($CallingInfo.ScriptName | Split-Path -Leaf):$($CallingInfo.ScriptLineNumber)", $userContext, $severity, $CallingInfo.ScriptName
					
				$logline = $LineTemplate -f $LineContent
				$global:ErrorLogStream.WriteLine($logline)
	
				#Write back to the host if $Writebacktohost is true.
				if ($WriteBackToHost) {
					switch ($LoggingPreference['InformationPreference']) {
						'SilentlyContinue' { $InformationPreference = 'Continue'; Write-Information -Message "INFORMATION: $message"; $InformationPreference = '' }						
						'Continue' { $InformationPreference = 'Continue'; Write-Information -Message "INFORMATION: $message"; $InformationPreference = '' }
						'Inquire' { $InformationPreference = 'Inquire'; Write-Information -Message "INFORMATION: $message"; $InformationPreference = '' }
						'Stop' { $InformationPreference = 'Stop'; Write-Information -Message "INFORMATION: $message"; $InformationPreference = '' }
						'Suspend' { $InformationPreference = 'Suspend'; Write-Information -Message "INFORMATION: $message"; $InformationPreference = '' }
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
}
#endregion

#region Script and Module Requirements
Clear-Host
# Checking For Windows PowerShell 5.1 - needed for standard logging function 
#Requires -Version 5.1

# Force TLS 1.2 for connections to PSGallery
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 

# If the proxy addresses will be written to ExchangeOnlineManagement make sure the module is installed
If ($RestoreToLocation -eq "ExchangeOnline") {
	# Checking for the ExchangeOnlineManagement Module - used for access to Get-AzureADUser commands
	$ExoModule = Get-Module -ListAvailable -Verbose:$false | Where-Object { $_.Name -eq "ExchangeOnlineManagement" }
	If (!$ExoModule) {
		# Exchange Online Management module is missing lets see if Nuget is available to install it now
		Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("ExchangeOnlineManagement module is not installed") -Type Information -WriteBackToHost
		$Nuget = Get-PackageProvider -ListAvailable | Where-Object { $_.Name -eq "Nuget" }
		If (!$Nuget) {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Nuget Package Provider is not installed currently") -Type Warning -WriteBackToHost 
			Try {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Nuget Package Provider will attempt to be installed") -Type Information -WriteBackToHost
				Install-PackageProvider -Name Nuget -Force -Scope AllUsers -ErrorAction Stop
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Nuget Package Provider is now installed") -Type Information -WriteBackToHost 
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to install the Nuget Package Provider from PSGallery, try running as admin") -Type Error -WriteBackToHost
				Exit
			}                
		}

		# Nuget is installed so we can try to install the module from PSGallery
		Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("ExchangeOnlineManagement module is not installed currently") -Type Warning -WriteBackToHost 
		Try {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("ExchangeOnlineManagement module will attempt to be installed") -Type Information -WriteBackToHost
			Install-Module -Name ExchangeOnlineManagement -Force -Scope AllUsers -ErrorAction Stop
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("ExchangeOnlineManagement module is now installed") -Type Information -WriteBackToHost 
		} Catch {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to install the ExchangeOnlineManagement module from PSGallery, try running as admin") -Type Error -WriteBackToHost
			Exit
		}
	} Else {
		# Check for multiple verisons of the ExchangeOnlineManagement Module and uninstall them all
		If ($ExoModule.count -ne 1) {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Attempting to remove all instances of ExchangeOnlineManagement module") -Type Warning -WriteBackToHost 
			Try {
				Uninstall-Module -Name ExchangeOnlineManagement -Force -AllVersions -ErrorAction Stop
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("All versions of the ExchangeOnlineManagement module are now uninstalled") -Type Information -WriteBackToHost
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to uninstall the ExchangeOnlineManagement module, try running as Admin") -Type Error -WriteBackToHost
				Exit
			}

			# Install the latest version from PSGallery
			Try {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Attempting to install the ExchangeOnlineManagement module") -Type Warning -WriteBackToHost 
				Install-Module -Name ExchangeOnlineManagement -Force -Scope AllUsers -ErrorAction Stop
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("ExchangeOnlineManagement module is now installed") -Type Information -WriteBackToHost 
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to install the ExchangeOnlineManagement module from PSGallery, try running as Admin") -Type Error -WriteBackToHost
				Exit
			}
		}
	}
}

# Connect to Exchange On-Premises or Exchange Online
If ($RestoreToLocation -eq "On-Premises") {
	If ($OrgMgmtAdmin) {
		$ExchangeCred = Import-Clixml -Path $OrgMgmtAdmin
		Connect-Exchange -ExchangeServer $ExchangeServer
	} Else {
		$ExchangeCred = Get-Credential
		Connect-Exchange -ExchangeServer $ExchangeServer
	}
} Else {
	If ($O365Admin) {
		If ($MFAEnabled) {
			$ExchangeCred = Import-Clixml -Path $O365Admin
			Connect-ExOnline -MFAEnabled
		} Else {
			$ExchangeCred = Import-Clixml -Path $O365Admin
			Connect-ExOnline
		}
	} Else {
		If ($MFAEnabled) {
			$ExchangeCred = Get-Credential
			Connect-ExOnline -MFAEnabled
		} Else {
			$ExchangeCred = Get-Credential
			Connect-ExOnline
		}
	}
}
#endregion

#region Main
# Get the list of accounts to target from a user selected CSV file
$inputFile = Get-FileName -StartDirectory $StartDirectory
$users = Import-Csv -Path $inputFile -Delimiter "`t"

$p = 0
ForEach ($user in $users) {
	# Progress Bar
	If ($users.count -gt 1) {
		$p++
		Write-Progress -Id 1 -Activity "Processing $p of $($users.count) total addresses" -Status ("{0:P2}" -f ($p / $($users).Count)) -CurrentOperation $user.masterTargetEmail
		Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("") -Type Information -WriteBackToHost
		If ($user.EmailAddress) {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Restoring address " + $user.EmailAddress + " for account " + $user.PrimarySmtpAddress) -Type Information -WriteBackToHost
		} Else {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Restoring address " + $user.LegacyExchangeDN + " for account " + $user.PrimarySmtpAddress) -Type Information -WriteBackToHost
		}
	}

	# Determine logic for user or group object in the target
	#If ($user.targetSamAccountName){
	$targetObject = $null
	Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message "Using alias of $($user.masterTargetEmail) to find a match in the target" -Type Information -WriteBackToHost
	#$targetObject = Get-Recipient -Identity $user.targetSamAccountName -ErrorAction Stop
	$targetObject = Get-Recipient -Identity $user.masterTargetEmail -Verbose:$false #-ErrorAction Stop
		
	# }
	# Else{
	# 	$targetObject = $null
	# 	Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The target object is of type group and will use the " + $user.Guid + " to find a match in the tagret") -Type Information -WriteBackToHost
	# 	#$targetObject = Get-Recipient -Filter "CustomAttribute8 -eq '$($user.Guid)'" -ErrorAction Stop
	# 	$targetObject = Get-Recipient -Identity $user.masterTargetEmail -ErrorAction Stop
	# }
	
	# Grab existing target object information and type of object (hard fail if one does not exist)
	If ($targetObject) {
		Switch ($targetObject.RecipientTypeDetails) {
			UserMailbox { $targetObjectType = 'UserMailbox' }
			RoomMailbox { $targetObjectType = 'UserMailbox' }
			EquipmentMailbox { $targetObjectType = 'UserMailbox' }
			SharedMailbox { $targetObjectType = 'UserMailbox' }
			RemoteUserMailbox { $targetObjectType = 'RemoteUserMailbox' }
			RemoteRoomMailbox { $targetObjectType = 'RemoteUserMailbox' }
			RemoteEquipmentMailbox { $targetObjectType = 'RemoteUserMailbox' }
			RemoteSharedMailbox { $targetObjectType = 'RemoteUserMailbox' }
			MailUser { $targetObjectType = 'MailUser' }
			MailUniversalDistributionGroup { $targetObjectType = 'MailUniversalDistributionGroup' }
			RoomList { $targetObjectType = 'MailUniversalDistributionGroup' }
			GroupMailbox { $targetObjectType = 'GroupMailbox' }
			Default { $targetObjectType = 'UnknownType' }
		}
	} Else {
		Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to find the target object") -Type Error -WriteBackToHost
		Continue
	}
	
	# Split the address to help identify the type SIP or SMTP address for additional processing later
	$addressType = $null
	If ($user.EmailAddress) {
		$addressType = $user.EmailAddress.Split(":")
	}

	# Section focused on MailUniversalDistributionGroup
	If ($targetObjectType -eq "MailUniversalDistributionGroup") {
		If ($ObjectType -eq "All" -or $ObjectType -eq "Groups") {
			#Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The target object type is a MailUniversalDistributionGroup") -Type Information
			If ($RestoreToLocation -eq "ExchangeOnline") {
				If ($addressType) {
					# Apply SMTP or smtp address to object
					If ($addressType[0] -clike "SMTP") {
						If ($ConvertToAlias) {
							#Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The ConvertToAlias switch was used this address with be added as an alias") -Type Information -WriteBackToHost
							Try {
								If ($DisableReportMode) {
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias address of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information 
									Set-DistributionGroup -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] } -ErrorAction Stop
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information 
								} Else {
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias address of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information 
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information 
								}
							} Catch {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias address of " + $addressType[1] + " to group mailbox " + $targetObject.Alias) -Type Error -WriteBackToHost
							}
						} Else {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The the address to replace on the group mailbox is a primary") -Type Information 
							Try {
								If ($DisableReportMode) {
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The primary address of " + $targetObject.PrimarySmtpAddress + " is about to be replaced with " + $addressType[1]) -Type Information 
									Set-DistributionGroup -Identity $targetObject.Alias -PrimarySmtpAddress $addressType[1] =-ErrorAction Stop
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The primary address of " + $addressType[1] + " was set on group mailbox " + $targetObject.Alias) -Type Information 
								} Else {
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The primary address of " + $targetObject.PrimarySmtpAddress + " is about to be replaced with " + $addressType[1]) -Type Information 
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The primary address of " + $addressType[1] + " was set on group mailbox " + $targetObject.Alias) -Type Information 
								}
							} Catch {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to update PrimarySmtpAddress for group mailbox " + $targetObject.Alias) -Type Error -WriteBackToHost
							}
						}
					} Else {
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias address of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information 
								Set-DistributionGroup -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] } -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information 
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias address of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information 
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information 
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias address of " + $addressType[1] + " to group mailbox " + $targetObject.Alias) -Type Error -WriteBackToHost
						}
					}
				}
					
				# Covert LegacyExchangeDN to x500 address and add as an alias
				If ($ConvertLegDNtoX500) {
					If ($user.LegacyExchangeDN) {
						$targetX500 = ConvertTo-X500 -LegacyDN $user.LegacyExchangeDN
						If ($targetX500) {
							Try {
								If ($DisableReportMode) {
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $targetX500 + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
									Set-DistributionGroup -Identity $targetObject.Alias -EmailAddresses @{Add = $targetX500 } -ErrorAction Stop
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $targetX500 + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
								} Else {
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $targetX500 + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $targetX500 + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
								}
							} Catch {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias of " + $targetX500 + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
							}
						}
					}
				}
			}
		} Else {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The objects in scope for restore was mailboxes only, so groups are being skipped") -Type Information 
		}
	}

	# Section focused on GroupMailbox
	If ($targetObjectType -eq "GroupMailbox") {
		If ($ObjectType -eq "All" -or $ObjectType -eq "Groups") {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The target object type is a group mailbox") -Type Information
			If ($RestoreToLocation -eq "ExchangeOnline") {
				If (!$IgnoreGroupMailboxes) {
					Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The ignore group mailboxes is set to false") -Type Information 
					If ($addressType) {
						# Apply SMTP or smtp address to object
						If ($addressType[0] -clike "SMTP") {
							If ($ConvertToAlias) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The ConvertToAlias switch was used this address with be added as an alias") -Type Information -WriteBackToHost
								Try {
									If ($DisableReportMode) {
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias address of " + $addressType[1] + " is about to be added to " + $targetObject.Identity) -Type Information 
										Set-UnifiedGroup -Identity $targetObject.Identity -EmailAddresses @{Add = $addressType[1] } -ErrorAction Stop
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias address of " + $addressType[1] + " was successfully added to " + $targetObject.Identity) -Type Information 
									} Else {
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias address of " + $addressType[1] + " is about to be added to " + $targetObject.Identity) -Type Information 
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias address of " + $addressType[1] + " was successfully added to " + $targetObject.Identity) -Type Information 
									}
								} Catch {
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias address of " + $addressType[1] + " to group mailbox " + $targetObject.Identity) -Type Error -WriteBackToHost
								}
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The address to replace on the group mailbox is a primary") -Type Information 
								Try {
									If ($DisableReportMode) {
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The primary address of " + $targetObject.PrimarySmtpAddress + " is about to be replaced with " + $addressType[1]) -Type Information 
										Set-UnifiedGroup -Identity $targetObject.Identity -PrimarySmtpAddress $addressType[1] =-ErrorAction Stop
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The primary address of " + $addressType[1] + " was set on group mailbox " + $targetObject.Identity) -Type Information 
									} Else {
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The primary address of " + $targetObject.PrimarySmtpAddress + " is about to be replaced with " + $addressType[1]) -Type Information 
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The primary address of " + $addressType[1] + " was set on group mailbox " + $targetObject.Identity) -Type Information 
									}
								} Catch {
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to update PrimarySmtpAddress for group mailbox " + $targetObject.Alias) -Type Error -WriteBackToHost
								}
							}
						} Else {
							Try {
								If ($DisableReportMode) {
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias address of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information 
									Set-GroupMailbox -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] } -ErrorAction Stop
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information 
								} Else {
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias address of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information 
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information 
								}
							} Catch {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias address of " + $addressType[1] + " to group mailbox " + $targetObject.Alias) -Type Error -WriteBackToHost
							}
						}
					}
					
					# Covert LegacyExchangeDN to x500 address and add as an alias
					If ($ConvertLegDNtoX500) {
						If ($user.LegacyExchangeDN) {
							$targetX500 = ConvertTo-X500 -LegacyDN $user.LegacyExchangeDN
							If ($targetX500) {
								Try {
									If ($DisableReportMode) {
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $targetX500 + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
										Set-GroupMailbox -Identity $targetObject.Alias -EmailAddresses @{Add = $targetX500 } -ErrorAction Stop
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $targetX500 + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
									} Else {
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $targetX500 + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
										Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $targetX500 + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
									}
								} Catch {
									Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias of " + $targetX500 + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
								}
							}
						}
					}
				}
			} Else {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The Group Mailboxes only exist in ExchangeOnline, Check the restore location defined") -Type Information 
			}
		} Else {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The objects in scope for restore was mailboxes only, so groups are being skipped") -Type Information 
		}
	}

	# Section focused on MailUser
	If ($targetObjectType -eq "MailUser") {
		If ($ObjectType -eq "All" -or $ObjectType -eq "Mailboxes") {
			If ($addressType) {
				# Apply SIP address to object
				If ($addressType[0] -clike "SIP") {
					Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The address we are working with is a SIP address") -Type Information -WriteBackToHost
					If (!$IgnoreSIP) {
						# Checking if we are re-applying SIP addresses
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The switch to ignore SIP addresses is set to False") -Type Information -WriteBackToHost
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The SIP address of " + $addressType[1] + " will be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Set-MailUser -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] } -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The SIP address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The SIP address of " + $addressType[1] + " will be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The SIP address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to set the SIP address of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}
					}
				}

				# Apply SMTP or smtp address to object
				If ($addressType[0] -clike "SMTP") {
					Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The address is of type SMTP") -Type Information -WriteBackToHost
					If ($ConvertToAlias) {
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The ConvertToAlias switch was used this address with be added as an alias") -Type Information -WriteBackToHost
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Set-MailUser -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] } -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}
					} Else {
						Try {	
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The PrimarySmtpAddress of " + $targetObject.PrimarySmtpAddress + " is about to be replaced by " + $addressType[1]) -Type Information -WriteBackToHost
								Set-MailUser -Identity $targetObject.Alias -ExternalEmailAddress $addressType[1] -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The PrimarySmtpAddress of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The PrimarySmtpAddress of " + $targetObject.PrimarySmtpAddress + " is about to be replaced by " + $addressType[1]) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The PrimarySmtpAddress of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to set the PrimarySmtpAddress of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}
					}
				} Else {
					Try {
						If ($DisableReportMode) {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							Set-MailUser -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] } -ErrorAction Stop
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
						} Else {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
						}
					} Catch {
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
					}
				}
			}

			# Covert LegacyExchangeDN to x500 address and add as an alias
			If ($ConvertLegDNtoX500) {
				If ($user.LegacyExchangeDN) {
					$targetX500 = ConvertTo-X500 -LegacyDN $user.LegacyExchangeDN
					If ($targetX500) {
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $targetX500 + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Set-MailUser -Identity $targetObject.Alias -EmailAddresses @{Add = $targetX500 } -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $targetX500 + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $targetX500 + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $targetX500 + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias of " + $targetX500 + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}
					}
				}
			}
		} Else {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The objects in scope for restore was groups only, so mail-users are being skipped") -Type Information 
		}
	}

	# Section focused on UserMailbox
	If ($targetObjectType -eq "UserMailbox") {
		If ($ObjectType -eq "All" -or $ObjectType -eq "Mailboxes") {
			If ($addressType) {
				# Apply SIP address to object
				If ($addressType[0] -clike "SIP") {
					Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The address we are working with is a SIP address") -Type Information -WriteBackToHost
					If (!$IgnoreSIP) {
						# Checking if we are re-applying SIP addresses
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The switch to ignore SIP addresses is set to False") -Type Information -WriteBackToHost
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The SIP address of " + $addressType[1] + " will be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Set-Mailbox -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] } # -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The SIP address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The SIP address of " + $addressType[1] + " will be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The SIP address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to set the SIP address of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}
					} Else {
						# The switch to IgnoreSIPAddress was used
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The switch to ignore all SIP addresses was used, the address of " + $addressType[1] + " was skipped") -Type Warning -WriteBackToHost 
						Continue # Move to next address SIP is not being re-applied
					}
				}
				
				# Apply SMTP or smtp address to object
				If ($addressType[0] -clike "SMTP") {
					#Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The address is of type SMTP") -Type Information -WriteBackToHost
					If ($ConvertToAlias) {
						#Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The ConvertToAlias switch was used this address with be added as an alias") -Type Information -WriteBackToHost
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Set-Mailbox -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] } # -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}
					} Else {
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The ConvertToAlias switch was not used this address will be added as a primary") -Type Information -WriteBackToHost
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The PrimarySmtpAddress of " + $targetObject.PrimarySmtpAddress + " is about to be replaced by " + $addressType[1]) -Type Information -WriteBackToHost
								Set-Mailbox -Identity $targetObject.Alias -PrimarySmtpAddress $addressType[1] # -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The PrimarySmtpAddress of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The PrimarySmtpAddress of " + $targetObject.PrimarySmtpAddress + " is about to be replaced by " + $addressType[1]) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The PrimarySmtpAddress of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to set the PrimarySmtpAddress of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}			 
					}
				} Else {
					Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The address was previously an alias and will be restored as an alias") -Type Information -WriteBackToHost
					Try {
						If ($DisableReportMode) {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							Set-Mailbox -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] } # -ErrorAction Stop
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
						} Else {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
						}
					} Catch {
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
					}
				}
			}

			# Covert LegacyExchangeDN to x500 address and add as an alias
			If ($ConvertLegDNtoX500) {
				If ($user.LegacyExchangeDN) {
					$targetX500 = ConvertTo-X500 -LegacyDN $user.LegacyExchangeDN
					If ($targetX500) {
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $targetX500 + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Set-Mailbox -Identity $targetObject.Alias -EmailAddresses @{Add = $targetX500 } # -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $targetX500 + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $targetX500 + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $targetX500 + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias of " + $targetX500 + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}
					}
				}
			}
			
		} Else {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The objects in scope for restore was groups only, so UserMailboxes are being skipped") -Type Information 
		}
	}

	# Section focused on RemoteUserMailbox
	If ($targetObjectType -eq "RemoteUserMailbox") {
		If ($ObjectType -eq "All" -or $ObjectType -eq "Mailboxes") {
			If ($addressType) {
				# Apply SIP address to object
				If ($addressType[0] -clike "SIP") {
					Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The address we are working with is a SIP address") -Type Information -WriteBackToHost
					If (!$IgnoreSIP) {
						# Checking if we are re-applying SIP addresses
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The switch to ignore SIP addresses is set to False") -Type Information -WriteBackToHost
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The SIP address of " + $addressType[1] + " will be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Set-RemoteMailbox -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] } -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The SIP address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The SIP address of " + $addressType[1] + " will be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The SIP address of " + $addressType[1] + " was successfully added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to set the SIP address of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}
					} Else {
						# The switch to IgnoreSIPAddress was used
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The switch to ignore all SIP addresses was used, the address of " + $addressType[1] + " was skipped") -Type Warning -WriteBackToHost 
						Continue # Move to next address SIP is not being re-applied
					}
				}
				
				# Apply SMTP or smtp address to object
				If ($addressType[0] -clike "SMTP") {
					Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The address is of type SMTP") -Type Information -WriteBackToHost
					If ($ConvertToAlias) {
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The ConvertToAlias switch was used this address with be added as an alias") -Type Information -WriteBackToHost
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Set-RemoteMailbox -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] } -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}
					} Else {
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The ConvertToAlias switch was not used this address will be added as a primary") -Type Information -WriteBackToHost
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The PrimarySmtpAddress of " + $targetObject.PrimarySmtpAddress + " is about to be replaced by " + $addressType[1]) -Type Information -WriteBackToHost
								Set-RemoteMailbox -Identity $targetObject.Alias -PrimarySmtpAddress $addressType[1]# -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The PrimarySmtpAddress of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The PrimarySmtpAddress of " + $targetObject.PrimarySmtpAddress + " is about to be replaced by " + $addressType[1]) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE The PrimarySmtpAddress of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to set the PrimarySmtpAddress of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}			 
					}
				} Else {
					Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The address was previously an alias and will be restored as an alias") -Type Information -WriteBackToHost
					Try {
						If ($DisableReportMode) {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							Set-RemoteMailbox -Identity $targetObject.Alias -EmailAddresses @{Add = $addressType[1] }# -ErrorAction Stop
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
						} Else {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $addressType[1] + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
						}
					} Catch {
						Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias of " + $addressType[1] + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
					}
				}
			}

			# Covert LegacyExchangeDN to x500 address and add as an alias
			If ($ConvertLegDNtoX500) {
				If ($user.LegacyExchangeDN) {
					$targetX500 = ConvertTo-X500 -LegacyDN $user.LegacyExchangeDN
					If ($targetX500) {
						Try {
							If ($DisableReportMode) {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $targetX500 + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Set-RemoteMailbox -Identity $targetObject.Alias -EmailAddresses @{Add = $targetX500 } -ErrorAction Stop
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("An alias of " + $targetX500 + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							} Else {
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $targetX500 + " is about to be added to " + $targetObject.Alias) -Type Information -WriteBackToHost
								Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("REPORT MODE An alias of " + $targetX500 + " was added successfully to " + $targetObject.Alias) -Type Information -WriteBackToHost
							}
						} Catch {
							Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to add an alias of " + $targetX500 + " onto account " + $targetObject.Alias) -Type Error -WriteBackToHost
						}
					}
				}
			}
		} Else {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The objects in scope for restore was groups only, so Remote UserMailboxes are being skipped") -Type Information 
		}
	}

	# Section focused on UnknownType
	If ($targetObjectType -eq "UnknownType") {
		Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The object was determined to be of Unknown Type, no further processing will happen.") -Type Information
	}
}
#endregion

#region CloseOut
Get-PSSession | Remove-PSSession
Disconnect-ExchangeOnline
#endregion
