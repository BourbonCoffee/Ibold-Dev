<#
.SYNOPSIS
This script is used to discover Mailboxes, MailUsers, Distribution, Security and Unified Groups using a given SMTP domain as a email address
The script by default will do a report of all above stated object types having an email address from a given SMTP domain
The script has a optional swicth parameter to remove all email addresses for a given SMTP domain
The csv file must contain a column called DomainName with each row having a value eg. contonso.com
* Note UserPrincipalName should be changed first when using Office 365, as it will lock certain proxyAddresses
* Note does not cover the Get-AzureADUser scenario where the license is just removed from the mailbox (License and then remove or delete the account) 

.DESCRIPTION
This script is used to discover Mailboxes, MailUsers, Distribution, Security and Unified Groups using a given SMTP domain as a email address
The script by default will do a report of all above stated object types having an email address from a given SMTP domain
The script has a optional swicth parameter to remove all email addresses for a given SMTP domain
The csv file must contain a column called DomainName with each row having a value eg. contonso.com
* Note UserPrincipalName should be changed first when using Office 365, as it will lock certain proxyAddresses
* Note does not cover the Get-AzureADUser scenario where the license is just removed from the mailbox (License and then remove or delete the account)

.PARAMETER StartDirectory
This parameter defines the default location when asking for the input csv file.
The csv file must contain a column called DomainName with each row having a value eg. contonso.com
    Example:"C:\Temp"

.PARAMETER Type
This parameter has two options "On-Premises" and "Office365". This parameter is used to decide where the object types are located.
    Default Value: "Office365"

.PARAMETER ObjectType
This parameter has three options "MailboxOnly", "GroupsOnly" and "All". This parameter is used to decide what object types are in scope. 
If Type equals "On-Premises" Unified Groups are automatically skipped, since that is not an option.
    Default Value: "All"

.PARAMETER IncludeLegacyExchangeDN
This parameter is a switch meaning if added or set to true if takes effect. This will include the LegacyExchangeDN to the export file
    Default Value: $False

.PARAMETER OrgMgmtAdmin
This parameter defines the location and xml file name to the encrypted credentials used for Exchange on-premises access, if none are provided Kerberos is tried.
The credentials must be encrypted using CliXML
    Example value: "C:\Temp\orgmgmt.xml"

.PARAMETER ExchangeServer
This parameter defines the fully qualified server name of the on-premises Exchange server.
    Example value: server1.contoso.com

.PARAMETER O365Admin
This parameter defines the location and xml file name to the encrypted credentials required for Office 365 access.
The credentials must be encrypted using CliXML
If MFA is required leave this parameter blank and you will be prompted twice for credentials (Exchange Online and MSOL)
    Example value: "C:\Temp\o365admin.xml"

.PARAMETER DestroyMethod
If Destroy is selected it takes the output from the previous discovery process and DELETES the email addresses it finds for a given domain.
If DestroyOnly is selected it will prompt for a CSV of email addresses to DELETE. 

.EXAMPLE
Proxy-Seek-and-Destroy.ps1 -StartDirectory C:\temp -Type On-Premises -ObjectType MailboxOnly -ExchangeServer server.domain.com -OrgMgmtAdmin C:\temp\creds.xml

.EXAMPLE
Proxy-Seek-and-Destroy.ps1 -StartDirectory C:\temp -Type On-Premises -ObjectType MailboxOnly -ExchangeServer server.domain.com -OrgMgmtAdmin C:\temp\creds.xml -IncludeLegacyExchangeDN

.EXAMPLE
Proxy-Seek-and-Destroy.ps1 -StartDirectory C:\temp -Type On-Premises -ObjectType MailboxOnly -ExchangeServer server.domain.com -OrgMgmtAdmin C:\temp\creds.xml -DestroyMethod Destroy

.EXAMPLE
Proxy-Seek-and-Destroy.ps1 -StartDirectory C:\temp -Type On-Premises -ObjectType MailboxOnly -ExchangeServer server.domain.com -OrgMgmtAdmin C:\temp\creds.xml -DestroyMethod DestroyOnly

.EXAMPLE
Proxy-Seek-and-Destroy.ps1 -StartDirectory C:\temp -Type Office365 -ObjectType All -O365Admin C:\temp\O365Creds.xml

.EXAMPLE
Proxy-Seek-and-Destroy.ps1 -StartDirectory C:\temp -Type Office365 -ObjectType All -O365Admin C:\temp\O365Creds.xml -IncludeLegacyExchangeDN

.EXAMPLE
Proxy-Seek-and-Destroy.ps1 -StartDirectory C:\temp -Type Office365 -ObjectType All -O365Admin C:\temp\O365Creds.xml -DestroyMethod Destroy

.EXAMPLE
Proxy-Seek-and-Destroy.ps1 -StartDirectory C:\temp -Type Office365 -ObjectType All -O365Admin C:\temp\O365Creds.xml -DestroyMethod DestroyOnly

#>

#region Parameters
[CmdletBinding()]
Param(
	[Parameter(Position = 1, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Specify the starting location for picking a CSV (ex: C:\Temp)")]
	[ValidateNotNullOrEmpty()]
	[String] $StartDirectory,

	[Parameter(Position = 2, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Select where the SMTP addresses should be located")]
	[ValidateSet ("On-Premises", "Office365")]
	[String] $Type = "Office365",

	[Parameter(Position = 3, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Select the type of objects with the SMTP addresses")]
	[ValidateSet ("MailboxOnly", "GroupsOnly", "All")]
	[String] $ObjectType = "All",

	[Parameter(Position = 4, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Do you want to include LegacyExchangeDN as an entry in the export")]
	[Switch] $IncludeLegacyExchangeDN,

	[Parameter(Position = 5, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Specify the path to the encrypted creds (ex: C:\Temp\OrgAdmin.xml)")]
	[ValidateNotNullOrEmpty()]
	[String] $OrgMgmtAdmin,

	[Parameter(Position = 6, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Specify the FQDN of the Exchange Server on-premises (ex: server1.contoso.com)")]
	[ValidateNotNullOrEmpty()]
	[String] $ExchangeServer,

	[Parameter(Position = 7, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Specify the path to the encrypted creds (ex: C:\Temp\O365Admin.xml)")]
	[ValidateNotNullOrEmpty()]
	[String] $O365Admin,

	[Parameter(Position = 9, Mandatory = $false, ValueFromPipeline = $false,
		HelpMessage = "Select the method to DELETE the email addresses")]
	[ValidateSet ("Destroy", "DestroyOnly")]
	[String] $DestroyMethod
)
#endregion

#region Define Variables / Setup Logging
# Standard Logging Setup
Set-Variable -Name scriptPath -Option AllScope -Scope Script -Value (Split-Path $MyInvocation.MyCommand.Path)
Set-Variable -Name scriptDir -Option AllScope -Scope Script -Value (Split-Path $ScriptPath)
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
Function Get-FileName ($StartDirectory) {
	# Prompts the user for the input file starting in $StartDirectory
	# Log file is created in the same directory as the input file that is selected
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$openFileDialog.initialDirectory = $StartDirectory
	$openFileDialog.filter = "All files (*.*)| *.*"
	$openFileDialog.ShowDialog() | Out-Null
	$openFileDialog.filename
}

Function Connect-Exchange-On-Premises {
	Param(
		[Parameter(Position = 1, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[String] $ExchangeServer,

		[Parameter(Position = 2, Mandatory = $false)]
		[String] $ExchangeCreds
	)

	$SessionOptions = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -OpenTimeout 20000
	
	# Attempt #1 http and current logged on user
	Try {
		$session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell") -Authentication Kerberos -AllowRedirection -SessionOption $SessionOptions -ErrorAction Stop
		Import-PSSession -Session $session -AllowClobber
		Set-ADServerSettings -ViewEntireForest:$true
		Return $session
	} Catch {
		# Attempt #2 https and current logged on user
		Try {
			$session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("https://" + $ExchangeServer + "/PowerShell") -Authentication Kerberos -AllowRedirection -SessionOption $SessionOptions -ErrorAction Stop
			Import-PSSession -Session $session -AllowClobber
			Set-ADServerSettings -ViewEntireForest:$true
			Return $session
		} Catch {
			# Attempt #3 https and provided creds
			Try {
				$session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("https://" + $ExchangeServer + "/PowerShell") -Credential $ExchangeCreds -AllowRedirection -SessionOption $SessionOptions -ErrorAction Stop
				Import-PSSession -Session $session -AllowClobber
				Set-ADServerSettings -ViewEntireForest:$true
				Return $session
			} Catch {
				# Attempt #4 http and provided creds
				Try {
					$session = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell") -Credential $ExchangeCreds -AllowRedirection -SessionOption $SessionOptions -ErrorAction Stop
					Import-PSSession -Session $session -AllowClobber
					Set-ADServerSettings -ViewEntireForest:$true
					Return $session
				} Catch {
					Exit
				}
			}
		}
	}
}

Function Get-Mailbox-Detail {
	# Find all mailboxes with the desired domain name being used
	$domainMatch = "*@" + $domain.DomainName
	$domainString = ("`'$domainMatch`'").ToString()
	$recipients = Get-Recipient -ResultSize Unlimited -Filter "EmailAddresses -like $domainString -and (RecipientTypeDetails -eq 'UserMailbox' -or RecipientTypeDetails -eq 'SharedMailbox' -or RecipientTypeDetails -eq 'RoomMailbox')"
	$totalcount = $recipients | Measure-Object
	Out-CMTraceLog -Logfile $logFileName -Message ("We discovered " + $totalcount.count + " mailboxes with " + $domainMatch + " used.") -Type Information -WriteBackToHost
	
	# Loop through the list of mailboxes and find the exact address that matched and output info about it
	$i = 0
	ForEach ($recipient in $recipients) {
		If ($recipients.count -gt 1) {
			$i++
			Write-Progress -ParentId 1 -Activity "Processing $i of $($recipients.count) users: $($recipient.PrimarySmtpAddress)" -Status ("{0:P2}" -f ($i / $($recipients).Count))
		}
		$addresses = $recipient.EmailAddresses -like $domainMatch
		ForEach ($address in $addresses) {
			$record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname
			Add-Content $outputFile -Value $record
		}

		If ($IncludeLegacyExchangeDN) {
			# Export the Exchange Legacy DN as well
			$record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname + "`t" + $recipient.ServerLegacyDN
			Add-Content $outputFile -Value $record
		}
	}
}

Function Get-InActive-Mailbox-Detail {
	# Find all mailboxes with the desired domain name being used
	$domainMatch = "*@" + $domain.DomainName
	$domainString = ("`'$domainMatch`'").ToString()
	$recipients = Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited -Filter "EmailAddresses -like $domainString"
	$totalcount = $recipients | Measure-Object
	Out-CMTraceLog -Logfile $logFileName -Message ("We discovered " + $totalcount.count + " inactive mailboxes with " + $domainMatch + " used.") -Type Information -WriteBackToHost
	
	# Loop through the list of mailboxes and find the exact address that matched and output info about it
	$i = 0
	ForEach ($recipient in $recipients) {
		If ($recipients.count -gt 1) {
			$i++
			Write-Progress -ParentId 1 -Activity "Processing $i of $($recipients.count) users: $($recipient.PrimarySmtpAddress)" -Status ("{0:P2}" -f ($i / $($recipients).Count))
		}
		$addresses = $recipient.EmailAddresses -like $domainMatch
		ForEach ($address in $addresses) {
			$record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + $recipient.IsInactiveMailbox + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname
			Add-Content $outputFile -Value $record
		}

		If ($IncludeLegacyExchangeDN) {
			# Export the Exchange Legacy DN as well
			$record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + $recipient.IsInactiveMailbox + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname + "`t" + $recipient.ServerLegacyDN
			Add-Content $outputFile -Value $record
		}
	}
}

Function Get-MailUser-Detail {
	# Find all mailusers with the desired domain name being used
	$domainMatch = "*@" + $domain.DomainName
	$domainString = ("`'$domainMatch`'").ToString()
	$recipients = Get-Recipient -ResultSize Unlimited -Filter "EmailAddresses -like $domainString -and RecipientTypeDetails -eq 'MailUser'"
	$totalcount = $recipients | Measure-Object
	Out-CMTraceLog -Logfile $logFileName -Message ("We discovered " + $totalcount.count + " mailusers with " + $domainMatch + " used.") -Type Information -WriteBackToHost

	# Loop through the list of mailusers and find the exact address that matched and output info about it
	$i = 0
	ForEach ($recipient in $recipients) {
		If ($recipients.count -gt 1) {
			$i++
			Write-Progress -ParentId 1 -Activity "Processing $i of $($recipients.count) users: $($recipient.PrimarySmtpAddress)" -Status ("{0:P2}" -f ($i / $($recipients).Count))
		}
		$addresses = $recipient.EmailAddresses -like $domainMatch
		ForEach ($address in $addresses) {
			$record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname
			Add-Content $outputFile -Value $record
		}
		
		If ($IncludeLegacyExchangeDN) {
			# Export the Exchange Legacy DN as well
			$record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname + "`t" + $recipient.ServerLegacyDN
			Add-Content $outputFile -Value $record
		}
	}
}

Function Get-DistributionGroup-Detail {
	# Find all mailboxes with the desired domain name being used
	$domainMatch = "*@" + $domain.DomainName
	$domainString = ("`'$domainMatch`'").ToString()
	$recipients = Get-DistributionGroup -Resultsize Unlimited -Filter "EmailAddresses -like $domainString -and (RecipientTypeDetails -eq 'MailUniversalDistributionGroup' -or RecipientTypeDetails -eq 'MailUniversalSecurityGroup')"
	$totalcount = $recipients | Measure-Object	
	Out-CMTraceLog -Logfile $logFileName -Message ("We discovered " + $totalcount.count + " groups with " + $domainMatch + " used.") -Type Information -WriteBackToHost

	# Loop through the list of groups and find the exact address that matched and output info about it
	$i = 0
	ForEach ($recipient in $recipients) {
		If ($recipients.count -gt 1) {
			$i++
			Write-Progress -ParentId 1 -Activity "Processing $i of $($recipients.count) groups: $($recipient.Displayname)" -Status ("{0:P2}" -f ($i / $($recipients).Count))
		}
		$addresses = $recipient.EmailAddresses -like $domainMatch
		ForEach ($address in $addresses) {
			$record = $recipient.Guid.Guid + "`t" + "`t" + $recipient.ExchangeObjectId + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname
			Add-Content $outputFile -Value $record
		}

		If ($IncludeLegacyExchangeDN) {
			# Export the Exchange Legacy DN as well
			$record = $recipient.Guid.Guid + "`t" + "`t" + $recipient.ExchangeObjectId + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname + "`t" + $recipient.LegacyExchangeDN
			Add-Content $outputFile -Value $record
		}		
	}
}

Function Get-RoomList-Detail {
	# Find all mailboxes with the desired domain name being used
	$domainMatch = "*@" + $domain.DomainName
	$domainString = ("`'$domainMatch`'").ToString()
	$recipients = Get-DistributionGroup -Resultsize Unlimited -Filter "EmailAddresses -like $domainString -and RecipientTypeDetails -eq 'RoomList'"
	$totalcount = $recipients | Measure-Object
	Out-CMTraceLog -Logfile $logFileName -Message ("We discovered " + $totalcount.count + " room lists with " + $domainMatch + " used.") -Type Information -WriteBackToHost

	# Loop through the list of groups/room lists and find the exact address that matched and output info about it
	$i = 0
	ForEach ($recipient in $recipients) {
		If ($recipients.count -gt 1) {
			$i++
			Write-Progress -ParentId 1 -Activity "Processing $i of $($recipients.count) groups: $($recipient.Displayname)" -Status ("{0:P2}" -f ($i / $($recipients).Count))
		}
		$addresses = $recipient.EmailAddresses -like $domainMatch
		ForEach ($address in $addresses) {
			$record = $recipient.Guid.Guid + "`t" + "`t" + $recipient.ExchangeObjectId + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname
			Add-Content $outputFile -Value $record
		}

		If ($IncludeLegacyExchangeDN) {
			# Export the Exchange Legacy DN as well
			$record = $recipient.Guid.Guid + "`t" + "`t" + $recipient.ExchangeObjectId + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname + "`t" + $recipient.LegacyExchangeDN
			Add-Content $outputFile -Value $record
		}	
	}
}

Function Get-UnifiedGroup-Detail {
	# Find all groups with the desired domain name being used
	$domainMatch = "*@" + $domain.DomainName
	$domainString = ("`'$domainMatch`'").ToString()
	$recipients = Get-Recipient -ResultSize Unlimited -Filter "EmailAddresses -like $domainString -and RecipientTypeDetails -eq 'GroupMailbox'"
	$totalcount = $recipients | Measure-Object	
	Out-CMTraceLog -Logfile $logFileName -Message ("We discovered " + $totalcount.count + " unified groups with " + $domainMatch + " used.") -Type Information -WriteBackToHost
	Out-CMTraceLog -Logfile $logFileName -Message ("") -Type Information -WriteBackToHost

	# Loop through the list of unified groups and find the exact address that matched and output info about it
	$i = 0
	ForEach ($recipient in $recipients) {
		If ($recipients.count -gt 1) {
			$i++
			Write-Progress -ParentId 1 -Activity "Processing $i of $($recipients.count) groups: $($recipient.Displayname)" -Status ("{0:P2}" -f ($i / $($recipients).Count))
		}
		$addresses = $recipient.EmailAddresses -like $domainMatch
		ForEach ($address in $addresses) {
			$record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + $address + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname
			Add-Content $outputFile -Value $record
		}

		If ($IncludeLegacyExchangeDN) {
			# Export the Exchange Legacy DN as well
			$record = $recipient.Guid.Guid + "`t" + $recipient.ExchangeGuid + "`t" + "`t" + $recipient.RecipientTypeDetails + "`t" + "`t" + $recipient.Alias + "`t" + "`t" + $recipient.PrimarySmtpAddress + "`t" + $recipient.Displayname + "`t" + $recipient.ServerLegacyDN
			Add-Content $outputFile -Value $record
		}
	}
}

Function Out-CMTraceLog {
	[CmdletBinding()] 
	Param( 
		[Parameter(Mandatory = $false,
			ValueFromPipeline = $false,
			HelpMessage = "Specify the file path for the log file. It will default to %temp%\CMTrace.log if not specified.")]
		[string]$Logfile = "C:\temp\CMTrace.log",
        
		[Parameter(Mandatory = $true,
			ValueFromPipeline = $false,
			HelpMessage = "Specify the object or string for the log")]
		$Message,

		[Parameter(Mandatory = $false,
			ValueFromPipeline = $false,
			HelpMessage = "Specify the severity of message to log.")]
		[ValidateSet('Warning', 'Error', 'Verbose', 'Debug', 'Information')] 
		[string]$Type = "Information",

		[Parameter(Mandatory = $false,
			ValueFromPipeline = $false,
			HelpMessage = "Specify whether or not to write the message back to the console. It will default to false if not specified.")]
		[switch]$WriteBackToHost,

		[Parameter(Mandatory = $false,
			ValueFromPipeline = $false,
			HelpMessage = "Specify a hashtable of the preference variables so that they can be honored. It will use the default PowerShell values if not specified.")]
		[hashtable]$LoggingPreference = @{"InformationPreference" = $InformationPreference; `
				"WarningPreference" = $WarningPreference; "ErrorActionPreference" = $ErrorActionPreference; "VerbosePreference" = $VerbosePreference; "DebugPreference" = $DebugPreference
		},

		[Parameter(Mandatory = $false,
			ValueFromPipeline = $false,
			HelpMessage = "Specify whether or not to force the message to be written to the log. It will default to false if not specified.")]
		[switch]$ForceWriteToLog = $true
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
		if ($null -eq $global:ErrorLogStream -or -not $global:ErrorLogStream.BaseStream) {
			If (-not (Test-Path -Path (Split-Path $LogFile))) { New-Item (Split-Path $LogFile) -ItemType Directory | Out-Null }
            
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
}
#endregion

#region Requirements

# Checking For Windows PowerShell 5.1 - needed for standard logging function 
#Requires -Version 5.1
# Force TLS 1.2 for connections to PSGallery
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 
Clear-Host

# Checking for the required modules - used for access to Azure and Exchange Online commands
If ($Type -eq "Office365") {
	# Checking for the Azure AD Module - used for access to Get-AzureADUser, Get-AzureADDevice, Get-ADMSGroup commands
	# Allows for the use of the Azure AD and AzureADPreview module
	$AzureADModule = Get-Module -ListAvailable -Verbose:$false | Where-Object { $_.Name -match "AzureAD" }
	If (!$AzureADModule) {
		# Azure AD module is missing lets see if Nuget is available to install it now
		Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Azure AD module is not installed") -Type Information -WriteBackToHost
		$Nuget = Get-PackageProvider -ListAvailable | Where-Object { $_.name -eq "Nuget" }
		If (!$Nuget) {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Nuget Package Provider is not installed currently") -Type Warning -WriteBackToHost -ForceWriteToLog
			Try {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Nuget Package Provider will attempt to be installed") -Type Information -WriteBackToHost
				Install-PackageProvider -Name Nuget -Force -Scope AllUsers -ErrorAction Stop
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Nuget Package Provider is now installed") -Type Information -WriteBackToHost -ForceWriteToLog
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to install the Nuget Package Provider from PSGallery, try running as admin") -Type Error -WriteBackToHost
				Exit $ExitCode
			}                
		} Else {
			# Nuget is installed so we can try to install the module from PSGallery
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Azure AD module is not installed currently") -Type Warning -WriteBackToHost -ForceWriteToLog
			Try {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Azure AD module will attempt to be installed") -Type Information -WriteBackToHost
				Install-Module -Name AzureAD -Force -Scope AllUsers -ErrorAction Stop
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Azure AD module is now installed") -Type Information -WriteBackToHost -ForceWriteToLog
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to install the Azure AD module from PSGallery, try running as admin") -Type Error -WriteBackToHost
				Exit $ExitCode
			}
		}
	}

	# Connect to Azure AD for access to Get-AzureADUser, Get-AzureADDevice, Get-ADMSGroup command
	$AzureADModule = Get-Module -ListAvailable -Verbose:$false | Where-Object { $_.Name -match "AzureAD" }
	Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Using " + $AzureADModule.Version + " version of the AzureAD module") -Type Information -ForceWriteToLog -WriteBackToHost
       
	# Determine the credential storage method for connecting via AzureAD
	$credMethod = $null
	If ($O365Admin) {
		# Check for using creds stored in encrypted XML
		Try {
			$AzureADCreds = Import-Clixml -Path $O365Admin -ErrorAction Stop
			# Creds were decrypted successfully attempting to connect to Azure AD
			Try {
				Connect-AzureAD -Credential $AzureADCreds | Out-Null
				$credMethod = "XML"
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Connected to Azure AD using credentials from " + $credMethod) -Type Information -WriteBackToHost
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable pull the users from Azure AD using credentials from XML due to " + $PSItem.Exception.Message) -Type Error -WriteBackToHost
				Exit $ExitCode
			}
		} Catch {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable pull the users from Azure AD, the stored credentials from XML are invalid") -Type Error -WriteBackToHost
			Exit $ExitCode
		}
	}
    
	If ($null -eq $credMethod) {
		# Attempt to pull them from SQL using custom encryption 
		# Read SQL looking for creds
		Try { $AzureADCreds = Get-MigratorAccountCredential -EntityName AAD -EnvironmentID $AzureADTenantName -CredentialID GlobalReader } Catch {}
		If ($AzureADCreds) {
			Try {
				Connect-AzureAD -Credential $AzureADCreds | Out-Null
				$credMethod = "SQL"
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Connected to Azure AD using credentials from " + $credMethod) -Type Information -WriteBackToHost
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable pull the users from Azure AD using credentials from SQL due to " + $PSItem.Exception.Message) -Type Warning -WriteBackToHost
				Exit $ExitCode
			}
		}
	}
     
	If ($null -eq $credMethod) {
		# Prompt for creds option (client using AD FS) 
		Try {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message "Connecting to Azure AD" -Type Information -WriteBackToHost
			Connect-AzureAD | Out-Null
			$credMethod = "Typed"
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Connected to Azure AD using credentials from " + $credMethod) -Type Information -WriteBackToHost
		} Catch {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable pull the users from Azure AD using credentials typed in due to " + $PSItem.Exception.Message) -Type Error -WriteBackToHost
			Exit $ExitCode
		}
	}

	# Checking for the Exchange Online Management Module for access to Get-MigrationUser, Get-MigrationBatch commands
	$ExchangeOnlineManagement = Get-Module -ListAvailable -Verbose:$false | Where-Object { $_.Name -eq "ExchangeOnlineManagement" }
	If (!$ExchangeOnlineManagement) {
		# Exchange Online Management module is missing lets see if Nuget is available to install it now
		$Nuget = Get-PackageProvider -ListAvailable | Where-Object { $_.Name -eq "Nuget" }
		If (!$Nuget) {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Nuget Package Provider is not installed currently") -Type Warning -WriteBackToHost -ForceWriteToLog
			Try {
				Install-PackageProvider -Name Nuget -Force -Scope AllUsers -ErrorAction Stop
				Out-CMTraceLog -Logfile $logFileName -Message ("Nuget Package Provider is now installed") -Type Information -WriteBackToHost
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("The Nuget Package Provider is now installed") -Type Information -WriteBackToHost -ForceWriteToLog
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to install the Nuget Package Provider from PSGallery, try running as admin") -Type Error -WriteBackToHost
				Exit
			}                
		} Else {
			# Nuget is installed so we can try to install the module from PSGallery
			Out-CMTraceLog -Logfile $logFileName -Message ("ExchangeOnlineManagement module is not installed currently") -Type Warning -WriteBackToHost
			Try {
				Install-Module -Name ExchangeOnlineManagement -Force -Scope AllUsers -ErrorAction Stop
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("ExchangeOnlineManagement module is now installed") -Type Information -WriteBackToHost -ForceWriteToLog
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to install the ExchangeOnlineManagement module from PSGallery, try running as admin") -Type Error -WriteBackToHost
				Exit
			}
		}
	} Else {
		# Check for multiple verisons of the ExchangeOnlineManagement Module and uninstall them all
		If ($ExchangeOnlineManagement.count -ne 1) {  
			Try {
				Uninstall-Module -Name ExchangeOnlineManagement -Force -AllVersions -ErrorAction Stop
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("All versions of the ExchangeOnlineManagement module are now uninstalled") -Type Warning -WriteBackToHost -ForceWriteToLog
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to uninstall the ExchangeOnlineManagement module, try running as Admin") -Type Error -WriteBackToHost
				Exit
			}
            
			# Install the latest version from PSGallery
			Try {
				Install-Module -Name ExchangeOnlineManagement -Force -Scope AllUsers -ErrorAction Stop
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("ExchangeOnlineManagement module is now installed") -Type Information -WriteBackToHost -ForceWriteToLog
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to install the ExchangeOnlineManagement module from PSGallery, try running as Admin") -Type Error -WriteBackToHost
				Exit
			}
		}
	}

	# Determine the credential storage method for connecting via Connect-ExchangeOnline and MSOL
	$credMethod = $null
	If ($O365Admin) {
		# Check for using creds stored in encrypted XML
		Try {
			$AdminCreds = Import-Clixml -Path $O365Admin -ErrorAction Stop
			# Creds were decrypted successfully attempting to connect to Connect-ExchangeOnline
			Try {
				Connect-ExchangeOnline -Credential $AdminCreds -ShowBanner:$false
				$credMethod = "XML"
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Connected to Exchange Online v2 (Connect-ExchangeOnline) using credentials from " + $credMethod) -Type Information -WriteBackToHost
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to connect to Exchange Online v2 (Connect-ExchangeOnline) using credentials from XML") -Type Error -WriteBackToHost
				Exit
			}
		} Catch {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to connect to Exchange Online v2 (Connect-ExchangeOnline), as the stored credentials from XML are invalid") -Type Error -WriteBackToHost
			Exit
		}
	}

	If ($null -eq $credMethod) {
		# Prompt for creds option 
		Try {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message "Connecting to Exchange Online" -Type Information -WriteBackToHost
			Connect-ExchangeOnline -ShowBanner:$false -Verbose:$false
			$credMethod = "Typed"
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Connected to Exchange Online v2 (Connect-ExchangeOnline) using credentials from " + $credMethod) -Type Information -WriteBackToHost            
		} Catch {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to connect to Exchange Online v2 (Connect-ExchangeOnline) using credentials typed in") -Type Error -WriteBackToHost
			Exit
		}
	}
} Else {
	# Determine the credential storage method for connecting to Exchange On-Premises
	$credMethod = $null
	If ($OrgMgmtAdmin) {
		# Check for using creds stored in encrypted XML
		Try {
			$Ex_OrgMgmtCreds = Import-Clixml -Path $OrgMgmtAdmin -ErrorAction Stop
			# Creds were decrypted successfully attempting to connect to Azure AD
			Try {
				Connect-Exchange-On-Premises -ExchangeServer $ExchangeServer -ExchangeCreds $Ex_OrgMgmtCreds
				$credMethod = "XML"
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Connected to Exchange on-premises using credentials from " + $credMethod) -Type Information -WriteBackToHost
			} Catch {
				Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable pull the users from Exchange on-premises using credentials from XML") -Type Error -WriteBackToHost
				Exit
			}
		} Catch {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to connect to Exchange on-premises, the stored credentials from XML are invalid") -Type Error -WriteBackToHost
			Exit
		}
	}
        
	If ($null -eq $credMethod) {
		# Prompt for creds option (client using AD FS) 
		Try {
			Connect-Exchange-On-Premises -ExchangeServer $ExchangeServer
			$credMethod = "Typed"
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Connected to Exchange on-premises using credentials from " + $credMethod) -Type Information -WriteBackToHost
        
		} Catch {
			Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $logFileName -Message ("Unable to connect to Exchange on-premises using credentials typed in") -Type Error -WriteBackToHost
			Exit
		}
	}
}
#endregion

#region Gather Addresses
If ($DestroyMethod -ne "DestroyOnly") {
	# Prompt for input file containing a column called DomainName and rows of domain names and create output file
	$filePicked = Get-FileName $StartDirectory
	$filePath = Split-Path $filePicked -Parent
	$fileName = Split-Path $filePicked -Leaf

	# Prepare output files or each object type requested
	$currentDate = Get-Date -Format "MM-dd-yyyy-hhmm-tt"
	$outputFile = ($filePath + '\' + $fileName.Substring(0, $fileName.Length - 4) + '_Domain-Results-' + $currentDate + '.txt')
	If (Test-Path -Path $outputFile) {
		Remove-Item -Path $outputFile
	}
	Add-Content $outputFile -Value "Guid`tExchangeGuid`tExchangeObjectId`tRecipientTypeDetails`tIsInactiveMailbox`tAlias`tEmailAddress`tPrimarySmtpAddress`tDisplayName`tLegacyExchangeDN"

	# Gather all object types requested with a primary or secondary addresses of the desired domain 
	$domains = Import-Csv -Path $filePicked
	$p = 0
	ForEach ($domain in $domains) {
		If ($domains.count -gt 1) {
			$p++
			Write-Progress -Id 1 -Activity "Processing $p of $($domains.count) total domains" -Status ("{0:P2}" -f ($p / $($domains).Count)) -CurrentOperation $domain.DomainName
		}

		# Gather all Mailboxes
		If ($ObjectType -eq "MailboxOnly" -or $ObjectType -eq "All") {
			Get-Mailbox-Detail
			Get-InActive-Mailbox-Detail
			Get-MailUser-Detail
		}

		# Gather all Distribution and Security Groups
		If ($ObjectType -eq "GroupsOnly" -or $ObjectType -eq "All") {
			Get-DistributionGroup-Detail
			Get-RoomList-Detail
			# Gather all Unified (Office 365) Groups
			If ($Type -ne "On-Premises") {
				Get-UnifiedGroup-Detail
			}
		}

		Write-Progress -Activity "Processing $i of $($domains.count) total objects" -Status "Ready" -Completed
	}
	Write-Progress -Activity "Processing $p of $($domains.count) total domains" -Status "Ready" -Completed
    
}
#endregion

#region Remove Addresses
# If DestroyMethod equals "Destroy" it will remove the addresses it found for that domain using the CSV just discovered
If ($DestroyMethod -eq "Destroy") {
	# Get the list of accounts to target from CSV
	If ($outputFile) {
		$users = Import-Csv $outputFile -Delimiter "`t"
	} Else {
		Out-CMTraceLog -Logfile $logFileName -Message ("Error no CSV file was found...exiting") -Type Error
		Exit
	}
}

# If DestroyMethod equals "DestroyOnly" it will prompt for a CSV of accounts from a previous run
If ($DestroyMethod -eq "DestroyOnly") {
	$outputFile = Get-FileName $StartDirectory
	# Get the list of accounts to target from CSV
	If ($outputFile) {
		$users = Import-Csv $outputFile -Delimiter "`t"
	} Else {
		Out-CMTraceLog -Logfile $logFileName -Message ("Error no CSV file was found...exiting") -Type Error
		Exit
	}
}

# Grab the tenant address to build the replacement address if removing the primary address
If ($Type -eq "Office365") {
	$TenantName = Get-AzureADDomain | Where-Object { $_.IsInitial -eq $true }
}

$p = 0
ForEach ($user in $users) {
	# Progress Bar
	If ($users.count -gt 1) {
		$p++
		Write-Progress -Id 1 -Activity "Processing $p of $($users.count) total objects" -Status ("{0:P2}" -f ($p / $($users).Count)) -CurrentOperation $user.Alias
	}

	# If it is a UserMailbox/Shared/Room Mailbox remove the desired address
	If (($user.RecipientTypeDetails -eq "UserMailbox" -or $user.RecipientTypeDetails -eq 'SharedMailbox' -or $user.RecipientTypeDetails -eq "RoomMailbox") -and $user.IsInactiveMailbox -ne $True) {
		$address = $null
		$address = $user.EmailAddress.Split(":")
		If ($address.SyncRoot[0] -clike "SMTP") {
			Try {
				# Change the primary address first 
				$newPrimary = ($user.Alias + "@" + $TenantName.Name)
				Set-Mailbox -Identity $user.Alias -WindowsEmailAddress $newPrimary -ErrorAction Stop
				Try {
					# Remove the old primary which was added by default when the primrary address changed
					Set-Mailbox -Identity $user.Alias -EmailAddresses @{Remove = $address.SyncRoot[1] } -ErrorAction Stop
				} Catch {
					Out-CMTraceLog -Logfile $logFileName -Message ("Falied to remove " + $address + " from user " + $user.PrimarySMTPAddress) -Type Error
				}
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to change primary address to " + $newPrimary + " on mailbox " + $user.Alias) -Type Error
			}
		} Else {
			Try {
				# Remove the old primary which was added by default when the primrary address changed
				Set-Mailbox -Identity $user.Alias -EmailAddresses @{Remove = $address.SyncRoot[1] } -ErrorAction Stop
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to remove " + $address + " from user " + $user.PrimarySMTPAddress) -Type Error
			}
		}
	} Else {
		# Mailbox is considered inActive
		$address = $null
		$address = $user.EmailAddress.Split(":")
		If ($address.SyncRoot[0] -clike "SMTP") {
			Try {
				# Change the primary address first 
				$newPrimary = ($user.Alias + "@" + $TenantName.Name)
				Set-Mailbox -Identity $user.Alias -WindowsEmailAddress $newPrimary -InactiveMailbox:$True -ErrorAction Stop
				Try {
					# Remove the old primary which was added by default when the primrary address changed
					Set-Mailbox -Identity $user.Alias -InactiveMailbox:$True -EmailAddresses @{Remove = $address.SyncRoot[1] } -ErrorAction Stop
				} Catch {
					Out-CMTraceLog -Logfile $logFileName -Message ("Falied to remove " + $address + " from user " + $user.PrimarySMTPAddress) -Type Error
				}
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to change primary address to " + $newPrimary + " on mailbox " + $user.Alias) -Type Error
			}
		} Else {
			Try {
				# Remove the old primary which was added by default when the primrary address changed
				Set-Mailbox -Identity $user.Alias -InactiveMailbox:$True -EmailAddresses @{Remove = $address.SyncRoot[1] } -ErrorAction Stop
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to remove " + $address + " from user " + $user.PrimarySMTPAddress) -Type Error
			}
		}
	}	

	# If it is a mail user, remove the desired address
	If ($user.RecipientTypeDetails -eq "MailUser") {
		$address = $null
		$address = $user.EmailAddress.Split(":")
		If ($address.SyncRoot[0] -clike "SMTP") {
			Try {
				# Change the primary address first 
				$newPrimary = ($user.Alias + "@" + $TenantName.Name)
				Set-MailUser -Identity $user.Alias -WindowsEmailAddress $newPrimary -ErrorAction Stop
				Try {
					# Remove the old primary which was added by default when the primrary address changed
					Set-MailUser -Identity $user.Alias -EmailAddresses @{Remove = $address.SyncRoot[1] } -ErrorAction Stop
				} Catch {
					Out-CMTraceLog -Logfile $logFileName -Message ("Falied to remove " + $address + " from user " + $user.PrimarySMTPAddress) -Type Error
				}
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to change primary address to " + $newPrimary + " on mailbox " + $user.Alias) -Type Error
			}
		} Else {
			Try {
				# Remove the old primary which was added by default when the primrary address changed
				Set-MailUser -Identity $user.Alias -EmailAddresses @{Remove = $address.SyncRoot[1] } -ErrorAction Stop
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to remove " + $address + " from user " + $user.PrimarySMTPAddress) -Type Error
			}
		}
	}

	# If it is a standard group, remove the desired address
	If ($user.RecipientTypeDetails -eq 'MailUniversalDistributionGroup' -or $user.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') {
		$address = $null
		$address = $user.EmailAddress.Split(":")
		If ($address.SyncRoot[0] -clike "SMTP") {
			Try {
				# Change the primary address first 
				$newPrimary = ($user.Alias + "@" + $TenantName.Name)
				Set-DistributionGroup -Identity $user.Alias -WindowsEmailAddress $newPrimary -ErrorAction Stop
				Try {
					# Remove the old primary which was added by default when the primrary address changed
					Set-DistributionGroup -Identity $user.Alias -EmailAddresses @{Remove = $address.SyncRoot[1] } -ErrorAction Stop
				} Catch {
					Out-CMTraceLog -Logfile $logFileName -Message ("Falied to remove " + $address + " from group " + $user.PrimarySMTPAddress) -Type Error
				}
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to change primary address to " + $newPrimary + " on group " + $user.Alias) -Type Error
			}
		} Else {
			Try {
				# Remove the old primary which was added by default when the primrary address changed
				Set-DistributionGroup -Identity $user.Alias -EmailAddresses @{Remove = $address.SyncRoot[1] } -ErrorAction Stop
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to remove " + $address + " from group " + $user.PrimarySMTPAddress) -Type Error
			}
		}
	}
	
	# If it is a unified group, remove the desired address
	If ($user.RecipientTypeDetails -eq "GroupMailbox") {
		$address = $null
		$address = $user.EmailAddress.Split(":")
		If ($address.SyncRoot[0] -clike "SMTP") {
			Try {
				# Change the primary address first 
				$newPrimary = ($user.Alias + "@" + $TenantName.Name)
				Set-UnifiedGroup -Identity $user.Alias -PrimarySmtpAddress $newPrimary -ErrorAction Stop
				Try {
					# Remove the old primary which was added by default when the primrary address changed
					Set-UnifiedGroup -Identity $user.Alias -EmailAddresses @{Remove = $address.SyncRoot[1] } -ErrorAction Stop
				} Catch {
					Out-CMTraceLog -Logfile $logFileName -Message ("Falied to remove " + $address + " from group " + $user.PrimarySMTPAddress) -Type Error
				}
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to change primary address to " + $newPrimary + " on group " + $user.Alias) -Type Error
			}
		} Else {
			Try {
				# Remove the old primary which was added by default when the primrary address changed
				Set-UnifiedGroup -Identity $user.Alias -EmailAddresses @{Remove = $address.SyncRoot[1] } -ErrorAction Stop
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to remove " + $address + " from group " + $user.PrimarySMTPAddress) -Type Error
			}
		}
	}

	# If it is a roomlist, set new primary address
	If ($user.RecipientTypeDetails -eq "RoomList") {
		$address = $null
		$address = $user.EmailAddress.Split(":")
		If ($address.SyncRoot[0] -clike "SMTP") {
			Try {
				# Change the primary address first 
				$newPrimary = ($user.Alias + "@" + $TenantName.Name)
				Set-DistributionGroup -Identity $user.Alias -PrimarySmtpAddress $newPrimary -ErrorAction Stop
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to change primary address to " + $newPrimary + " on group " + $user.Alias) -Type Error
			}
		} Else {
			Try {
				# Remove the old primary which was added by default when the primrary address changed
				Set-DistributionGroup -Identity $user.Alias -EmailAddresses @{Remove = $address.SyncRoot[1] } -ErrorAction Stop
			} Catch {
				Out-CMTraceLog -Logfile $logFileName -Message ("Falied to remove " + $address + " from group " + $user.PrimarySMTPAddress) -Type Error
			}
		}
	}        
}
#endregion

#region Close down
# Close out main program
Get-PSSession | Remove-PSSession
#endregion