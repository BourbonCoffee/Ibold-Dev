#region Restore- Cmdlets

function Restore-InactiveMailbox {
    <#
    .SYNOPSIS
        This cmdlet will recover or restore an Inactive Mailbox.

    .DESCRIPTION
        This cmdlet can be used to recover or restore an Inactve Mailbox.
        Recovering an Inactive Mailbox consists of creating a new user and mailbox for the content to move into. The mailbox is moved and will no longer exist as an Inactive Mailbox.
        Restoring an Inactive Mailbox consists of copying the mail objects from the Inactive Mailbox into an existing mailbox and, optionally, a folder within that mailbox. A restore
            is a copy and the Inactrive Mailbox will still exist after the operation. The X500 address of the inactive mailbox is added to the EmailAddresses / proxyAddress of the 
            target mailbox to maintain reply-ability.
        
        This cmdlet uses parameter sets to separate the 'recover' and 'restore' processes. 'Recover' is the default parameter set.

    .PARAMETER SourceMailbox
        The source inactive mailbox you would like to recover or restore.

    .PARAMETER Prefix
        Add an optional prefix to the recovered inactive mailbox. Can be used with $Suffix

    .PARAMETER Suffix
        Add an optional suffix to the recovered inacitve mailbox. Can be used with $Prefix

    .PARAMETER TargetMailbox
        The target mailbox you would like to restore the mail objects into.
        The contents of the inactive mailbox (source mailbox) will be merged into the corresponding folders in the existing mailbox (target mailbox). If the folder does not exist
            at the source, it will be created.

    .PARAMETER TargetFolder
        The target folder, in the top-level folder structure of the target mailbox that the inactive mailbox (source mailbox) will be copied into.
        The folder structure of the inactive mailbox will be preserved within the TargetFolder.

    .INPUTS
        None. You cannot pipe objects into this function.

    .OUTPUTS
        Recover parameter set only: A CSV containing the DisplayName and PrimarySmtpAddress values of the recovered inactive mailbox(es) on the desktop.

    .EXAMPLE
        Recover Examples:
            Restore-InactiveMailbox -SourceMailbox bruce.banner@ibold.dev -Verbose
            Restore-InactiveMailbox -SourceMailbox tony.stark@ibold.dev -Prefix "Recovered - "
        From CSV or variable:
            foreach ($mbx in $csv) {Restore-InactiveMailbox -SourceMailbox $_.primarysmtpaddress}
    
    .EXAMPLE
        Restore Examples:
            Restore-InactiveMailbox -SourceMailbox miles.morales@ibold.dev -TargetMailbox peter.parker@ibold.dev
            Restore-InactiveMailbox -SourceMailbox miles.morales@ibold.dev -TargetMailbox peter.parker@ibold.dev -TargetFolder "Miles Morales Restored Mailbox"
        From CSV or variable:
            foreach ($mbx in $csv) {Restore-InactiveMailbox -SourceMailbox $mb.primarysmtpaddress -TargetMailbox $mb.TargetMailbox -TargetFolder $mb.DisplayName -Verbose}

    .LINK
    Recovering a mailbox: https://learn.microsoft.com/en-us/purview/recover-an-inactive-mailbox

    .LINK
    Restoring a mailbox: https://learn.microsoft.com/en-us/purview/restore-an-inactive-mailbox

        .NOTES
        Version:
            - 3.10.2024.1615:   -New function
            - 5.28.2024.2312:   -Logic fixes for when the DN between an InactiveMailbox and its SoftDeleted counterpart are different

    #>
    [CmdletBinding(DefaultParameterSetName = 'Recover')]
    param (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Inactive mailbox to be restored.", ParameterSetName = 'Recover')]
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Inactive mailbox to be restored.", ParameterSetName = 'Restore')]
        [string]$SourceMailbox,

        [Parameter(Mandatory = $false, HelpMessage = "Add a prefix to the Display Name of the recovered inactive mailbox", ParameterSetName = 'Recover')]
        [string]$Prefix,

        [Parameter(Mandatory = $false, HelpMessage = "Add a suffix to the Display Name of the recovered inactive mailbox", ParameterSetName = 'Recover')]
        [string]$Suffix,

        [Parameter(Mandatory = $true, HelpMessage = "Target mailbox for restored items.", ParameterSetName = 'Restore')]
        [string]$TargetMailbox,

        [Parameter(Mandatory = $false, HelpMessage = "Target folder for restored items.", ParameterSetName = 'Restore')]
        [string]$TargetFolder
    )


    begin {
        $functionName = "$($PSCmdlet.MyInvocation.MyCommand.Name)"
        $startTime = Get-Date
        Write-Verbose @"
    `r`n  Function: $functionName
    Starting at $($startTime.ToString('yyyy-MM-dd hh:mm:ss tt')) `n
"@ #| Out-Log
    }#begin

    process {
        if ($PSCmdlet.ParameterSetName -eq 'Recover') {
            $characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!@#$%^&*()"
            $results = @()

            Write-Verbose "Inactive mailbox RECOVER to a new mailbox is in progress..." #| Out-Log
                
            try {
                Write-Host "Creating new mailbox for $SourceMailbox" -ForegroundColor Green #| Out-Log

                $inactiveMailbox = $null
                $randomNumber = $null
                
                Write-Verbose "Retrieving inactive mailbox for $SourceMailbox" #| Out-Log
                $inactiveMailbox = Get-Mailbox -InactiveMailboxOnly -Identity $SourceMailbox
                $randomNumber = Get-Random -Minimum 100 -Maximum 999 #Generate a random three-digit number to make sure -Name parameter is unique
                
                $uniqueName = $inactiveMailbox.Alias + $randomNumber
                $displayName = $inactiveMailbox.DisplayName
                $primarySMTPAddress = $inactiveMailbox.MicrosoftOnlineServicesID

                # Create a secure password
                Write-Verbose "Generating secure password for $primarySMTPAddress." #| Out-Log
                $password = -join (1..16 | ForEach-Object { Get-Random -InputObject $characters.ToCharArray() })
                $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
                
                # Splat params
                $recoverParams = @{
                    InactiveMailbox           = $primarySMTPAddress
                    Name                      = $uniqueName
                    DisplayName               = "{0}{1}{2}" -f $Prefix, $displayName, $Suffix
                    MicrosoftOnlineServicesID = $primarySMTPAddress
                    Password                  = $securePassword
                    ResetPasswordOnNextLogon  = $false
                }

                # Create new mailbox
                Write-Verbose "Creating new mailbox for $primarySMTPAddress" #| Out-Log
                New-Mailbox @recoverParams
                Write-Verbose "$primarySMTPAddress recovered. `n" #| Out-Log

                # Add details to the results array
                $results += [PSCustomObject]@{
                    DisplayName        = "{0}{1}{2}" -f $Prefix, $displayName, $Suffix
                    PrimarySmtpAddress = $primarySMTPAddress
                }#PSCustomObject
            }#try
            catch {
                Write-Error "An error occurred with mailbox: $($inactiveMailbox.PrimarySMTPAddress). Error: $_" #| Out-Log
            }#catch       

            $path = Get-KnownFolderPath "Desktop"
            $ExportToCSV = "$path\RecoveredMailboxes.csv"
            Write-Verbose "Exporting recovered mailboxes to $path" #| Out-Log

            try {
                # Export the results to CSV
                $results | Export-Csv -Path $ExportToCSV -NoTypeInformation -Append
                Write-Verbose "Exported mailbox recovery details to $ExportToCSV `n" #| Out-Log
                Write-Host "Done. Exported to: $ExportToCSV" -ForegroundColor Green
            } catch {
                Write-Error "An error occurred while exporting to CSV. Error: $_" #| Out-Log
            }#catch
        }#if
        elseif ($PSCmdlet.ParameterSetName -eq 'Restore') {

            Write-Verbose "Inactive mailbox RESTORE to existing mailbox is in progress..." #| Out-Log

            try {
                # Variables
                $inactiveMailbox = $null

                Write-Verbose "Retrieving inactive mailbox for $SourceMailbox" #| Out-Log
                $inactiveMailbox = Get-Mailbox -InactiveMailboxOnly -Identity $SourceMailbox

                # Splat params
                $restoreParams = @{
                    SourceMailbox = $inactiveMailbox.DistinguishedName
                    TargetMailbox = $TargetMailbox
                }

                # Check if the TargetFolder parameter is provided. If it is, add it to the parameter hashtable.
                if (![string]::IsNullOrWhiteSpace($TargetFolder)) {
                    $restoreParams.TargetRootFolder = $TargetFolder
                    Write-Verbose "Mail objects will be restored to: $($TargetFolder) in mailbox: $($TargetMailbox)" #| Out-Log
                }

                Write-Verbose "Adding the LegacyExchangeDN of the inactive mailbox as an X500 proxy address to the target mailbox. `n" #| Out-Log
                Write-Host "Starting mailbox restore for $SourceMailbox" -ForegroundColor Green #| Out-Log
                Set-Mailbox $TargetMailbox -EmailAddresses @{Add = "X500:" + $($inactiveMailbox.LegacyExchangeDN) }
                Write-Verbose "Submitting request to restore and merge $($SourceMailbox) into $($TargetMailbox) `n" #| Out-Log
                New-MailboxRestoreRequest @restoreParams
            }#try
            catch {
                Write-Error "An error occurred while submitting the mailbox restore request. Error: $_" #| Out-Log
            }#catch
        }#elseif
    }#process

    end {
        $endTime = Get-Date 
        $runTime = $endTime - $startTime 
        # Format the output time
        if ($runTime.TotalSeconds -lt 1)
        { $elapsed = "$($runTime.TotalMilliseconds.ToString('#,0.0000')) Milliseconds" }
        elseif ($runTime.TotalSeconds -gt 60)
        { $elapsed = "$($runTime.TotalMinutes.ToString('#,0.0000')) Minutes" }
        else
        { $elapsed = "$($runTime.TotalSeconds.ToString('#,0.0000')) Seconds" }

        Write-Verbose @"
    `r`n  Function: $functionName
    Finished at $($endTime.ToString('yyyy-MM-dd hh:mm:ss tt'))
    Elapsed Time $elapsed `n
"@ #| Out-Log
    }#end
}#function Restore-InactiveMailbox

#endregion

#region Export Module Members

Export-ModuleMember -Function Restore-InactiveMailbox

#endregion