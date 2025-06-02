<#
.SYNOPSIS
    Adds a specified user (or list of users) as owner to all public Microsoft Teams, with optional support for private teams.

.DESCRIPTION
    Connects to Microsoft Teams, loops through all teams, and adds each user as an owner to public teams. Optionally includes private teams.
    Supports WhatIf dry-run mode and logs actions to a CSV on the user's desktop.

.PARAMETER UserPrincipalName
    The UPN (email) of a single user to be added as owner.

.PARAMETER UserListPath
    Path to a text file containing one UPN per line. Used for batch processing.

.PARAMETER Private
    If specified, the script includes private teams.

.PARAMETER ConfirmPrivate
    If $true, the script skips prompting for each private team. If $false (default), prompts for confirmation on each private team.

.PARAMETER WhatIf
    Simulates the operation without making changes.

.EXAMPLE
    Add-TeamsOwner -UserListPath ".\users.txt" -Private -Confirm:$true -WhatIf

.NOTES
    Requires MicrosoftTeams module.
#>

function Add-TeamsOwner {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param (
        [Parameter(Mandatory = $false)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory = $false)]
        [string]$UserListPath,

        [switch]$Private,

        [Parameter(Mandatory = $false)]
        [bool]$ConfirmPrivate = $false
    )

    # Prep logging
    $Desktop = [Environment]::GetFolderPath('Desktop')
    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $LogPath = Join-Path $Desktop "TeamsOwnerLog_$Timestamp.csv"
    $Log = @()

    # Check Teams module
    if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
        Write-Error "MicrosoftTeams module not found. Run: Install-Module MicrosoftTeams"
        return
    }

    # Connect if needed
    try {
        Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
        Connect-MicrosoftTeams | Out-Null
        Write-Host "Connected!" -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect to Microsoft Teams: $($_.Exception.Message)"
        return
    }

    # Load UPNs
    $UPNs = @()
    if ($UserPrincipalName) { $UPNs += $UserPrincipalName }

    if ($UserListPath) {
        if (-not (Test-Path $UserListPath)) {
            Write-Error "User list file not found: $UserListPath"
            return
        }
        $FileUPNs = Get-Content $UserListPath | Where-Object { $_ -match '@' } | ForEach-Object { $_.Trim() }
        $UPNs += $FileUPNs
    }

    if (-not $UPNs.Count) {
        Write-Warning "No UPNs provided. Use -UserPrincipalName or -UserListPath."
        return
    }

    # Get all Teams
    $AllTeams = Get-Team
    Write-Host "Found $($AllTeams.Count) teams." -ForegroundColor Cyan

    foreach ($UPN in $UPNs) {
        Write-Host "`nProcessing user: $UPN" -ForegroundColor Yellow

        foreach ($Team in $AllTeams) {
            $TeamName = $Team.DisplayName
            $TeamId = $Team.GroupId
            $IsPrivate = $Team.Visibility -eq "Private"

            if ($IsPrivate -and -not $Private) {
                $Log += [pscustomobject]@{
                    UserUPN   = $UPN
                    TeamName  = $TeamName
                    TeamId    = $TeamId
                    Action    = "Skipped"
                    Status    = "Private team (not selected)"
                    Timestamp = (Get-Date)
                }
                continue
            }

            if ($IsPrivate -and $Private -and -not $ConfirmPrivate) {
                $Answer = Read-Host "Add $UPN as OWNER to private team '$TeamName'? (Y/N)"
                if ($Answer -notin @("Y", "y")) {
                    Write-Host "Skipped." -ForegroundColor DarkGray
                    $Log += [pscustomobject]@{
                        UserUPN   = $UPN
                        TeamName  = $TeamName
                        TeamId    = $TeamId
                        Action    = "Skipped"
                        Status    = "Declined in prompt"
                        Timestamp = (Get-Date)
                    }
                    continue
                }
            }

            try {
                $Members = Get-TeamUser -GroupId $TeamId
                $UserEntry = $Members | Where-Object { $_.User -eq $UPN }

                if (-not $UserEntry) {
                    if ($PSCmdlet.ShouldProcess("$TeamName", "Add $UPN as OWNER")) {
                        Add-TeamUser -GroupId $TeamId -User $UPN -Role Owner
                    }
                    $Log += [pscustomobject]@{
                        UserUPN   = $UPN
                        TeamName  = $TeamName
                        TeamId    = $TeamId
                        Action    = "Added"
                        Status    = "User added as owner"
                        Timestamp = (Get-Date)
                    }
                } elseif ($UserEntry.Role -ne "Owner") {
                    if ($PSCmdlet.ShouldProcess("$TeamName", "Promote $UPN to OWNER")) {
                        Add-TeamUser -GroupId $TeamId -User $UPN -Role Owner
                    }
                    $Log += [pscustomobject]@{
                        UserUPN   = $UPN
                        TeamName  = $TeamName
                        TeamId    = $TeamId
                        Action    = "Promoted"
                        Status    = "User promoted to owner"
                        Timestamp = (Get-Date)
                    }
                } else {
                    Write-Verbose "Already an owner: $TeamName"
                    $Log += [pscustomobject]@{
                        UserUPN   = $UPN
                        TeamName  = $TeamName
                        TeamId    = $TeamId
                        Action    = "AlreadyOwner"
                        Status    = "No action needed"
                        Timestamp = (Get-Date)
                    }
                }
            } catch {
                $Log += [pscustomobject]@{
                    UserUPN   = $UPN
                    TeamName  = $TeamName
                    TeamId    = $TeamId
                    Action    = "Error"
                    Status    = $_.Exception.Message
                    Timestamp = (Get-Date)
                }
                Write-Warning "Error in '$TeamName': $($_.Exception.Message)"
            }
        }
    }

    # Write log
    try {
        $Log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
        Write-Host "`n📄 Log saved to: $LogPath" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to write log file: $($_.Exception.Message)"
    }

    Write-Host "`n✔️ Script complete." -ForegroundColor Green
}
