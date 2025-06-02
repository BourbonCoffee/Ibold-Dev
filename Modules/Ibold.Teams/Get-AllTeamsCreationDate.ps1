# This script will get all the teams in a tenant and their creation date.

Import-Module ExchangeOnline
Import-Module MicrosoftTeams

# Connect to Teams and EXO powershell
Connect-MicrosoftTeams 
Connect-ExchangeOnline

Write-Host "Getting Teams..."
$Teams = Get-Team

$teamdata = @()

Write-Host "Getting UnifiedGroup data..."
foreach($Team in $Teams)
{
	$TeamUG = Get-UnifiedGroup -Identity $Team.GroupId
	$teamdata += @(
		[pscustomobject]@{
		DisplayName = $Team.DisplayName
		CreationDate = $TeamUG.WhenCreated
		}
	)
}

# Display results
$teamdata | sort displayname | Export-csv -path "$env:OneDrive\Desktop\TeamsCreationDate.csv" -NotypeInformation