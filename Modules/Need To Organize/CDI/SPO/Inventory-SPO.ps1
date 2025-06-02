Param (
    [Parameter(Mandatory = $true,
        HelpMessage = "Format: https://tenantname.sharepoint.com")]
    [string]$SourceSharePointUrl,

    [Parameter(Mandatory = $false,
        HelpMessage = "Destination service account credential object from Get-Credential")]
    [System.Management.Automation.PSCredential]$SourceCredential
)

#Check for PnP.PowerShell module, install if not found
if ($null -eq (Get-Module PnP.PowerShell -ListAvailable)) {
    Install-Module -Name PnP.PowerShell -Scope CurrentUser
}
Import-Module PnP.PowerShell

Write-Host "Connecting to source tenant..."
$srcTenant = Connect-PnPOnline -Url $SourceSharePointUrl -Interactive -ReturnConnection

$report = @()

$sites = Get-PnPTenantSite -Detailed -IncludeOneDriveSites -Connection $srcTenant

# We do not want to migrate sites based on these site templates
$templateRegEx = "^(STS#-1|RedirectSite#|SPSPERS#|SRCHCEN#|APPCATALOG#|TEAMCHANNEL#|PWA#|SPSMSITEHOST#|POINTPUBLISHINGTOPIC#|POINTPUBLISHINGHUB#)\d+$"
$filteredSites = $sites | Where-Object { $_.Template -notmatch $templateRegEx -and $_.IsTeamsConnected -eq $false }

$filteredSites | ForEach-Object {
    $_.Title
    $myObject = [PSCustomObject]@{
        Title                   = $_.Title
        Url                     = $_.Url
        LastContentModifiedDate = $_.LastContentModifiedDate
        Status                  = $_.Status
        IsOneDrive              = $_.Url.Contains("/personal/")
        IsHubSite               = $_.IsHubSite
        Owner                   = $_.Owner
        Template                = $_.Template
        StorageUsageCurrent     = $_.StorageUsageCurrent
        TeamsChannelType        = $_.TeamsChannelType
        IsTeamsConnected        = $_.IsTeamsConnected
        IsTeamsChannelConnected = $_.IsTeamsChannelConnected
    }
    $report += $myObject
}

$report | Export-Csv "C:\Users\CIbold0\OneDrive - CBIZ, Inc\Desktop\Chris\CDI\SPO\SPOSites.csv" -NoTypeInformation