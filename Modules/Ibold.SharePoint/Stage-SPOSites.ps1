[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true)]
    [string]$srcTenantUrl,

    [Parameter(Mandatory = $true)]
    [string]$dstTenantUrl,

    [Parameter(Mandatory = $true)]
    [string]$sitePrefix,

    [Parameter(Mandatory = $true)]
    [string]$filePath
)

Import-Module -Name Microsoft.Online.SharePoint.PowerShell
Import-Module -Name PnP.PowerShell
 
$srcTenantAdminUrl = "https://" + $srcTenantUrl.Replace("https://","").Replace(".sharepoint.com","") + "-admin.sharepoint.com"
$dstTenantAdminUrl = "https://" + $dstTenantUrl.Replace("https://","").Replace(".sharepoint.com","") + "-admin.sharepoint.com"

Write-Host "Connecting to source tenant..."
$srcTenant = Connect-SPOService -Url $srcTenantAdminUrl 
Write-Host "Connecting to destination tenant..."
$dstTenant = Connect-PnPOnline -Url $dstTenantAdminUrl -Interactive
 
$allDstSites = Get-PnPTenantSite -Connection $dstTenant
 
$outputFile = "$([Environment]::GetFolderPath('Desktop'))\SPOSitesMapping.csv"

$siteTemplateHash = @{
    "GROUP#0" = "TeamSite";
    "SITEPAGEPUBLISHING#0" = "CommunicationSite";
    "STS#3" = "TeamSiteWithoutMicrosoft365Group"
}

Get-Content $filePath | ForEach-Object {
    $site = $_

    Write-Host "Processing site $site..."

    $spoSite = Get-SPOSite -Identity "$site"
    $newSPOSiteName = "$sitePrefix $($spoSite.Title)"
    $newSPOUrl = "$dstTenantUrl/sites/$sitePrefix" + $spoSite.Url.Split('/')[-1]
    $alias = "$sitePrefix" + $spoSite.Url.Split('/')[-1]

    if (-not $spoSite.IsTeamsConnected) {
        $newSPOSite = $allDstSites | Where-Object { $_.Url -eq $newSPOUrl } | Select-Object -ExpandProperty Url

        if ($newSPOSite -eq $null) {
            Write-Host "Creating $newSPOSite..."
            $siteTemplate = $siteTemplateHash[$spoSite.Template]

            switch ($siteTemplate) {
                "TeamSite" {
                    $newSPOSite = New-PnPSite -Type TeamSite -Title $newSPOSiteName -Alias $alias -Connection $dstTenant #-Url $newSPOUrl
                    break
                }
                "TeamSiteWithoutMicrosoft365Group" {
                    $newSPOSite = New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title $newSPOSiteName -Url $newSPOUrl -Connection $dstTenant
                    break
                }
                "CommunicationSite" {
                    $newSPOSite = New-PnPSite -Type CommunicationSite -Title $newSPOSiteName -Url $newSPOUrl -Connection $dstTenant
                    break
                }
                default {
                    #-Owner switch is currently hard coded. See if there is a way to pull UPN from PnP Connection...
                    $newSPOSite = New-PnPTenantSite -template $spoSite.Template -Title $newSPOSiteName -Url $newSPOUrl -Connection $dstTenant -Owner "" -TimeZone 11 -StorageQuota ($spoSite.StorageQuota + 500)
                    break
                }
            }
        }
    }
    else {
        Write-Warning "'$site' should be a Team"
    }

    [PSCustomObject]@{
        SourceUrl = $site; DestinationUrl = $newSPOSite 
    } | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter "," -Append
}
