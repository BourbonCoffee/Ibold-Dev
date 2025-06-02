Param (
    [Parameter(Mandatory = $true)]
    [string]$InputCsv,
    
    [Parameter(Mandatory = $true)]
    [string]$Prefix,

    [Parameter(Mandatory = $true,
        HelpMessage = "Format: https://tenantname.sharepoint.com")]
    [string]$DestinationSharePointUrl,

    [Parameter(Mandatory = $false,
        HelpMessage = "Destination service account credential object from Get-Credential")]
    [System.Management.Automation.PSCredential]$DestinationCredential
)

#Check for PnP.PowerShell module, install if not found
if ((Get-Module PnP.PowerShell -ListAvailable) -eq $null) {
    Install-Module -Name PnP.PowerShell -Scope CurrentUser
}

Import-Module PnP.PowerShell
Connect-PnPOnline -Url $DestinationSharePointUrl -Interactive

$inputFile = Import-Csv -LiteralPath $InputCsv.trim('"')
$outputFile = "C:\Users\CIbold0\OneDrive - CBIZ, Inc\Desktop\Chris\CDI\SPO\SPOSitesMapping.csv"

$siteTemplateHash = @{
    "GROUP#0"              = "TeamSite"
    "SITEPAGEPUBLISHING#0" = "CommunicationSite"
    "STS#3"                = "TeamSiteWithoutMicrosoft365Group"
}

$allDstSites = Get-PnPTenantSite -Connection $dstTenant
$report = @()

foreach ($site in $inputFile) {
    $alias = $site.Url.Split('/')[-1]
    $siteUrl = "/sites/$Prefix$alias"

    if ($site.Template -eq "GROUP#0") {
        $newSPOSite = New-PnPSite -Type TeamSite -Title $site.Title -Alias "$Prefix$alias" -Connection $dstTenant
    } else {
        
        $HashArguments = @{
            Template                 = $site.Template
            Title                    = $site.Title
            Url                      = $siteUrl
            Owner                    = "svcCDIQuestMigration@cbizcorp.onmicrosoft.com"
            StorageQuota             = $site.StorageUsageCurrent * 1.2
            StorageQuotaWarningLevel = ($site.StorageUsageCurrent * 1.2) * 0.8
        }

        New-PnPTenantSite @HashArguments -Connection $dstTenant -TimeZone 10
    }

    $item = [PSCustomObject]@{SourceUrl = $site.Url; DestinationUrl = $DestinationSharePointUrl + $siteUrl }
    $item
    $report += $item
}

$report | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter ","