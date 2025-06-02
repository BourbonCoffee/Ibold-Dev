# Parameters
$TenantAdminURL = "https://ibolddev-Admin.SharePoint.com"
$CSVPath = "C:\Temp\AccessRequestData.csv"

# Function to Get access request Configuration for a SharePoint Online site
Function Get-PnPAccessRequestConfig {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)] $Web
    )
  
    Try {
        Write-Host -ForegroundColor Yellow "Getting Access Request Settings on: $($web.Url)"
        
        if ($Web.HasUniqueRoleAssignments) {
            if ($Web.RequestAccessEmail -ne [string]::Empty) {
                $AccessRequest = "Enabled"
                $EmailOrGroup = "Email"
                $AccessRequestConfig = $Web.RequestAccessEmail
            } elseif ($Web.UseAccessRequestDefault -eq $true) {
                $AccessRequest = "Enabled"
                $EmailOrGroup = "Default Owner Group"
                $OwnersGroup = Get-PnPGroup -AssociatedOwnerGroup
                $AccessRequestConfig = $OwnersGroup.Title
            } else {
                $AccessRequest = "Disabled"
                $EmailOrGroup = [string]::Empty
                $AccessRequestConfig = [string]::Empty
            }
        } else {
            $AccessRequest = "Inherits from Parent"
            $EmailOrGroup = [string]::Empty
            $AccessRequestConfig = [string]::Empty
        }
 
        # Return an object instead of modifying a global variable
        [PSCustomObject]@{
            WebURL              = $Web.URL
            AccessRequest       = $AccessRequest
            EmailOrGroup        = $EmailOrGroup
            AccessRequestConfig = $AccessRequestConfig
        }
    } Catch {
        Write-Host "Error Getting Access Requests for $($Web.URL): $($_.Exception.Message)" -ForegroundColor Red
        $null  # Return null to handle the error gracefully
    }
}
  
# Ensure the directory exists
$CSVDirectory = Split-Path -Path $CSVPath -Parent
if (-not (Test-Path -Path $CSVDirectory)) {
    New-Item -ItemType Directory -Path $CSVDirectory | Out-Null
}

# Remove existing CSV if it exists
if (Test-Path $CSVPath) { 
    Remove-Item $CSVPath -Force 
}

# Connect to Admin Center (ensure this works)
try {
    Connect-PnPOnline -Url $TenantAdminURL -Interactive -ClientId 455cb93e-76f9-4337-93b9-a7699566c823
} catch {
    Write-Host "Error connecting to SharePoint Admin Center: $($_.Exception.Message)" -ForegroundColor Red
    exit
}
 
# Get All Site collections - Exclude specific site templates
try {
    $SitesCollections = Get-PnPTenantSite | 
        Where-Object { $_.Template -notin ("SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1") }
} catch {
    Write-Host "Error retrieving tenant sites: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Collect all access request data
$AccessRequestData = @()
   
# Loop through each site collection
foreach ($Site in $SitesCollections) {
    try {
        # Connect to site collection
        Connect-PnPOnline -Url $Site.Url -Interactive -ClientId 455cb93e-76f9-4337-93b9-a7699566c823
  
        # Collect data for Root Web and all Subwebs
        $SiteAccessData = Get-PnPSubWeb -IncludeRootWeb -Recurse -Includes HasUniqueRoleAssignments, RequestAccessEmail, UseAccessRequestDefault | 
            ForEach-Object { Get-PnPAccessRequestConfig $_ } |
            Where-Object { $_ -ne $null }
        
        $AccessRequestData += $SiteAccessData
    } catch {
        Write-Host "Error processing site $($Site.Url): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Export collected data to CSV once at the end
try {
    $AccessRequestData | Export-Csv -Path $CSVPath -NoTypeInformation
    Write-Host "Data exported successfully to $CSVPath" -ForegroundColor Green
} catch {
    Write-Host "Error exporting data to CSV: $($_.Exception.Message)" -ForegroundColor Red
}