# Install the PnP PowerShell module if not already installed
# Install-Module -Name PnP.PowerShell -Force

$urls = "$([Environment]::GetFolderPath('Desktop'))\urls.txt"

$output = "$([Environment]::GetFolderPath('Desktop'))\missingOwners.csv"
$results = @()

Connect-PnPOnline -Url "https://cbizcorp.sharepoint.com" -UseWebLogin

foreach ($siteUrl in Get-Content $urls) {
    Write-Host "Checking site: $siteUrl"
    try {
        $site = Get-PnPTenantSite -Url $siteUrl -ErrorAction Stop

        if ($site.Owner -eq $null) {
            Write-Host "Site does not have an owner!"
            
            $results += [PSCustomObject]@{
                'Site Title' = $site.Title
                'URL' = $site.Url
                'Has Owner' = $false
                'Owner' = "MISSING!"
                'Error' = $null
            }
        } else {
            Write-Host "Site has an owner: $($site.Owner)"
            $results += [PSCustomObject]@{
                'Site Title' = $site.Title
                'URL' = $site.Url
                'Has Owner' = $true
                'Owner' = $site.Owner
                'Error' = $null
            }
        }
    } catch {
        Write-Host "Error checking site: $siteUrl - $_"
        $results += [PSCustomObject]@{
            'Site Title' = $null
            'URL' = $siteUrl
            'Has Owner' = $false
            'Owner' = $null
            'Error' = $_.Exception.Message
        }
    }
}
$results | Export-Csv -Path $output -NoTypeInformation
