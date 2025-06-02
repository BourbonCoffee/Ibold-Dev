#Change the file locations within lines 7 & 8 to reflect the source (UPN mapping csv) and the output which provides the OneDrive URLs
Write-Host "Connecting to source tenant..."
$srcTenant = Connect-Site -Url https://SRCTENANT-admin.sharepoint.com -Browser -DisableSSO
Write-Host "Connecting to destination tenant..."
$dstTenant = Connect-Site -Url https://DSTTENANT-admin.sharepoint.com -Browser -DisableSSO

$inputFile = "Location of source UPN mapping file"
$outputFile = "Location of output file\OneDriveSiteMappingALL.csv"

$table = Import-Csv $inputFile -Delimiter ","
$first = $true
foreach ($row in $table) {
    try {
        $sourceEmail = $row.SourceUPN 
        $sourceUrl = Get-OneDriveUrl -Tenant $srcTenant -Email $sourceEmail

        $targetEmail = $row.TargetUPN 
        $targetUrl = Get-OneDriveUrl -Tenant $dstTenant -Email $targetEmail -ProvisionIfRequired

        $item = [PSCustomObject]@{Username = ''; SourceEmail = $sourceEmail; SourceUrl = $sourceUrl; DestinationEmail = $targetEmail; DestinationUrl = $targetUrl }

        if ($first) {
            $item | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter "," 
            $first = $false
        }
        else {
            $item | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter "," -Append
        }
     
        Write-Output $item
    }
    catch {
        Write-Output "ERROR: $($_.Exception.Message)"
    }
}