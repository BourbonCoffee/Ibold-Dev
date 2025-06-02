$LogFile = "C:\Users\svcEBKMigration02\Downloads\Development\Development\OneDrive\OneDriveSites.log"

$tenant = Connect-Site -Url "https://cbizcorp-admin.sharepoint.com" -Browser
$csvFile = "C:\Users\svcEBKMigration02\Downloads\Development\Development\OneDrive\fred.csv"
$table = Import-Csv $csvFile -Delimiter ","
$first = $true

    foreach ($row in $table) {
        try {
            $targetEmail = $row.Email
            $targetUrl = Get-OneDriveUrl -Tenant $tenant -Email $targetEmail -ProvisionIfRequired

            if ($targetUrl) {
                $item = [PSCustomObject]@{DestinationEmail = $targetEmail; DestinationUrl = $targetUrl }

                if ($first) {
                    $item | Export-Csv -Path "C:\Users\svcEBKMigration02\Downloads\Development\Development\OneDrive\targetURLs-1.csv" -NoTypeInformation -Delimiter "," 
                    $first = $false
                }
                else {
                    $item | Export-Csv -Path "C:\Users\svcEBKMigration02\Downloads\Development\Development\OneDrive\targetURLs-1.csv" -NoTypeInformation -Delimiter "," -Append
                }
            
                Write-Output $item
            } else {
                Write-Warning "No URL found for $($targetEmail)" | Out-File $LogFile -Force
            }
        }
        catch {
            Write-Output "ERROR: $($_.Exception.Message)" | Out-File $LogFile -Force
        }
    }