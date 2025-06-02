Connect-ExchangeOnline

$hiddenGroups = 0

# Log will output to a text file on the desktop
$logFile = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop', 'HideGroupsLog.txt')
Add-Content -Path $logFile -Value "Script started: $(Get-Date)"
$message = "Finding team-enabled Microsoft 365 Groups and checking for any which are visible to Exchange clients."
Write-Host $message
Add-Content -Path $logFile -Value $message

try {

    $groups = Get-UnifiedGroup -ResultSize Unlimited
    
    # Filter the groups to only include those that are visible to Exchange clients
    $groups = $groups | Where-Object {$_.HiddenFromExchangeClientsEnabled -eq $False}

    # Check if any groups need to be hidden from Exchange clients and hide from GAL for non-Teams enabled groups
    if ($groups.Count -ne 0) {
        foreach ($group in $groups) {
            $message = "Hiding $($group.DisplayName)"
            Write-Host $message
            Add-Content -Path $logFile -Value $message
            
            $hiddenGroups++
            # Set group based off Entra ID rather than Exchange GUID to remain consistent
            Set-UnifiedGroup -Identity $group.ExternalDirectoryObjectId -HiddenFromExchangeClientsEnabled:$True -HiddenFromAddressListsEnabled:$True
        }
    } else {
        $message = "No team-enabled Microsoft 365 groups are visible to Exchange clients and address lists."
        Write-Host $message
        Add-Content -Path $logFile -Value $message
    }

    # Display and log a summary of the groups that were hidden
    $message = "All done. $hiddenGroups team-enabled groups hidden from Exchange clients."
    Write-Host $message -ForegroundColor Green
    Add-Content -Path $logFile -Value $message
} catch {
    $errorMessage = "An error occurred: $($_.Exception.Message)"
    Write-Host $errorMessage -ForegroundColor Red
    Add-Content -Path $logFile -Value $errorMessage
}

Add-Content -Path $logFile -Value "Script ended: $(Get-Date)"
