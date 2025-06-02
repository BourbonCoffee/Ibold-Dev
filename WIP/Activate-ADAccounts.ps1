# Define the path to the input CSV file
$inputCsvPath = "$([Environment]::GetFolderPath('Desktop'))\UserList.csv"

# Define the path to the output CSV file
$outputCsvPath = "$([Environment]::GetFolderPath('Desktop'))\activation_results.csv"

# Import the Active Directory module
Import-Module ActiveDirectory

# Read the input CSV file
$inputCSV = Import-Csv $inputCSVPath

# Initialize an empty array to store the output data
$outputData = @()

# Iterate through each row in the input CSV file
foreach ($row in $inputCSV) {
    $upn = $row.UPN

    # Get the deactivated account
    $deactivatedAccount = Get-ADUser -Filter "UserPrincipalName -eq '$upn' -and Enabled -eq '$false'"

    if ($deactivatedAccount) {
        # Activate the account
        $deactivatedAccount | Enable-ADAccount

        $outputData += [PSCustomObject] @{
            UPN    = $upn
            Status = "Activated"
        }

        Write-Host "Account '$upn' has been activated."
    }
    else {
        $outputData += [PSCustomObject] @{
            UPN    = $upn
            Status = "Not Found or Already Active"
        }

        Write-Host "Account '$upn' not found or already active."
    }
}

# Export the output data to a CSV file
$outputData | Export-Csv -Path $outputCSVPath -NoTypeInformation

Write-Host "Results have been exported to '$outputCSVPath'."
