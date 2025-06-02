Connect-MgGraph -Scopes "User.Read.All"

# Define the properties we want to retrieve
$select = @(
    "DisplayName",
    "Mail",
    "EmployeeId",
    "EmployeeNumber"
)
# Add extension attributes to the select list
1..15 | ForEach-Object { $select += "onPremisesExtensionAttributes" }

# Get all users with specified properties
$users = Get-MgUser -All -Select $select -ConsistencyLevel eventual

# Create an array to store the results
$results = @()

# Process each user
foreach ($user in $users) {
    # Create a hashtable for the current user's properties
    $userProperties = @{
        DisplayName    = $user.DisplayName
        Email          = $user.Mail
        EmployeeID     = $user.EmployeeId
        EmployeeNumber = $user.EmployeeNumber
    }

    # Add extension attributes 1-15 to the hashtable
    if ($user.AdditionalProperties.onPremisesExtensionAttributes) {
        $extensionAttrs = $user.AdditionalProperties.onPremisesExtensionAttributes
        for ($i = 1; $i -le 15; $i++) {
            $extensionKey = "extensionAttribute$i"
            $userProperties[$extensionKey] = $extensionAttrs[$extensionKey]
        }
    }

    # Create a custom object from the hashtable and add it to results
    $results += [PSCustomObject]$userProperties
}

# Export the results to CSV
$exportPath = "C:\temp\UserAttributes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $exportPath -NoTypeInformation

Write-Host "Export completed. File saved as: $exportPath"
