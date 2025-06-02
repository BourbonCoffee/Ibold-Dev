$aliases = Import-Csv -Path "C:\Users\CIbold0\OneDrive - CBIZ, Inc\Desktop\Chris\Somerset\data\Somerset CBIZ Aliases.csv"

$missingAliases = @()

foreach ($alias in $aliases) {
    $recipient = Get-Recipient -Filter "Alias -eq '$($alias.AliasNoPrefix)'"

    if (-not $recipient) {
        $missingAliases += [PSCustomObject]@{
            DisplayName        = $alias.DisplayName
            AliasNoPrefix      = $alias.AliasNoPrefix
            PrimarySmtpAddress = $alias.PrimarySmtpAddress
        }
    }
}

# Export the results to a new CSV file
$missingAliases | Export-Csv -Path "C:\Users\CIbold0\OneDrive - CBIZ, Inc\Desktop\Chris\Somerset\data\Somerset Missing Aliases.csv" -NoTypeInformation
