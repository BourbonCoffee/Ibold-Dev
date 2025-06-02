Import-Csv -Path "C:\Users\ChrisIbold\OneDrive - Sterling Consulting\Desktop\MailboxesWithoutOnMicrosoft.csv" | ForEach-Object {
    $Recipient = $_.PrimarySmtpAddress
    $ProxyAddresses = $_.AliasToAdd
    try {
        Set-Mailbox -Identity $Recipient -EmailAddresses @{add = $ProxyAddresses }
        Write-Host "$ProxyAddresses has been added to $Recipient"
    } catch {
        $executionError = $Error
        Write-Host $executionError
    }
}