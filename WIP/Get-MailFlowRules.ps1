$path = "$([Environment]::GetFolderPath('Desktop'))"
Connect-ExchangeOnline

$mailFlowRules = Get-TransportRule
$mailFlowRules | Export-Csv -Path "$path\Get-MailFlowRules.csv" -NoTypeInformation
