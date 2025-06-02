$path = "$([Environment]::GetFolderPath('Desktop'))"
Connect-ExchangeOnline

$inboundMailConnectors = Get-InboundConnector
$inboundMailConnectors | Export-Csv -Path "$path\Get-InboundConnectors.csv" -NoTypeInformation

$outboundMailConnectors = Get-OutboundConnector
$outboundMailConnectors | Export-Csv -Path "$path\Get-OutboundConnectors.csv" -NoTypeInformation
