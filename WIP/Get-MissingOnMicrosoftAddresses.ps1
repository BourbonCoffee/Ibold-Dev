$mbx = Get-Mailbox -ResultSize Unlimited -Filter { EmailAddresses -notlike "*Compudata0365.onmicrosoft.com" }
 
$mbx | Export-Csv -Path "C:\Users\ChrisIbold\OneDrive - Sterling Consulting\Desktop\MailboxesWithoutOnMicrosoft.csv" -NoTypeInformation
