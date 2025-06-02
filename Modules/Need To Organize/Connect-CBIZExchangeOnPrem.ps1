$UserCredential = Import-Clixml "C:\Users\CIbold0\OneDrive - CBIZ, Inc\Desktop\Chris\OnPremCreds.xml"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://CLE1VEX2016.ad.cbiz.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential

Import-PSSession $Session -AllowClobber