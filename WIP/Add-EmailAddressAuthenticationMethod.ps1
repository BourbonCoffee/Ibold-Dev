# This script was created to quickly add all EBK users' CBIZ accounts as an authentication method to enable SSPR.
# Script 

Connect-Graph -Scopes "UserAuthenticationMethod.Read.All", "UserAuthenticationMethod.ReadWrite.All"

$Users = Import-Csv "C:\Users\ChrisIbold\Sterling Consulting\CBIZ - EBK Acquisition - General\Working Documents\Mappings\ebk-mw-mappings.csv"
$i = 0
ForEach ($user in $users) {
    $i++
    Write-Host "Working on User $i of $($users.Count) - $($user."Source Email")" -ForegroundColor Yellow
    Write-Host "Setting Email $($user."Destination Email")" -ForegroundColor DarkYellow

    try {
        $cbizEmail = (Get-MgUserAuthenticationEmailMethod -UserId $user."Source Email" | Where-Object { $_.EmailAddress -like "*" }).EmailAddress
        If ($cbizEmail) {
            If ($cbizEmail -eq $user."Destination Email") {
                Write-Host "Email already matches - No action taken`n" -ForegroundColor Red
            }
            If ($cbizEmail -ne $user."Destination Email") {
                Write-Host "Different email is already populated with: $cbizEmail `nPlease remove and add their CBIZ one!`n" -ForegroundColor Red
            }
        }
        If (!$cbizEmail) {
            $gobbleThemOutputs = New-MgUserAuthenticationEmailMethod -UserId $user."Source Email" -EmailAddress $user."Destination Email"
            Write-Host "No current email - populated with $($user."Destination Email")`n" -ForegroundColor Green
        }
    } catch {
        Write-Host "An error occurred: `n`n$_" -ForegroundColor Red
    }
    #run once
    #break
}