$users = Import-Csv "C:\Users\ChrisIbold\Sterling Consulting\CBIZ - EBK Acquisition - General\Working Documents\Mappings\ebk-mw-mappings.csv"

Connect-MgGraph -Scopes "Directory.ReadWrite.All", "User.ReadWrite.All"
Write-Host "Processing" $($users | Measure-Object).Count

foreach ($user in $users) {
    try {
        Write-Host "Removing $($user.UserPrincipalName)'s login token..." -ForegroundColor Yellow
        Revoke-MgUserSignInSession -UserId $user.Id
    } catch {
        Write-Error "Error removing login token for $($user.UserPrincipalName)"
        Write-Error $_
    }
}