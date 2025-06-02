
Connect-AzureAD

$csvPath = "$([Environment]::GetFolderPath('Desktop'))\UserList.csv"

$blockedAccounts = @()

$csvData = Import-Csv $csvPath


foreach ($row in $csvData) {
    $email = $row.SMTP

    $user = Get-AzureADUser -ObjectId $email -ErrorAction SilentlyContinue

    if ($user) {
        if ($user.AccountEnabled -eq $false) {
            $blockedAccount = [PSCustomObject]@{
                Email = $email
                Status = "Blocked"
            }
            $blockedAccounts += $blockedAccount
            Write-Host "Email: $email"
            Write-Host "Sign-in is blocked."
            Write-Host "------------------------"
        } else {
            Write-Host "Email: $email"
            Write-Host "Sign-in is not blocked."
            Write-Host "------------------------"
        }
    } else {
        Write-Host "Email: $email"
        Write-Host "Account not found."
        Write-Host "------------------------"
    }
}

$blockedAccounts | Export-Csv -Path $csvPath -NoTypeInformation  
