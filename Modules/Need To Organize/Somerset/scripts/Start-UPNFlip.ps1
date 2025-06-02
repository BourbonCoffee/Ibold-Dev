Get-MsolUser -All | Where-Object { $_.UserPrincipalName -like "*@somersetcpas.com" } | ForEach-Object {
    $upn = $_.UserPrincipalName
    $newUpn = $upn.Split('@')[0] + "@somr.onmicrosoft.com"
    try {
        Set-MsolUserPrincipalName -UserPrincipalName $upn -NewUserPrincipalName $newUpn
        Write-Host "Successfully changed UPN from $upn to $newUpn"
    } catch {
        Write-Host "Failed to change UPN from $upn to $newUpn. Error: $_"
    }
}