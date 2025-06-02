$users = Import-Csv "C:\Users\ChrisIbold\Sterling Consulting\CBIZ - EBK Acquisition - General\Working Documents\Mappings\ebk-mw-mappings.csv"

Connect-MgGraph -Scopes DeviceManagementManagedDevices.Read.All, DeviceManagementManagedDevices.ReadWrite.All, DeviceManagementManagedDevices.PrivilegedOperations.All
Write-Host "Processing" $($users | Measure-Object).Count

foreach ($user in $users) {
    Write-Host "Retiring $($user."Source Email")"
    $devices = $null
    $devices = Get-MgDeviceManagementManagedDevice -All -Filter "(OperatingSystem eq 'Android' or OperatingSystem eq 'iOS') and userPrincipalName eq '$($user."Source Email")'"

    if ($devices) {
        Write-Host "Removing devices from Intune..." -ForegroundColor DarkYellow

        foreach ($device in $devices) {
            try {
                Write-Host $device.ManagedDeviceName -ForegroundColor Green
                Invoke-MgRetireDeviceManagementManagedDevice -ManagedDeviceId $device.Id -ErrorAction Stop
            } catch {
                Write-Error "Error retiring $($device.Id)"
                Write-Error $_
            }
        }
    }
}