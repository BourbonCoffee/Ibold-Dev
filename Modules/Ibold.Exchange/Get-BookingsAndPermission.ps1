$BookingsMailboxesWithPermissions = New-Object 'System.Collections.Generic.List[System.Object]'
# Get all Booking Mailboxes
$allBookingsMailboxes = Get-EXOMailbox -RecipientTypeDetails SchedulingMailbox -ResultSize:Unlimited

# Loop through the list of Mailboxes
$BookingsMailboxesWithPermissions = foreach ($bookingsMailbox in $allBookingsMailboxes) {
    # Get Permissions for this Mailbox
    $allPermissionsForThisMailbox = Get-EXOMailboxPermission -UserPrincipalName $bookingsMailbox.UserPrincipalName -ResultSize:Unlimited | Where-Object { ($_.User -like '*@*') -and ($_.AccessRights -eq "FullAccess") }
    foreach ($permission in $allPermissionsForThisMailbox) {
        # Output PSCustomObject with infos to the foreach loop, so it gets saved into $BookingsMailboxesWithPermissions
        [PSCustomObject]@{
            'Bookings Mailbox DisplayName'    = $bookingsMailbox.DisplayName
            'Bookings Mailbox E-Mail-Address' = $bookingsMailbox.PrimarySmtpAddress
            'User'                            = $permission.User
            'AccessRights'                    = "Administrator"
        }
    }
}
$BookingsMailboxesWithPermissions | Export-Csv C:\temp\bookings-permissions.csv -Encoding utf8 -Delimiter ";" -NoTypeInformation
