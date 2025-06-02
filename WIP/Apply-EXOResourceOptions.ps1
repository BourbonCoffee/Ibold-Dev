# Import the CSV file containing mailbox settings
$ImportPath = "$([Environment]::GetFolderPath('Desktop'))\ResourceMailboxOptionsALL.csv"
$MailboxSettings = Import-Csv -Path $ImportPath

# Loop through each mailbox setting and apply it to the corresponding resource mailbox
foreach ($Setting in $MailboxSettings) {
    $Mailbox = Get-Mailbox -Identity $Setting.sourceSMTP

    try {
        Write-Host "Applying settings to mailbox: $($Mailbox.DisplayName)"

        Set-CalendarProcessing -Identity $Mailbox `
            -AutomateProcessing $Setting.AutomateProcessing `
            -AllowConflicts ([bool]$Setting.AllowConflicts) `
            -BookingWindowInDays $Setting.BookingWindowInDays `
            -MaximumDurationInMinutes $Setting.MaximumDurationInMinutes `
            -AllowRecurringMeetings ([bool]$Setting.AllowRecurringMeetings) `
            -EnforceSchedulingHorizon ([bool]$Setting.EnforceSchedulingHorizon) `
            -ScheduleOnlyDuringWorkHours ([bool]$Setting.ScheduleOnlyDuringWorkHours) `
            -RemoveOldMeetingMessages ([bool]$Setting.RemoveOldMeetingMessages) `
            -AddOrganizerToSubject ([bool]$Setting.AddOrganizerToSubject) `
            -DeleteSubject ([bool]$Setting.DeleteSubject) `
            -DeleteComments ([bool]$Setting.DeleteComments) `
            -RemovePrivateProperty ([bool]$Setting.RemovePrivateProperty) `
            -DeleteNonCalendarItems ([bool]$Setting.DeleteNonCalendarItems) `
            -EnableResponseDetails ([bool]$Setting.EnableResponseDetails) `
            -EnforceCapacity ([bool]$Setting.EnforceCapacity) `
            -ConflictPercentageAllowed $Setting.ConflictPercentageAllowed `
            -MaximumConflictInstances $Setting.MaximumConflictInstances `
            -ForwardRequestsToDelegates ([bool]$Setting.ForwardRequestsToDelegates) `
            -DeleteAttachments ([bool]$Setting.DeleteAttachments) `
            -TentativePendingApproval ([bool]$Setting.TentativePendingApproval) `
            -OrganizerInfo ([bool]$Setting.OrganizerInfo) `
            -ResourceDelegates $Setting.ResourceDelegates `
            -AllRequestOutOfPolicy ([bool]$Setting.AllRequestOutOfPolicy) `
            -AllBookInPolicy ([bool]$Setting.AllBookInPolicy) `
            -AllRequestInPolicy ([bool]$Setting.AllRequestInPolicy) `
            -AddAdditionalResponse ([bool]$Setting.AddAdditionalResponse) `
            -AdditionalResponse $Setting.AdditionalResponse `
            -AddNewRequestsTentatively ([bool]$Setting.AddNewRequestsTentatively) `
            -ProcessExternalMeetingMessages ([bool]$Setting.ProcessExternalMeetingMessages) `
            -RemoveForwardedMeetingNotifications ([bool]$Setting.RemoveForwardedMeetingNotifications) `
            -RemoveCanceledMeetings ([bool]$Setting.RemoveCanceledMeetings) `
            -EnableAutoRelease ([bool]$Setting.EnableAutoRelease) `
            -PostReservationMaxClaimTimeInMinutes $Setting.PostReservationMaxClaimTimeInMinutes

        Write-Host "Settings applied successfully to mailbox: $($Mailbox.DisplayName)"
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Error applying settings to mailbox: $($Mailbox.DisplayName) - $errorMessage"
    }
}

Write-Host "Mailbox settings applied to resource mailboxes."