# Get all resource mailboxes
$ResourceMailboxes = Get-Mailbox -RecipientTypeDetails RoomMailbox, EquipmentMailbox

# Export mailbox settings to CSV
$ExportPath = "$([Environment]::GetFolderPath('Desktop'))\ResourceMailboxOptionsALL.csv"
$MailboxSettings = @()

foreach ($Mailbox in $ResourceMailboxes) {
    $MailboxSetting = Get-CalendarProcessing -Identity $Mailbox.Identity

    $MailboxSettingObj = [PSCustomObject]@{
        MailboxName                          = $Mailbox.Name
        PrimarySMTP                          = $Mailbox.PrimarySmtpAddress
        AutomateProcessing                   = $MailboxSetting.AutomateProcessing
        AllowConflicts                       = $MailboxSetting.AllowConflicts
        BookingWindowInDays                  = $MailboxSetting.BookingWindowInDays
        MaximumDurationInMinutes             = $MailboxSetting.MaximumDurationInMinutes
        AllowRecurringMeetings               = $MailboxSetting.AllowRecurringMeetings
        EnforceSchedulingHorizon             = $MailboxSetting.EnforceSchedulingHorizon
        ScheduleOnlyDuringWorkHours          = $MailboxSetting.ScheduleOnlyDuringWorkHours
        RemoveOldMeetingMessages             = $MailboxSetting.RemoveOldMeetingMessages
        AddOrganizerToSubject                = $MailboxSetting.AddOrganizerToSubject
        DeleteSubject                        = $MailboxSetting.DeleteSubject
        DeleteComments                       = $MailboxSetting.DeleteComments
        RemovePrivateProperty                = $MailboxSetting.RemovePrivateProperty
        DeleteNonCalendarItems               = $MailboxSetting.DeleteNonCalendarItems
        EnableResponseDetails                = $MailboxSetting.EnableResponseDetails
        AllowMultipleResources               = $MailboxSetting.AllowMultipleResources
        BookingType                          = $MailboxSetting.BookingType
        EnforceAdjacencyAsOverlap            = $MailboxSetting.EnforceAdjacencyAsOverlap
        EnforceCapacity                      = $MailboxSetting.EnforceCapacity
        ConflictPercentageAllowed            = $MailboxSetting.ConflictPercentageAllowed
        MaximumConflictInstances             = $MailboxSetting.MaximumConflictInstances
        ForwardRequestsToDelegates           = $MailboxSetting.ForwardRequestsToDelegates
        DeleteAttachments                    = $MailboxSetting.DeleteAttachments
        TentativePendingApproval             = $MailboxSetting.TentativePendingApproval
        OrganizerInfo                        = $MailboxSetting.OrganizerInfo
        ResourceDelegates                    = $MailboxSetting.ResourceDelegates
        RequestOutOfPolicy                   = $MailboxSetting.RequestOutOfPolicy
        AllRequestOutOfPolicy                = $MailboxSetting.AllRequestOutOfPolicy
        BookInPolicy                         = $MailboxSetting.BookInPolicy
        AllBookInPolicy                      = $MailboxSetting.AllBookInPolicy
        RequestInPolicy                      = $MailboxSetting.RequestInPolicy
        AllRequestInPolicy                   = $MailboxSetting.AllRequestInPolicy
        AddAdditionalResponse                = $MailboxSetting.AddAdditionalResponse
        AdditionalResponse                   = $MailboxSetting.AdditionalResponse
        AddNewRequestsTentatively            = $MailboxSetting.AddNewRequestsTentatively
        ProcessExternalMeetingMessages       = $MailboxSetting.ProcessExternalMeetingMessages
        RemoveForwardedMeetingNotifications  = $MailboxSetting.RemoveForwardedMeetingNotifications
        AutoRSVPConfiguration                = $MailboxSetting.AutoRSVPConfiguration
        RemoveCanceledMeetings               = $MailboxSetting.RemoveCanceledMeetings
        EnableAutoRelease                    = $MailboxSetting.EnableAutoRelease
        PostReservationMaxClaimTimeInMinutes = $MailboxSetting.PostReservationMaxClaimTimeInMinutes
        MailboxOwnerId                       = $MailboxSetting.MailboxOwnerId
        Identity                             = $MailboxSetting.Identity
        IsValid                              = $MailboxSetting.IsValid
        ObjectState                          = $MailboxSetting.ObjectState

    }

    $MailboxSettings += $MailboxSettingObj 
}

$MailboxSettings | Export-Csv -Path $ExportPath -NoTypeInformation
