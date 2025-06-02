<# -------------------------------------------
Initial draft script. In testing phase. In order for this to currently run you must be on-prem to check if the ADUser is disabled or not.
You must also have the SharePoint PnP PowerShell module installed for the MS list commands.
Looking into if it is possible to just use the Azure AD Graph API to pull enable/disable status from AAD.
--------------------------------------------#>

# Connect to SharePoint Online
Connect-PnPOnline "https://yourtenant.sharepoint.com/sites/yoursite"

# Define variables for the security group name and SharePoint list name
$securityGroupName = "Your Security Group Name"
$sharePointListName = "Your SharePoint List Name"

# Get the members of the security group
$members = Get-ADGroupMember -Identity $securityGroupName | Where-Object {$_.objectClass -eq 'user'}

# Get the SharePoint list
$list = Get-PnPList -Identity $sharePointListName

# Loop through the members and add or remove them from the SharePoint list as necessary
foreach ($member in $members) {
    # Check if the user is already in the SharePoint list
    if (Get-PnPListItem -List $list -Filter "Title eq '$($member.SamAccountName)'") {
        Write-Host "$($member.SamAccountName) is already in the SharePoint list"
    }
    else {
        # Add the user to the SharePoint list
        $itemProperties = @{
            Title = $member.SamAccountName
            # Add additional properties as needed
        }
        Add-PnPListItem -List $list -Values $itemProperties
        Write-Host "Added $($member.SamAccountName) to the SharePoint list"
    }
}

# Loop through the SharePoint list items and remove any that are not in the security group
$listItems = Get-PnPListItem -List $list
foreach ($item in $listItems) {
    # Check if the user is in the security group
    if ($members.SamAccountName -contains $item.FieldValues.Title) {
        Write-Host "$($item.FieldValues.Title) is still in the security group"
    }
    else {
        # Remove the user from the SharePoint list
        Remove-PnPListItem -List $list -Identity $item.Id
        Write-Host "Removed $($item.FieldValues.Title) from the SharePoint list"
    }
}
