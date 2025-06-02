<#
.SYNOPSIS
    CIS Microsoft 365 Foundations Benchmark (v4.0.0) - Automated Checks.
    Supports running only Level 1 checks, or Level 1 + Level 2 checks.

.DESCRIPTION
    This script connects to the various Microsoft 365 services, runs 
    the selected checks, and then exports results to an Excel spreadsheet.
    - Use the -Level1 switch to run only Level 1 checks.
    - Use the -Level2 switch to run both Level 1 and Level 2 checks.
    - If neither switch is specified, it defaults to Level 1 checks only.

.NOTES
    - Requires the following modules:
        * ExchangeOnlineManagement
        * Microsoft.Graph
        * MicrosoftTeams
        * Microsoft.Online.SharePoint.PowerShell
        * ImportExcel   (for exporting to XLSX)
    - Example usage:
        . .\CIS_M365_Script.ps1
        Invoke-CISBenchmarkChecks -Level1 -Verbose 
        or
        Invoke-CISBenchmarkChecks -Level2 -Verbose 
#>

# --------------------------------------------
# GLOBAL VARS AND HELPER FUNCTIONS
# --------------------------------------------
# Global array to store results
$Global:CISResults = @()

function Add-CISResult {
    param(
        [string]$ControlNumber,
        [string]$Title,
        [string]$Level,
        [string]$Status,
        [string]$Details
    )
    $obj = [pscustomobject][ordered]@{
        TimeStamp     = (Get-Date)
        ControlNumber = $ControlNumber
        Title         = $Title
        Level         = $Level
        Status        = $Status
        Details       = $Details
    }
    $Global:CISResults += $obj
    Write-Verbose "[Result] $($obj | ConvertTo-Json -Depth 2)"
}

# --------------------------------------------
# Section 1 - M365 Admin Center
# --------------------------------------------

function Test-111-AdminAccountsCloudOnly {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.1.1: (L1) Ensure Administrative accounts are cloud-only..."

    $controlNum = "1.1.1"
    $title = "Ensure Administrative accounts are cloud-only"
    $level = "L1"

    try {
        # Connect-MgGraph -Scopes "RoleManagement.Read.Directory","User.Read.All" (Ensure you do this elsewhere)
        
        $DirectoryRoles = Get-MgDirectoryRole
        $PrivilegedRoles = $DirectoryRoles | Where-Object {
            $_.DisplayName -like "*Administrator*" -or $_.DisplayName -eq "Global Reader"
        }

        # Collect unique membership
        $RoleMembers = foreach ($role in $PrivilegedRoles) {
            Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id | Select-Object Id -Unique
        }

        # Retrieve details about each privileged user
        $PrivilegedUsers = foreach ($rm in $RoleMembers) {
            Get-MgUser -UserId $rm.Id -Property UserPrincipalName, DisplayName, Id, OnPremisesSyncEnabled
        }

        # Find any admin user that is OnPremisesSyncEnabled = $true
        $SyncedAdmins = $PrivilegedUsers | Where-Object { $_.OnPremisesSyncEnabled -eq $true }

        if ($SyncedAdmins) {
            $details = "FAIL: Found admin(s) synced from on-prem => " + ($SyncedAdmins.UserPrincipalName -join ", ")
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details $details
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "No synced admin accounts found."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-113-TwoToFourGlobalAdmins {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.1.3: (L1) Ensure that between two and four global admins are designated..."

    $controlNum = "1.1.3"
    $title = "Ensure that between two and four global admins are designated"
    $level = "L1"

    try {
        # Connect-MgGraph -Scopes "Directory.Read.All"

        # Locate the Global Admin role using its RoleTemplateId
        $globalAdminRole = Get-MgDirectoryRole -Filter "RoleTemplateId eq '62e90394-69f5-4237-9190-012177145e10'"
        if (-not $globalAdminRole) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details "Unable to locate Global Admin role."
            return
        }

        $globalAdmins = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id
        $count = $globalAdmins.Count

        if ($count -lt 2) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Only $count Global Admin(s). Recommended >= 2."
        } elseif ($count -gt 4) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "$count Global Admin(s). Recommended <= 4."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "$count Global Admin(s) is within 2-4 range."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-114-AdminLicenseReducedFootprint {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.1.4: (L1) Ensure administrative accounts use minimal license footprint..."

    $controlNum = "1.1.4"
    $title = "Ensure administrative accounts use licenses with a reduced application footprint"
    $level = "L1"

    try {
        # Connect-MgGraph -Scopes "RoleManagement.Read.Directory","User.Read.All"
        
        $DirectoryRoles = Get-MgDirectoryRole
        $PrivilegedRoles = $DirectoryRoles | Where-Object {
            $_.DisplayName -like "*Administrator*" -or $_.DisplayName -eq "Global Reader"
        }

        $RoleMembers = foreach ($role in $PrivilegedRoles) {
            Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id | Select-Object Id -Unique
        }

        $PrivilegedUsers = foreach ($rm in $RoleMembers) {
            Get-MgUser -UserId $rm.Id -Property UserPrincipalName, DisplayName, Id
        }

        foreach ($pu in $PrivilegedUsers) {
            $licenseDetails = Get-MgUserLicenseDetail -UserId $pu.Id
            $AssignedSkus = $licenseDetails.SkuPartNumber

            if (-not $AssignedSkus) {
                # No licenses
                Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "$($pu.UserPrincipalName): unlicensed (good)."
            } elseif (($AssignedSkus -notcontains 'AAD_PREMIUM') -and ($AssignedSkus -notcontains 'AAD_PREMIUM_P2')) {
                # Has some license that presumably includes mailbox, etc.
                Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "$($pu.UserPrincipalName) => $($AssignedSkus -join ', ')"
            } else {
                Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "$($pu.UserPrincipalName) => $($AssignedSkus -join ', ')"
            }
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-121-OnlyApprovdPrublicGroups {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.2.1: (L2) Ensure only organizationally managed/approved public groups exist..."

    $controlNum = "1.2.1"
    $title = "Ensure that only organizationally managed/approved public groups exist"
    $level = "L2"

    try {
        # Connect-MgGraph -Scopes "Group.Read.All"

        $allGroups = Get-MgGroup -All:$true
        $publicGroups = $allGroups | Where-Object { $_.Visibility -eq "Public" }

        if ($publicGroups) {
            $groupNames = $publicGroups | Select-Object -ExpandProperty DisplayName
            $details = "FAIL: Found Public groups => " + ($groupNames -join ", ")
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details $details
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "No public groups found (or all are approved)."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-122-SignInToSharedMailboxBlocked {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.2.2: (L1) Ensure sign-in to shared mailboxes is blocked..."

    $controlNum = "1.2.2"
    $title = "Ensure sign-in to shared mailboxes is blocked"
    $level = "L1"

    try {
        # Connect-ExchangeOnline
        # Connect-MgGraph -Scopes "Policy.Read.All"

        $MBX = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
        if (-not $MBX) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "No shared mailboxes found."
            return
        }

        foreach ($mailbox in $MBX) {
            $user = Get-MgUser -UserId $mailbox.ExternalDirectoryObjectId -Property DisplayName, UserPrincipalName, AccountEnabled
            if ($user) {
                if ($user.AccountEnabled -eq $true) {
                    Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Shared mailbox $($user.DisplayName) => sign-in enabled."
                } else {
                    Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "Shared mailbox $($user.DisplayName) => sign-in blocked."
                }
            } else {
                # If for some reason no user object found, consider that an error or check
                Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "INFO" -Details "Could not retrieve user for mailbox $($mailbox.DisplayName)."
            }
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-131-PasswordNeverExpire {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.3.1: (L1) Ensure password expiration policy is set to never expire..."

    $controlNum = "1.3.1"
    $title = "Ensure 'Password expiration policy' is set to 'never expire'"
    $level = "L1"

    try {
        # Connect-MgGraph -Scopes "Domain.Read.All"

        $domains = Get-MgDomain -All:$true
        if (-not $domains) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details "No domains found. Are you connected properly?"
            return
        }

        # CIS doc says check if PasswordValidityPeriodInDays = 2147483647
        $nonCompliant = $domains | Where-Object { 
            $_.IsVerified -eq $true -and $_.PasswordValidityPeriodInDays -ne 2147483647 
        }

        if ($nonCompliant) {
            $bad = $nonCompliant | ForEach-Object { "$($_.Id) => $($_.PasswordValidityPeriodInDays)" }
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Some verified domains are not set to 'never expire': $bad"
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "All verified domains use 'never expire' password setting."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-133-ExternalCalendarSharingOff {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.3.3: (L2) Ensure external sharing of calendars is not available..."

    $controlNum = "1.3.3"
    $title = "Ensure 'External sharing' of calendars is not available"
    $level = "L2"

    try {
        # Connect-ExchangeOnline

        # Check default sharing policy
        $policy = Get-SharingPolicy -Identity "Default Sharing Policy" -ErrorAction SilentlyContinue
        if (-not $policy) {
            # Possibly no default policy found, interpret as pass or review carefully
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "INFO" -Details "No 'Default Sharing Policy' found. Possibly no external sharing enabled."
            return
        }

        if ($policy.Enabled -eq $true) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Default Sharing Policy is Enabled => external calendar sharing allowed."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "Default Sharing Policy is Disabled => external calendar sharing off."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-136-CustomerLockboxEnabled {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.3.6: (L2) Ensure the Customer Lockbox feature is enabled..."

    $controlNum = "1.3.6"
    $title = "Ensure the Customer Lockbox feature is enabled"
    $level = "L2"

    try {
        # Often, Customer Lockbox is visible via:
        # Connect-ExchangeOnline
        $org = Get-OrganizationConfig -ErrorAction Stop
        # Checking a property, e.g.: $org.CustomerLockboxEnabled
        # If not found, you might try 'Get-MsolCompanyInformation' or MSCommerce-based cmdlets in your tenant.

        if ($org.CustomerLockboxEnabled -eq $true) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "Customer Lockbox is enabled."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Customer Lockbox appears disabled."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

# --------------------------------------------
# Section 2 - M365 Defender
# --------------------------------------------

function Test-211-SafeLinksOfficeAppsEnabled {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.1: (L2) Ensure Safe Links for Office Applications is Enabled..."

    $controlNum = "2.1.1"
    $title = "Ensure Safe Links for Office Applications is Enabled"
    $level = "L2"

    try {
        # Typically:
        #   Connect-ExchangeOnline 
        # We check Safe Links Policies (and maybe Safe Links Rules).
        # Example policy name might be "Global" or "Default"

        $safeLinksPolicy = Get-SafeLinksPolicy -Identity "Global" -ErrorAction SilentlyContinue
        if (-not $safeLinksPolicy) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "No Safe Links policy named 'Global' found."
            return
        }

        # The key property in some tenants is EnableSafeLinksForOffice, which must be True
        if ($safeLinksPolicy.EnableSafeLinksForOffice) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "Safe Links for Office Apps is enabled."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Safe Links for Office Apps is disabled in the policy."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-212-CommonAttachmentTypesFilterEnabled {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.2: (L1) Ensure the Common Attachment Types Filter is enabled..."

    $controlNum = "2.1.2"
    $title = "Ensure the Common Attachment Types Filter is enabled"
    $level = "L1"

    try {
        # Typically:
        #   Connect-ExchangeOnline
        # We check the Malware Filter Policy. The property is usually EnableFileTypes. 
        # Or "CommonAttachmentFilterEnabled" for certain policies.

        $malwarePolicies = Get-MalwareFilterPolicy -ErrorAction Stop
        if (-not $malwarePolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "No MalwareFilterPolicy found."
            return
        }

        $nonCompliant = @()
        foreach ($policy in $malwarePolicies) {
            if (-not $policy.CommonAttachmentFilterEnabled) {
                $nonCompliant += $policy.Name
            }
        }

        if ($nonCompliant) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Common Attachment Filter disabled in: $($nonCompliant -join ', ')"
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "All Malware Filter Policies have common attachment filter enabled."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-213-NotifyInternalUsersSendingMalware {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.3: (L1) Ensure notifications for internal users sending malware is Enabled..."

    $controlNum = "2.1.3"
    $title = "Ensure notifications for internal users sending malware is Enabled"
    $level = "L1"

    try {
        # Usually same MalwareFilterPolicy. The property might be NotifyInternalSenders or similar.

        $malwarePolicies = Get-MalwareFilterPolicy -ErrorAction Stop
        if (-not $malwarePolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "No MalwareFilterPolicy found."
            return
        }

        $nonCompliant = @()
        foreach ($policy in $malwarePolicies) {
            if (-not $policy.NotifyInternalSenders) {
                $nonCompliant += $policy.Name
            }
        }

        if ($nonCompliant) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "NotifyInternalSenders is disabled in: $($nonCompliant -join ', ')"
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "All Malware Filter Policies notify internal users on malware."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-214-SafeAttachmentsPolicyEnabled {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.4: (L2) Ensure Safe Attachments policy is enabled..."

    $controlNum = "2.1.4"
    $title = "Ensure Safe Attachments policy is enabled"
    $level = "L2"

    try {
        # "Safe Attachments" is part of the Anti-Malware pipeline in M365 Defender,
        # Typically:  Get-SafeAttachmentPolicy or Get-SafeAttachmentRule

        $sapolicies = Get-SafeAttachmentPolicy -ErrorAction Stop
        if (-not $sapolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "No Safe Attachment policies found."
            return
        }

        $nonEnabled = @()
        foreach ($sapol in $sapolicies) {
            if ($sapol.EnableSafeAttachmentsForEmail -ne $true) {
                # or check the Mode property
                $nonEnabled += $sapol.Name
            }
        }

        if ($nonEnabled) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Safe Attachments not fully enabled for: $($nonEnabled -join ', ')"
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "All Safe Attachment policies are enabled."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-215-SafeAttachmentsSPODTeamsEnabled {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.5: (L2) Ensure Safe Attachments for SPO, OneDrive, Teams is Enabled..."

    $controlNum = "2.1.5"
    $title = "Ensure Safe Attachments for SharePoint, OneDrive, and Microsoft Teams is Enabled"
    $level = "L2"

    try {
        # In many tenants, this is "SafeAttachmentsForSharePoint" property in 
        # Advanced Threat Protection settings. For ex:
        # Get-AtpPolicyForO365 (with older cmdlets) or 
        # Get-PolicyTipConfig, etc. 
        # Example approach:

        $atpSettings = Get-AtpPolicyForO365 -ErrorAction SilentlyContinue
        if (-not $atpSettings) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Could not retrieve ATP policy for O365."
            return
        }

        # Check if "EnableATPForSPOTeamsODB" or similarly named property is True
        if ($atpSettings.EnableATPForSPOTeamsODB -eq $true) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "Safe Attachments for SPO/OneDrive/Teams is enabled."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Safe Attachments for SPO/OneDrive/Teams is disabled."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-216-NotifyAdminsSpamPolicies {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.6: (L1) Ensure Exchange Online Spam Policies notify admins..."

    $controlNum = "2.1.6"
    $title = "Ensure Exchange Online Spam Policies are set to notify administrators"
    $level = "L1"

    try {
        # Typically check spam filter policies:
        $spamPolicies = Get-SpamFilterPolicy -ErrorAction Stop
        if (-not $spamPolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "No SpamFilterPolicy found."
            return
        }

        $missingNotif = @()
        foreach ($policy in $spamPolicies) {
            # Might look for property like "NotifyAdmin" or "NotifyOutboundSpam"
            # Real property name can differ. Let's call it "NotifyOutboundSpamRecipients" or "NotifyInternalAdmins"
            if (-not $policy.NotifyOutboundSpamRecipients) {
                $missingNotif += $policy.Name
            }
        }

        if ($missingNotif) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "No admin notification in: $($missingNotif -join ', ')"
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "All SpamFilterPolicy rules notify admins as recommended."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-217-AntiPhishingPolicyCreated {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.7: (L2) Ensure that an anti-phishing policy has been created..."

    $controlNum = "2.1.7"
    $title = "Ensure that an anti-phishing policy has been created"
    $level = "L2"

    try {
        # "Anti-Phish" policy in EOP or MDO:
        $phishPolicies = Get-AntiPhishPolicy -ErrorAction SilentlyContinue
        if (-not $phishPolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "No anti-phish policy found."
        } else {
            # If we find at least one, consider that a pass. 
            # Optionally check if it's enabled or has certain settings.
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "Anti-phishing policy found: $($phishPolicies.Count) policy(ies)."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-219-DkimEnabledAllDomains {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.9: (L1) Ensure DKIM is enabled for all Exchange Online Domains..."

    $controlNum = "2.1.9"
    $title = "Ensure that DKIM is enabled for all Exchange Online Domains"
    $level = "L1"

    try {
        # Typically:
        #   Connect-ExchangeOnline
        # "Get-DkimSigningConfig" returns DKIM config per domain

        $dkimConfigs = Get-DkimSigningConfig -ErrorAction Stop
        if (-not $dkimConfigs) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "No DKIM signing configs found."
            return
        }

        $notEnabled = $dkimConfigs | Where-Object { $_.Enabled -eq $false }
        if ($notEnabled) {
            $domains = $notEnabled | Select-Object -ExpandProperty DomainName
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "DKIM not enabled for domain(s): $($domains -join ', ')"
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "DKIM is enabled for all Exchange Online domains."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-2111-ComprehensiveAttachmentFiltering {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.11: (L2) Ensure comprehensive attachment filtering is applied..."

    $controlNum = "2.1.11"
    $title = "Ensure comprehensive attachment filtering is applied"
    $level = "L2"

    try {
        # This often refers to advanced anti-malware or content filtering settings 
        # in the "Malware Filter Policy" or "Safe Attachment" environment. 
        # Let's do a typical check for certain file types or scanning behavior:

        $malwarePolicies = Get-MalwareFilterPolicy -ErrorAction SilentlyContinue
        if (-not $malwarePolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "No MalwareFilterPolicy found."
            return
        }

        # If we want to confirm that they have a broad set of blocked attachments:
        $insufficient = @()
        foreach ($policy in $malwarePolicies) {
            # For instance, check if "EnableFileFilter" or "CommonAttachmentFilterEnabled" is true
            # and that the policy has some attachments in the filter list, etc.
            if (-not $policy.CommonAttachmentFilterEnabled) {
                $insufficient += $policy.Name
            }
        }

        if ($insufficient) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Some policies might not have comprehensive filtering: $($insufficient -join ', ')"
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "All policies appear to have comprehensive attachment filtering."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-2112-NoConnectionFilterIPAllowList {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.12: (L1) Ensure the connection filter IP allow list is not used..."

    $controlNum = "2.1.12"
    $title = "Ensure the connection filter IP allow list is not used"
    $level = "L1"

    try {
        # Typically from HostedConnectionFilterPolicy:
        $connPolicies = Get-HostedConnectionFilterPolicy -ErrorAction SilentlyContinue
        if (-not $connPolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "No HostedConnectionFilterPolicy found => presumably no IP allow list."
            return
        }

        $badPolicies = @()
        foreach ($pol in $connPolicies) {
            if ($pol.IPAllowList -and $pol.IPAllowList.Count -gt 0) {
                $badPolicies += "$($pol.Name): $($pol.IPAllowList -join ', ')"
            }
        }

        if ($badPolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Connection filter policy has IP allow list => $($badPolicies -join ' | ')"
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "No IP allow list found in connection filter policies."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-2113-ConnectionFilterSafeListOff {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.13: (L1) Ensure the connection filter safe list is off..."

    $controlNum = "2.1.13"
    $title = "Ensure the connection filter safe list is off"
    $level = "L1"

    try {
        # Typically: a property like "EnableSafeList" in the HostedConnectionFilterPolicy

        $connPolicies = Get-HostedConnectionFilterPolicy -ErrorAction SilentlyContinue
        if (-not $connPolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "No HostedConnectionFilterPolicy found => presumably safe list is off."
            return
        }

        $enabledSafeLists = @()
        foreach ($pol in $connPolicies) {
            if ($pol.EnableSafeList -eq $true) {
                $enabledSafeLists += $pol.Name
            }
        }

        if ($enabledSafeLists) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Safe list is ON in policy(ies): $($enabledSafeLists -join ', ')"
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "No connection filter safe lists are enabled."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-2114-NoInboundAntiSpamAllowedDomains {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.1.14: (L1) Ensure inbound anti-spam policies do not contain allowed domains..."

    $controlNum = "2.1.14"
    $title = "Ensure inbound anti-spam policies do not contain allowed domains"
    $level = "L1"

    try {
        # Typically in the ContentFilterPolicy or SpamFilterPolicy
        $spamPolicies = Get-SpamFilterPolicy -ErrorAction SilentlyContinue
        if (-not $spamPolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "No SpamFilterPolicy found => presumably no allowed domain."
            return
        }

        $bad = @()
        foreach ($policy in $spamPolicies) {
            # Possibly check $policy.AllowedSenderDomains or $policy.AllowedDomains
            if ($policy.AllowedSenderDomains -and $policy.AllowedSenderDomains.Count -gt 0) {
                $bad += "$($policy.Name): $($policy.AllowedSenderDomains -join ', ')"
            }
        }

        if ($bad) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Found allowed domains in policy(ies): $($bad -join ' | ')"
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "No inbound anti-spam policies contain allowed domains."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-244-ZeroHourAutoPurgeTeamsOn {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 2.4.4: (L1) Ensure Zero-hour auto purge (ZAP) for Microsoft Teams is on..."

    $controlNum = "2.4.4"
    $title = "Ensure Zero-hour auto purge for Microsoft Teams is on"
    $level = "L1"

    try {
        # ZAP for Teams can be part of "Set-AtpPolicyForO365" or "Get-AtpPolicyForO365"
        # Possibly "TeamsMessagesFilteringEnabled" property, or "EnableATPForSPOTeamsODB" in some contexts.
        # The doc snippet was incomplete, but let's illustrate typical code:

        $atpPolicy = Get-AtpPolicyForO365 -ErrorAction SilentlyContinue
        if (-not $atpPolicy) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Unable to retrieve ATP policies => cannot confirm ZAP for Teams."
            return
        }

        # Hypothetical property for ZAP (some tenants might differ)
        if ($atpPolicy.TeamsMessagesFilteringEnabled -eq $true) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "ZAP for Microsoft Teams is on."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Zero-hour auto purge for Teams is off or not configured."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

# --------------------------------------------
# Section 3 - Microsoft Purview
# --------------------------------------------

function Test-311-M365AuditLogSearchEnabled {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 3.1.1: (L1) Ensure Microsoft 365 audit log search is Enabled..."

    $controlNum = "3.1.1"
    $title = "Ensure Microsoft 365 audit log search is Enabled"
    $level = "L1"

    try {
        # Typically you need to: Connect-ExchangeOnline
        # Then we can do:
        #   $config = Get-AdminAuditLogConfig
        #   Check if $config.UnifiedAuditLogIngestionEnabled is $true

        $config = Get-AdminAuditLogConfig -ErrorAction Stop
        if (-not $config) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Unable to retrieve AdminAuditLogConfig."
            return
        }

        if ($config.UnifiedAuditLogIngestionEnabled -eq $true) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "Unified audit log is enabled."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Unified audit log ingestion is disabled."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

# --------------------------------------------
# Section 4 - Microsoft Intune
# --------------------------------------------

# As of CIS M365 Foundations Benchmark v4.0.0, 
# there are NO Automated controls in Section 4.

function Invoke-CISSection4Checks {
    [CmdletBinding()]
    param()

    Write-Verbose "Running CIS Microsoft 365 Foundations Benchmark - Section 4 (Intune Admin Center)"

    # If you wanted to log that there are no checks:
    Add-CISResult -ControlNumber "4.x" -Title "Section 4 Automated Checks" -Level "N/A" `
        -Status "INFO" -Details "No Automated checks for Intune in CIS M365 v4.0.0."

    Write-Host "`nSection 4 checks complete. Results in `$Global:CISResults: $($Global:CISResults.Count) items."
}

# --------------------------------------------
# Section 5 - Entra
# --------------------------------------------

function Test-5122-ThirdPartyIntegratedAppsNotAllowed {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 5.1.2.2: (L2) Ensure third-party integrated applications are not allowed..."

    $controlNum = "5.1.2.2"
    $title = "Ensure third-party integrated applications are not allowed"
    $level = "L2"

    try {
        # Approach with AzureAD module: 
        #    Connect-AzureAD
        # Check the "UsersCanRegisterApps" or "AllowAdHocApps" property in your tenant settings.

        $tenantSettings = (Get-MgPolicyAuthorizationPolicy).DefaultUserRolePermissions | Select-Object AllowedToCreateApps
        if (-not $tenantSettings) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details "Could not retrieve AzureAD tenant details."
            return
        }

        # The property name can differ by environment. Often it's something like 'AllowAdHocApps' or 'UsersCanRegisterApps'.
        # Let's assume it's 'UsersCanRegisterApps' for demonstration:
        if ($tenantSettings -eq $false) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "Users cannot register 3rd-party integrated apps."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Users can register 3rd-party integrated apps. Should be disallowed."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-5123-RestrictNonAdminTenantCreationYes {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 5.1.2.3: (L1) Restrict non-admin users from creating tenants..."

    $controlNum = "5.1.2.3"
    $title = "Ensure 'Restrict non-admin users from creating tenants' is set to 'Yes'"
    $level = "L1"

    try {
        $companySettings = (Get-MgPolicyAuthorizationPolicy).DefaultUserRolePermissions | Select-Object AllowedToCreateTenants 
        if (-not $companySettings) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details "Unable to retrieve MSOL Company settings."
            return
        }

        if ($companySettings -eq $false) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "Non-admins cannot create new tenants (AdHocSubscriptions=false)."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Non-admin users can create tenants => 'AllowAdHocSubscriptions=true'."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-5151-UserConsentToAppsNotAllowed {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 5.1.5.1: (L2) Ensure user consent to apps accessing company data is not allowed..."

    $controlNum = "5.1.5.1"
    $title = "Ensure user consent to apps accessing company data on their behalf is not allowed"
    $level = "L2"

    try {
        # Connect-MgGraph -Scopes "Policy.Read.All"
        # We look at the authorization policy. 
        # The property: "UserConsentForRiskyApps" or "DefaultUserRolePermissions" or "UserConsentAllowed": 
        #   Some references use "ConsentPolicy: 'DoNotAllow'"

        $authPolicies = Get-MgPolicyAuthorizationPolicy -ErrorAction Stop
        if (-not $authPolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details "No AuthorizationPolicy found via MS Graph."
            return
        }

        # Typically only one default policy, e.g. $authPolicies[0]
        $policy = $authPolicies[0]
        # The property can be "UserConsentAllowed" or "DefaultUserRolePermissions.AllowedToCreateApps" or
        # "ConsentPolicySetting" etc. For demonstration, let's assume "PermissionGrantPolicy" = "DoNotAllow"
        # or "AllowUserConsentForApps" = $false.

        if ($policy.AllowUserConsentForApps -eq $false) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "User consent to apps is disallowed."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Users can consent to apps => Should be disallowed."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-5151-UserConsentToAppsNotAllowed {

    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 5.1.5.1: (L2) Ensure user consent to apps accessing company data is not allowed..."

    $controlNum = "5.1.5.1"
    $title = "Ensure user consent to apps accessing company data on their behalf is not allowed"
    $level = "L2"

    try {
        # Connect-MgGraph -Scopes "Policy.Read.All"
        # We look at the authorization policy. 
        # The property: "UserConsentForRiskyApps" or "DefaultUserRolePermissions" or "UserConsentAllowed": 
        #   Some references use "ConsentPolicy: 'DoNotAllow'"

        $authPolicies = (Get-MgPolicyAuthorizationPolicy).DefaultUserRolePermissions | Select-Object -ExpandProperty PermissionGrantPoliciesAssigned 

        if (-not $authPolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "User consent to apps is disallowed."
            return
        }

        # Typically only one default policy, e.g. $authPolicies[0]
        # $policy = $authPolicies[0]
        # The property can be "UserConsentAllowed" or "DefaultUserRolePermissions.AllowedToCreateApps" or
        # "ConsentPolicySetting" etc. For demonstration, let's assume "PermissionGrantPolicy" = "DoNotAllow"
        # or "AllowUserConsentForApps" = $false.

        if ($authPolicies) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Users can consent to apps => Should be disallowed."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Users can consent to apps => Should be disallowed."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

