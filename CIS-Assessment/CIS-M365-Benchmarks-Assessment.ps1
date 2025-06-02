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
# 1. GLOBAL VARS AND HELPER FUNCTIONS
# --------------------------------------------
$Global:CISResults = @()

function Add-CISResult {
    param(
        [string]$ControlNumber,
        [string]$Title,
        [string]$Level, # "L1" or "L2"
        [string]$Status, # e.g. "PASS", "FAIL", "INFO", "ERROR"
        [string]$Details
    )
    # Create object and add to global array
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
# 2. SERVICE CONNECTION FUNCTIONS
# --------------------------------------------
function Connect-ExchangeService {
    [CmdletBinding()]
    param()
    Write-Verbose "Attempting to connect to Exchange Online..."
    try {
        Connect-ExchangeOnline -ErrorAction Stop
        Write-Verbose "Connected to Exchange Online."
    } catch {
        Write-Warning "Failed to connect to Exchange Online: $($_.Exception.Message)"
    }
}

function Connect-MgService {
    [CmdletBinding()]
    param()
    Write-Verbose "Attempting to connect to Microsoft Graph..."
    try {
        # Request read scopes needed for typical CIS checks:
        Connect-MgGraph -Scopes "Directory.Read.All", "RoleManagement.Read.Directory", "User.Read.All", "Group.Read.All" -ErrorAction Stop
        Write-Verbose "Connected to Microsoft Graph."
    } catch {
        Write-Warning "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    }
}

function Connect-TeamsService {
    [CmdletBinding()]
    param()
    Write-Verbose "Attempting to connect to Microsoft Teams..."
    try {
        Connect-MicrosoftTeams -ErrorAction Stop
        Write-Verbose "Connected to Microsoft Teams."
    } catch {
        Write-Warning "Failed to connect to Microsoft Teams: $($_.Exception.Message)"
    }
}

function Connect-SPOSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AdminUrl
    )
    Write-Verbose "Attempting to connect to SharePoint Online using $AdminUrl..."
    try {
        Connect-SPOService -Url $AdminUrl -ErrorAction Stop
        Write-Verbose "Connected to SharePoint Online."
    } catch {
        Write-Warning "Failed to connect to SharePoint Online: $($_.Exception.Message)"
    }
}

function Connect-CISModules {
    [CmdletBinding()]
    param(
        [string]$SharePointAdminUrl = "https://<YourTenant>-admin.sharepoint.com"
    )

    Write-Verbose "Starting all M365 service connections..."
    Connect-ExchangeService
    Connect-MgService
    Connect-TeamsService
    Connect-SPOSite -AdminUrl $SharePointAdminUrl
    Write-Verbose "All service connections attempted."
}

# --------------------------------------------
# 3. SAMPLE CHECK FUNCTIONS (L1 & L2)
# --------------------------------------------

#
# -- Level 1 EXAMPLES --
#

function Test-111-AdminAccountsCloudOnly {
    # L1
    <#
.LINK
    CIS 1.1.1 (L1, Automated)
#>
    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.1.1: (L1) Ensure Administrative accounts are cloud-only..."

    $controlNum = "1.1.1"
    $title = "Ensure Administrative accounts are cloud-only"
    $level = "L1"

    try {
        $DirectoryRoles = Get-MgDirectoryRole
        $PrivilegedRoles = $DirectoryRoles | Where-Object {
            $_.DisplayName -like "*Administrator*" -or $_.DisplayName -eq "Global Reader"
        }

        $RoleMembers = foreach ($role in $PrivilegedRoles) {
            Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id | Select-Object Id -Unique
        }

        $PrivilegedUsers = foreach ($rm in $RoleMembers) {
            Get-MgUser -UserId $rm.Id -Property UserPrincipalName, DisplayName, Id, OnPremisesSyncEnabled
        }

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
    # L1
    <#
.LINK
    CIS 1.1.3 (L1, Automated)
#>
    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.1.3: (L1) Ensure between two and four global admins..."

    $controlNum = "1.1.3"
    $title = "Ensure that between two and four global admins are designated"
    $level = "L1"

    try {
        $globalAdminRole = Get-MgDirectoryRole -Filter "RoleTemplateId eq '62e90394-69f5-4237-9190-012177145e10'"
        if (-not $globalAdminRole) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details "Could not locate Global Admin role object."
            return
        }

        $globalAdmins = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id
        $count = $globalAdmins.Count

        if ($count -lt 2) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "Only $count Global Admin(s). Should be at least 2."
        } elseif ($count -gt 4) {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "$count Global Admin(s). Should be no more than 4."
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "$count Global Admin(s) is within the recommended 2-4."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

function Test-114-AdminLicenseReducedFootprint {
    # L1
    <#
.LINK
    CIS 1.1.4 (L1, Automated)
#>
    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.1.4: (L1) Admin accounts with minimal license footprint..."

    $controlNum = "1.1.4"
    $title = "Ensure administrative accounts use licenses with a reduced application footprint"
    $level = "L1"

    try {
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
                # no licenses
                Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "$($pu.UserPrincipalName): unlicensed (good)."
            } elseif (($AssignedSkus -notcontains 'AAD_PREMIUM') -and ($AssignedSkus -notcontains 'AAD_PREMIUM_P2')) {
                Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details "$($pu.UserPrincipalName) has non-minimal license(s): $($AssignedSkus -join ', ')"
            } else {
                Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "$($pu.UserPrincipalName) has license(s): $($AssignedSkus -join ', ')"
            }
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}


#
# -- Level 2 EXAMPLE --
#

function Test-121-OnlyApprovedPublicGroups {
    # L2
    <#
.LINK
    CIS 1.2.1 (L2, Automated)
#>
    [CmdletBinding()]
    param()
    Write-Verbose "Running CIS 1.2.1: (L2) Ensure only organizationally approved public groups exist..."

    $controlNum = "1.2.1"
    $title = "Ensure that only organizationally managed/approved public groups exist"
    $level = "L2"

    try {
        $allGroups = Get-MgGroup -All:$true
        $publicGroups = $allGroups | Where-Object { $_.Visibility -eq "Public" }

        if ($publicGroups) {
            $list = $publicGroups | ForEach-Object { $_.DisplayName } | Sort-Object
            $details = "FAIL: Found public M365 group(s). Check if truly approved: " + ($list -join ", ")
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "FAIL" -Details $details
        } else {
            Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "PASS" -Details "No public groups found (or none unapproved)."
        }
    } catch {
        Add-CISResult -ControlNumber $controlNum -Title $title -Level $level -Status "ERROR" -Details $_.Exception.Message
    }
}

# (Similarly, you would create other L2 checks as needed.)
# --------------------------------------------------------

# --------------------------------------------
# 4. MASTER RUNNER & EXPORT
# --------------------------------------------
function Invoke-CISBenchmarkChecks {
    [CmdletBinding()]
    param(
        [switch]$Level1,
        [switch]$Level2,
        [string]$SharePointAdminUrl = "https://<YourTenant>-admin.sharepoint.com",
        [string]$OutputXlsx = ".\CIS-M365-Checks-$((Get-Date).ToString('yyyyMMdd-HHmmss')).xlsx"
    )

    Write-Host "==========================================="
    Write-Host "Starting CIS Microsoft 365 Benchmarks..."
    Write-Host "==========================================="

    # If user does not explicitly supply either switch, default to L1
    if ((-not $Level1) -and (-not $Level2)) {
        $Level1 = $true
    }

    # Connect to all services first
    Connect-CISModules -SharePointAdminUrl $SharePointAdminUrl

    # If Level1 is specified or Level2 is specified => run L1 checks.
    # (If -Level2 alone is specified, we'll run both L1 & L2.)
    if ($Level1 -or $Level2) {
        # -- L1 checks:
        Test-111-AdminAccountsCloudOnly
        # 1.1.2 is manual so we skip or create a placeholder
        Test-113-TwoToFourGlobalAdmins
        Test-114-AdminLicenseReducedFootprint
        # (Add more L1 checks as needed)
    }

    # If Level2 is specified, run L2 checks as well
    if ($Level2) {
        Test-121-OnlyApprovedPublicGroups
        # (Add more L2 checks as needed)
    }

    Write-Host "`nAll selected checks complete. Found $($Global:CISResults.Count) result records."

    # Export to XLSX (Requires ImportExcel)
    Write-Verbose "Exporting results to Excel file: $OutputXlsx"
    try {
        if (!(Get-Module -Name ImportExcel -ListAvailable)) {
            Write-Verbose "ImportExcel module not found. Installing now..."
            Install-Module ImportExcel -Scope CurrentUser -Force
        }
        Import-Module ImportExcel -ErrorAction Stop

        $Global:CISResults | Export-Excel -Path $OutputXlsx -WorksheetName "CIS Checks" -AutoSize
        Write-Host "Export complete. XLSX file saved to: $OutputXlsx"
    } catch {
        Write-Warning "Error exporting to Excel: $($_.Exception.Message)"
        Write-Warning "Consider exporting to CSV or HTML instead."
    }

    Write-Host "Done."
}
