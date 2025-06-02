<#
.SYNOPSIS
    This script uses PowerShell to automate the collection of domain configuration 
    and settings as part of an AD Domain Discovery.

.DESCRIPTION
    Intended to be utilized as part of a review of the domain's configuration 
    prior to a migration, consolidation, or upgrade. The information is output 
    into Desktop\Domain Discovery\DomainInfo.txt with certain reports generating 
    individual files.

    The script uses Get-DnsServer commands. These require the DNS Server Tools installed,
    and you typically need to run them from a DNS server or a management system that has the
    DNS Server role tools installed. If you lack permissions or the server is missing DNS,
    those sections will gracefully catch the error and note it in the output.

    If you want to target a specific domain/forest, pass -Domain example.corp.local

.NOTES
    Comet Consulting Group ("Comet") Confidential - All Rights Reserved.
    This script contains proprietary information owned by Comet and should 
    be regarded as confidential. 

    Author:  Chris Ibold
    Version: 3.0

    Original reference: 
    https://social.technet.microsoft.com/wiki/contents/articles/38512.active-directory-domain-discovery-checklist.aspx

    -SkipUsers, -SkipComputers, -SkipDNS, -SkipGPO, -SkipReplication let you skip large or optional queries.
-Domain allows you to specify a domain (defaults to $env:USERDNSDOMAIN).
-OutputPath specifies the target folder for all outputs (CSV, HTML, transcript, text file).

If you do not skip replication checks, the script collects replication partner metadata and runs repadmin /replsummary to produce a text summary.
The repadmin command must be available on the system (e.g., with RSAT / AD DS Tools installed).

#>
[CmdletBinding()]
param(
    [string]$Domain = $env:USERDNSDOMAIN,

    # Output folder (default: "Documents\DomainDiscovery"), e.g. "C:\Temp\Discovery"
    [string]$OutputPath = (Join-Path $Home 'Documents\DomainDiscovery'),

    # Skip parameters for large or optional sections
    [switch]$SkipUsers,
    [switch]$SkipComputers,
    [switch]$SkipDNS,
    [switch]$SkipGPO,
    [switch]$SkipReplication,
    [switch]$SkipTrusts
)

###############################################################################
#                         1. Initialize & Logging                             #
###############################################################################

try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Host "ERROR: Unable to load ActiveDirectory module. $($_.Exception.Message)"
    return
}

# Create output directory if not present
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory | Out-Null
}

# Start transcript logging
$TranscriptPath = Join-Path $OutputPath 'DiscoveryTranscript.txt'
try {
    Start-Transcript -Path $TranscriptPath -ErrorAction Stop
} catch {
    Write-Host "Warning: Could not start transcript. $($_.Exception.Message)"
}

# GPO backup subfolder
$GPOFolder = Join-Path $OutputPath 'GPO Backup'
if (-not (Test-Path $GPOFolder)) {
    New-Item -Path $GPOFolder -ItemType Directory | Out-Null
}

# Text file for general domain info
$OutTxtFile = Join-Path $OutputPath 'DomainInfo.txt'

# Basic date info
$Date = Get-Date -f yyyy/MM/dd

# Gather domain/forest info
$ForestInfo = Get-ADForest -Server $Domain
$ForestFQDN = $ForestInfo.Name
$DomainInfo = Get-ADDomain -Server $Domain

# Write header to DomainInfo.txt
"#################################################################################################" | Out-File $OutTxtFile
"Active Directory Domain Discovery - Enhanced Script" | Out-File $OutTxtFile -Append
"Date: $Date" | Out-File $OutTxtFile -Append
"Domain: $Domain" | Out-File $OutTxtFile -Append
"Output Path: $OutputPath" | Out-File $OutTxtFile -Append
"Transcript Log: $TranscriptPath" | Out-File $OutTxtFile -Append
"#################################################################################################`n" | Out-File $OutTxtFile -Append

###############################################################################
#                         2. Helper Functions                                 #
###############################################################################

function Write-Info {
    param([string]$Message)
    $Message | Out-File $OutTxtFile -Append
    Write-Host $Message
}

function Export-CSVandHTML {
    <#
    .SYNOPSIS
        Exports a PowerShell object to both CSV and HTML files.

    .PARAMETER Data
        The collection of objects (e.g., from a pipeline) to export.

    .PARAMETER BaseFileName
        The full path and base name (no extension) for the output files.
        e.g., "C:\Temp\ADUsers" -> Exports C:\Temp\ADUsers.csv and ...\ADUsers.html

    .PARAMETER Title
        Optional HTML title to embed in the generated HTML.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $Data,
        [Parameter(Mandatory = $true)] [string]$BaseFileName,
        [string]$Title = "Export Report"
    )

    try {
        # Export to CSV
        $csvPath = "$BaseFileName.csv"
        $Data | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Info "Exported CSV: $csvPath"

        # Export to HTML
        $htmlPath = "$BaseFileName.html"
        $htmlContent = $Data | 
            ConvertTo-Html -Title $Title -PreContent "<h2>$Title (Generated $(Get-Date))</h2>"
        $htmlContent | Out-File $htmlPath
        Write-Info "Exported HTML: $htmlPath"
    } catch {
        Write-Info "ERROR exporting CSV/HTML for $($BaseFileName): $($_.Exception.Message)"
    }
}

function SafeExecute {
    <#
    .SYNOPSIS
        Safely executes a script block inside a try/catch,
        and logs any error without stopping the entire script.
    #>
    param(
        [ScriptBlock]$Block,
        [string]$ErrorContext = "General Section"
    )
    try {
        & $Block
    } catch {
        Write-Info "ERROR in [$ErrorContext]: $($_.Exception.Message)"
    }
}

#------------------------------------------------------------------------------
# Progress Bar Helpers
#------------------------------------------------------------------------------

# Calculate total steps based on skip parameters
[int]$TotalSteps = 0
$TotalSteps += 1 # Domain/Forest Info
$TotalSteps += 1 # Default Domain Password Policy
if (-not $SkipReplication) { $TotalSteps += 1 } # Replication
if (-not $SkipTrusts) { $TotalSteps += 1 } # Trusts
$TotalSteps += 1 # Sites/Subnets
$TotalSteps += 1 # Domain Controllers
$TotalSteps += 1 # FSMO Roles
$TotalSteps += 1 # OU Structure
if (-not $SkipUsers) { $TotalSteps += 1 } # AD Users (incl. stale users)
if (-not $SkipComputers) { $TotalSteps += 1 } # AD Computers
$TotalSteps += 1 # AD Groups (incl. privileged)
$TotalSteps += 1 # Domain Admins
$TotalSteps += 1 # Service Accounts
if (-not $SkipGPO) { $TotalSteps += 1 } # GPO
$TotalSteps += 1 # PSOs
$TotalSteps += 1 # AD Optional Features
if (-not $SkipDNS) { $TotalSteps += 1 } # DNS

# We'll use a counter to increment each time a step is started
[int]$StepCounter = 0

function Update-Progress {
    param([string]$Activity)
    $StepCounter++
    Write-Progress -Activity "Domain Discovery" -Status $Activity -PercentComplete (($StepCounter / $TotalSteps) * 100)
}

###############################################################################
#                         3. Collect Data & Export                            #
###############################################################################

#-------------------------------------------
# 1) Basic Domain / Forest Info
#-------------------------------------------
Update-Progress "Collecting Domain / Forest Info..."
SafeExecute -ErrorContext "Domain/Forest Info" -Block {
    Write-Info "================= DOMAIN / FOREST INFO ================="
    Write-Info ("Forest FQDN: {0}" -f $ForestFQDN)
    Write-Info ("Forest Functional Level: {0}" -f $ForestInfo.ForestMode)
    Write-Info ("Domain FQDN: {0}" -f $DomainInfo.Name)

    Write-Info "All Domains in Forest:"
    $ForestInfo.Domains | ForEach-Object { Write-Info " - $_" }
    Write-Info ""
}

#-------------------------------------------
# 2) Default Domain Password Policy
#-------------------------------------------
Update-Progress "Retrieving Default Domain Password Policy..."
SafeExecute -ErrorContext "Default Domain Password Policy" -Block {
    Write-Info "================= DEFAULT DOMAIN PASSWORD POLICY ================="
    $defaultDomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy -Server $Domain
    $defaultDomainPasswordPolicy | Out-File $OutTxtFile -Append
    Write-Info ""
}

#-------------------------------------------
# 3) AD Replication (optional)
#-------------------------------------------
if (-not $SkipReplication) {
    Update-Progress "Checking AD Replication..."
    SafeExecute -ErrorContext "Replication Checks" -Block {
        Write-Info "================= REPLICATION INFO ================="
        
        # Method 1: Basic replication metadata
        $replMeta = Get-ADReplicationPartnerMetadata -Target $DomainInfo.Name -Scope Domain
        Export-CSVandHTML -Data $replMeta -BaseFileName (Join-Path $OutputPath "ReplicationPartnerMetadata") -Title "AD Replication Partner Metadata"

        # Method 2: repadmin /replsummary for a quick summary
        # Note: repadmin must be available in the environment.
        $repadminSummary = & repadmin /replsummary
        $repadminSummary | Out-File (Join-Path $OutputPath "RepAdminReplSummary.txt")
        Write-Info "Replication summary has been saved to RepAdminReplSummary.txt"
        Write-Info ""
    }
}

#-------------------------------------------
# 4) Trust Discovery (optional)
#-------------------------------------------
if (-not $SkipTrusts) {
    Update-Progress "Discovering Domain Trusts..."
    SafeExecute -ErrorContext "Domain Trusts" -Block {
        Write-Info "================= TRUST DISCOVERY (Get-ADTrust) ================="
        # This retrieves basic information about each trust in the forest
        $trusts = Get-ADTrust -Filter * -Server $Domain

        # Just an example of selecting specific properties
        # Feel free to expand as needed
        $trustsSelect = $trusts | Select-Object Name, Direction, TrustType, TGTDelegation, IsForest, IsExternal, IsRealm

        Export-CSVandHTML -Data $trustsSelect -BaseFileName (Join-Path $OutputPath "DomainTrusts") -Title "Domain Trusts"
        Write-Info "Trust information has been exported to CSV/HTML."
        Write-Info ""
    }
}

#-------------------------------------------
# 5) Sites & Subnets
#-------------------------------------------
Update-Progress "Enumerating Sites and Subnets..."
SafeExecute -ErrorContext "Sites & Subnets" -Block {
    Write-Info "================= SITES & SUBNETS ================="
    $Sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites
    $SitesSubnets = foreach ($Site in $Sites) {
        foreach ($Subnet in $Site.Subnets) {
            [PSCustomObject]@{
                Site   = $Site.Name
                Subnet = $Subnet
            }
        }
    }
    Export-CSVandHTML -Data $SitesSubnets -BaseFileName (Join-Path $OutputPath "Sites") -Title "AD Sites and Subnets"
    Write-Info ""
}

#-------------------------------------------
# 6) Domain Controllers
#-------------------------------------------
Update-Progress "Enumerating Domain Controllers..."
SafeExecute -ErrorContext "Domain Controllers" -Block {
    Write-Info "================= DOMAIN CONTROLLERS ================="
    $DCs = Get-ADDomainController -Filter * -Server $Domain | 
        Select-Object Name, IPv4Address, OperatingSystem, Site
    Export-CSVandHTML -Data $DCs -BaseFileName (Join-Path $OutputPath "DomainControllers") -Title "Domain Controllers"
    Write-Info ""
}

#-------------------------------------------
# 7) FSMO Roles
#-------------------------------------------
Update-Progress "Retrieving FSMO Role Holders..."
SafeExecute -ErrorContext "FSMO Roles" -Block {
    Write-Info "================= FSMO ROLES ================="
    Write-Info ("Schema Master: {0}" -f $ForestInfo.SchemaMaster)
    Write-Info ("Domain Naming Master: {0}" -f $ForestInfo.DomainNamingMaster)
    $domainObj = Get-ADDomain -Server $Domain
    Write-Info ("Infrastructure Master: {0}" -f $domainObj.InfrastructureMaster)
    Write-Info ("PDC Emulator: {0}" -f $domainObj.PDCEmulator)
    Write-Info ("RID Master: {0}" -f $domainObj.RIDMaster)
    Write-Info ""
}

#-------------------------------------------
# 8) OU Structure
#-------------------------------------------
Update-Progress "Enumerating OU Structure..."
SafeExecute -ErrorContext "OU Structure" -Block {
    Write-Info "================= OU STRUCTURE ================="
    $OUs = Get-ADOrganizationalUnit -Filter 'name -like "*"' -Server $Domain |
        Select-Object Name, DistinguishedName
    Export-CSVandHTML -Data $OUs -BaseFileName (Join-Path $OutputPath "OUStructure") -Title "Organizational Units"
    Write-Info ""
}

#-------------------------------------------
# 9) AD Users (option to skip)
#-------------------------------------------
if (-not $SkipUsers) {
    Update-Progress "Collecting AD Users..."
    SafeExecute -ErrorContext "All AD Users" -Block {
        Write-Info "================= AD USERS ================="
        $users = Get-ADUser -Filter * -Properties * -Server $Domain
        $UserCount = $users.Count
        Write-Info "User Count: $UserCount"
        Export-CSVandHTML -Data $users -BaseFileName (Join-Path $OutputPath "ADUsers") -Title "AD Users"

        # Stale Users (e.g. older than 90 days)
        Write-Info "----------------- Stale Users (90+ Days) -----------------"
        $Cutoff = (Get-Date).AddDays(-90)
        $staleUsers = $users | Where-Object {
            ($_.LastLogonDate -and $_.LastLogonDate -lt $Cutoff) -or
            ($_.PasswordLastSet -lt $Cutoff)
        }
        Export-CSVandHTML -Data $staleUsers -BaseFileName (Join-Path $OutputPath "StaleUsers") -Title "Stale AD Users"
        Write-Info "Stale user count: $($staleUsers.Count)"
        Write-Info ""
    }
}

#-------------------------------------------
# 10) AD Computers (option to skip)
#-------------------------------------------
if (-not $SkipComputers) {
    Update-Progress "Collecting AD Computers..."
    SafeExecute -ErrorContext "AD Computers" -Block {
        Write-Info "================= AD COMPUTERS ================="
        $computers = Get-ADComputer -Filter * -Properties * -Server $Domain
        $ComputerCount = $computers.Count
        Write-Info "Computer Count: $ComputerCount"
        Export-CSVandHTML -Data $computers -BaseFileName (Join-Path $OutputPath "ADComputers") -Title "AD Computers"
        Write-Info ""
    }
}

#-------------------------------------------
# 11) AD Groups
#-------------------------------------------
Update-Progress "Collecting AD Groups..."
SafeExecute -ErrorContext "AD Groups" -Block {
    Write-Info "================= AD GROUPS ================="
    $groups = Get-ADGroup -Filter * -Properties * -Server $Domain
    $GroupCount = $groups.Count
    Write-Info "Group Count: $GroupCount"
    Export-CSVandHTML -Data $groups -BaseFileName (Join-Path $OutputPath "ADGroups") -Title "AD Groups"
    Write-Info ""

    # Privileged Groups (AdminCount=1)
    Write-Info "----------------- Privileged Groups (AdminCount=1) -----------------"
    $privGroups = Get-ADGroup -Filter 'AdminCount -eq 1' -Properties * -Server $Domain
    Export-CSVandHTML -Data $privGroups -BaseFileName (Join-Path $OutputPath "PrivilegedGroups") -Title "Privileged Groups"
    Write-Info ""
}

#-------------------------------------------
# 12) Domain Admins
#-------------------------------------------
Update-Progress "Collecting Domain Admins..."
SafeExecute -ErrorContext "Domain Admins" -Block {
    Write-Info "================= DOMAIN ADMINS ================="
    $domainAdmins = Get-ADGroupMember -Identity "Domain Admins" -Server $Domain
    Export-CSVandHTML -Data $domainAdmins -BaseFileName (Join-Path $OutputPath "DomainAdmins") -Title "Domain Admins"
    Write-Info ("Domain Admins Count: {0}" -f $domainAdmins.Count)
    Write-Info ""
}

#-------------------------------------------
# 13) Service Accounts
#-------------------------------------------
Update-Progress "Collecting Service Accounts..."
SafeExecute -ErrorContext "Service Accounts" -Block {
    Write-Info "================= SERVICE ACCOUNTS ================="
    
    # Regular "svc"/"service" accounts
    $svcUsers = Get-ADUser -Filter { (Name -like "*svc*") -or (Name -like "*service*") } -Properties * -Server $Domain
    Export-CSVandHTML -Data $svcUsers -BaseFileName (Join-Path $OutputPath "ServiceAccounts-User") -Title "User-based Service Accounts"
    Write-Info ("Found {0} user-based service accounts." -f $svcUsers.Count)

    # gMSA
    $gMSAs = Get-ADServiceAccount -Filter * -Properties * -Server $Domain
    Export-CSVandHTML -Data $gMSAs -BaseFileName (Join-Path $OutputPath "ServiceAccounts-gMSA") -Title "gMSA Accounts"
    Write-Info ("Found {0} gMSA accounts." -f $gMSAs.Count)
    Write-Info ""
}

#-------------------------------------------
# 14) GPO Backup & Discovery (option to skip)
#-------------------------------------------
if (-not $SkipGPO) {
    Update-Progress "Collecting and Backing Up GPOs..."
    SafeExecute -ErrorContext "GPO Backup and Discovery" -Block {
        Write-Info "================= GROUP POLICY OBJECTS ================="
        $GPOCount = (Get-GPO -All -Domain $ForestFQDN -Server $Domain).Count
        Write-Info ("GPO Count: {0}" -f $GPOCount)
        Write-Info "Backing up all GPOs..."

        Backup-GPO -All -Path $GPOFolder | Out-Null

        Write-Info "Exporting list of GPOs to CSV + HTML..."
        $allGPOs = Get-GPO -All -Domain $ForestFQDN -Server $Domain | 
            Sort-Object DisplayName |
            Select-Object DisplayName, Id, Owner, GpoStatus
        Export-CSVandHTML -Data $allGPOs -BaseFileName (Join-Path $OutputPath "DomainGPOs") -Title "Domain GPOs"

        Write-Info "Generating full HTML GPO Report..."
        Get-GPOReport -Domain $ForestFQDN -All -ReportType Html -Path (Join-Path $OutputPath "FullGPOReport.html")

        Write-Info "Compressing GPO backup folder..."
        $ArchiveFullPath = Join-Path $OutputPath 'GPOBackup.zip'
        Compress-Archive -Path $GPOFolder -DestinationPath $ArchiveFullPath

        Write-Info "Removing uncompressed GPO backup folder..."
        Remove-Item -Path $GPOFolder -Recurse -Force
        Write-Info "GPO backup compressed to: $ArchiveFullPath"
        Write-Info ""
    }
}

#-------------------------------------------
# 15) PSOs (Fine-Grained Password Policies)
#-------------------------------------------
Update-Progress "Collecting Fine-Grained Password Policies..."
SafeExecute -ErrorContext "Fine-Grained Password Policies (PSOs)" -Block {
    Write-Info "================= FINE-GRAINED PASSWORD POLICIES (PSOs) ================="
    $pso = Get-ADFineGrainedPasswordPolicy -Filter * -Properties * -Server $Domain
    Export-CSVandHTML -Data $pso -BaseFileName (Join-Path $OutputPath "DomainPSOs") -Title "Fine-Grained Password Policies"
    Write-Info ("PSO Count: {0}" -f $pso.Count)
    Write-Info ""
}

#-------------------------------------------
# 16) AD Optional Features
#-------------------------------------------
Update-Progress "Collecting AD Optional Features..."
SafeExecute -ErrorContext "AD Optional Features" -Block {
    Write-Info "================= AD OPTIONAL FEATURES ================="
    $adFeatures = Get-ADOptionalFeature -Filter * -Properties * -Server $Domain |
        Select-Object Name, EnabledScopes
    Export-CSVandHTML -Data $adFeatures -BaseFileName (Join-Path $OutputPath "ADOptionalFeatures") -Title "AD Optional Features"
    Write-Info ""
}

#-------------------------------------------
# 17) DNS Info (option to skip)
#-------------------------------------------
if (-not $SkipDNS) {
    Update-Progress "Collecting DNS Configuration..."
    SafeExecute -ErrorContext "DNS Info" -Block {
        Write-Info "================= DNS INFORMATION ================="

        # DNS Forwarders
        Write-Info "---- DNS Forwarders ----"
        try {
            $dnsForwarders = Get-DnsServerForwarder
            Export-CSVandHTML -Data $dnsForwarders -BaseFileName (Join-Path $OutputPath "DNSForwarders") -Title "DNS Forwarders"
        } catch {
            Write-Info "ERROR: Unable to retrieve DNS Forwarders - $($_.Exception.Message)"
        }

        # DNS A Records for forest root zone
        Write-Info "---- DNS A Records ($ForestFQDN) ----"
        try {
            $dnsARecords = Get-DnsServerResourceRecord -ZoneName $ForestFQDN -RRType A
            Export-CSVandHTML -Data $dnsARecords -BaseFileName (Join-Path $OutputPath "DNS-A-Records") -Title "DNS A Records"
        } catch {
            Write-Info "ERROR: Unable to retrieve DNS A records for $ForestFQDN - $($_.Exception.Message)"
        }

        # DNS Zones
        Write-Info "---- DNS Zones ----"
        try {
            $dnsZones = Get-DnsServerZone
            Export-CSVandHTML -Data $dnsZones -BaseFileName (Join-Path $OutputPath "DNSZones") -Title "DNS Zones"
        } catch {
            Write-Info "ERROR: Unable to retrieve DNS Zones - $($_.Exception.Message)"
        }

        # DNS Zone Aging (Scavenging)
        Write-Info "---- DNS Scavenging (Zone Aging) ----"
        try {
            $dnsZoneAging = Get-DnsServerZoneAging -Name $ForestFQDN
            $dnsZoneAging | Out-File $OutTxtFile -Append
            Write-Info "(DNS scavenging info appended to DomainInfo.txt)"
        } catch {
            Write-Info "ERROR: Unable to retrieve DNS Zone Aging for $ForestFQDN - $($_.Exception.Message)"
        }
        Write-Info ""
    }
}

###############################################################################
#                         4. Finalization                                     #
###############################################################################

# Mark the progress as completed
Write-Progress -Activity "Domain Discovery" -Status "Complete" -Completed

Write-Info "================= SCRIPT COMPLETE ================="
Write-Info "All requested data has been collected in $OutputPath."
Write-Info "View $OutTxtFile for a textual summary/log."
Write-Host "See transcript at: $TranscriptPath" 

# Stop transcript
try {
    Stop-Transcript
} catch {
    Write-Host "Warning: Unable to stop transcript. $($_.Exception.Message)"
}
