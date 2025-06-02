<###################################################################################################################
##	This script contains proprietary information owned by Sterling Consulting Company ("Sterling") and should be
##  regarded as confidential. This script, any output and files, related information, and all copies of same remain
##  the confidential property of Sterling and shall be returned to Sterling upon request.
##
##	These materials and the information contained herein are not to be duplicated or used, in whole or in part, for
##  any purpose other than for the purposes for Sterling and their clients
####################################################################################################################

.SYNOPSIS
This script uses PowerShell to automate the collection of domain configuration and settings as part of an AD Domain 
Discovery.

.DESCRIPTION
Intended to be utilized as part of a review of the domain's configuration prior to a migration, consolidation or
upgrade. The information is output into Desktop\Domain Discovery\DomainInfo.txt with certain reports generating 
individual files.

.NOTES
Author: Chris Ibold
Creation Date: 2023-03-8
Version: 1.0 (2023-03-8)
    * Initial creation and testing.

Version: 2.0 (2023-03-17)
    * Large overhaul of a number of systems
    * Automated output
    * Separated certain discovery items into a text file using Out-File that just don't make sense in a CSV
    * Moved all relavant discovery to their own CSV files.
    * Gave the GPO backup its own directory
    * Removed problematic variables that were calling built-in variables incorrectly

Version 2.5 (2023-03-20)
    * Added a progress bar helper and steps for each section.

Version 2.5.1 (2023-03-21)
    * Corrected a Format-Table to Select-Object instance left from testing.

Version 2.6 (2023-03-23)
    * Spatted Out-File variables.
    * Script now compresses the GPO Backup folder and then deletes the uncompressed folder for easier transport.

Version 2.7 (2023-03-26)
    * Better script readibility. Defined Initializations, Declarations, Function and Execution.

##################################################################################################################>

Import-Module ActiveDirectory

#---------------------------------------------------------[Initializations]--------------------------------------------------------
$stepCounter = 0
$script:steps = ([System.Management.Automation.PsParser]::Tokenize((Get-Content "$PSScriptRoot\$($MyInvocation.MyCommand.Name)"), [ref]$null) | Where-Object { $_.Type -eq 'Command' -and $_.Content -eq 'Write-ProgressHelper' }).Count


#----------------------------------------------------------[Declarations]----------------------------------------------------------
## Variables
$OutDirectory = [Environment]::GetFolderPath("Desktop") + "\Domain Discovery"
$GPOFolder = "$OutDirectory\GPO Backup"
$OutTxtFile = "$OutDirectory\DomainInfo.txt"

$Date = Get-Date -f yyyy/MM/dd
$ForestInfo = Get-ADForest $env:USERDNSDOMAIN
$ForestFQDN = $ForestInfo.Name
$DomainPDC = $ForestInfo.SchemaMaster
$DomainInfo = Get-ADDomain $env:USERDNSDOMAIN
$DomainDN = $DomainInfo.Name

## Create output directory
New-Item -Path $OutDirectory -ItemType Directory
New-Item -Path $GPOFolder -ItemType Directory

#Splat Out-File params
$OutFile = @{
    Append   = $true
    Filepath = "$OutTxtFile"
}

## Output text file
"The information below is the output of the this script. This script uses PowerShell to automate the" | Out-File @OutFile
"collection of information as outlined in the TechNet article, Active Directory Domain Discovery Checklist." | Out-File @OutFile
"https://social.technet.microsoft.com/wiki/contents/articles/38512.active-directory-domain-discovery-checklist.aspx`n" | Out-File @OutFile
"Created: $Date" | Out-File @OutFile
"########################################################################################################################`n" | Out-File @OutFile


#-----------------------------------------------------------[Functions]------------------------------------------------------------
# Create the Progress bar function
function Write-ProgressHelper {
    param (
        [int]$StepNumber,
        [string]$Message
    )

    Write-Progress -Activity 'Running Domain Discovery' -Status $Message -PercentComplete (($StepNumber / $steps) * 100)
}


#-----------------------------------------------------------[Execution]------------------------------------------------------------
## Domain Name
"Fully Qualified Domain Name (FQDN): $env:USERDNSDOMAIN" | Out-File @OutFile

#### Forest Functional Level
$ForestFunctionalLevel = $ForestInfo.ForestMode
"Forest Functional Level: $ForestFunctionalLevel" | Out-File @OutFile

#### Forest Architecture
$ForestDomains = $ForestInfo.Domains
$ForestCount = $ForestDomains.Count
$ChildDomains = $DomainInfo.ChildDomains
$ChildCount = $ChildDomains.Count
"Domains within Forest: <$ForestCount>" | Out-File @OutFile
"Parent Domain: $ForestFQDN" | Out-File @OutFile
"Child Domain(s): <$ChildCount> $ChildDomains" | Out-File @OutFile
"########################################################################################################################`n" | Out-File @OutFile

## Domain Trust(s)
Write-ProgressHelper -Message 'Discovering Trusts...' -StepNumber ($stepCounter++)
Function Set-TrustAttributes {
    [cmdletbinding()]
    Param(
        [parameter(Mandatory = $false, ValueFromPipeline = $True)]
        [int32]$Value
    )
    If ($value) {
        $input = $value
    }
    [String[]]$TrustAttributes = @() 
    Foreach ($key in $input) {

        if ([int32]$key -band 0x00000001) { $TrustAttributes += "Non Transitive" } 
        if ([int32]$key -band 0x00000002) { $TrustAttributes += "UpLevel" } 
        if ([int32]$key -band 0x00000004) { $TrustAttributes += "Quarantine (SID Filtering enabled)" } #SID Filtering 
        if ([int32]$key -band 0x00000008) { $TrustAttributes += "Forest Transitive" } 
        if ([int32]$key -band 0x00000010) { $TrustAttributes += "Cross Organization (Selective Authentication enabled)" } #Selective Auth 
        if ([int32]$key -band 0x00000020) { $TrustAttributes += "Within Forest" } 
        if ([int32]$key -band 0x00000040) { $TrustAttributes += "Treat as External" } 
        if ([int32]$key -band 0x00000080) { $TrustAttributes += "Uses RC4 Encryption" }
    } 
    return $trustattributes
}
Try { $TrustQuery = Get-WmiObject -Class Microsoft_DomainTrustStatus -Namespace root\microsoftactivedirectory -ComputerName $DomainPDC -ErrorAction SilentlyContinue }
Catch { $_ }
If ($TrustQuery) {
    $TrustOutput = $TrustQuery | Select-Object -Property @{L = "Trusted Domain"; e = { $_.TrustedDomain } }, @{L = "Trusts Direction"; e = { switch ($_.TrustDirection) {
                "1" { "Inbound" }
                "2" { "Outbound" }
                "3" { "Bi-directional" }
                Default { "N/A" }
            } }
    }, @{L = "Trusts Attributes"; e = { ($_.TrustAttributes | Set-TrustAttributes) } }
} 
$TrustOutput | Out-File -Append -FilePath $OutTxtFile
"########################################################################################################################`n" | Out-File @OutFile

## Sites
Write-ProgressHelper -Message 'Discovering Sites...' -StepNumber ($stepCounter++)
$Sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites
$SitesSubnets = @()
foreach ($Site in $Sites) {
    foreach ($Subnet in $Site.Subnets) {
        $SitesTemp = New-Object PSCustomObject -Property @{
            'Site'   = $Site.Name
            'Subnet' = $Subnet 
        }
        $SitesSubnets += $SitesTemp
    }
} $SitesSubnets | Export-Csv -Path "$OutDirectory\Sites.csv" -NoTypeInformation

## Domain Controller(s)
Write-ProgressHelper -Message 'Discovering Domain Controllers...' -StepNumber ($stepCounter++)
Get-ADDomainController -Filter * | Select-Object Name, ipv4Address, OperatingSystem, site | Sort-Object -Property Name | Export-Csv -Path "$OutDirectory\DomainControllers.csv" -NoTypeInformation

## FSMO Role Holder(s)
Write-ProgressHelper -Message 'Discovering FSMO Role Holders...' -StepNumber ($stepCounter++)
"SchemaMaster: $DomainPDC" | Out-File -Append -FilePath $OutTxtFile
"FSMO Role Holder(s):" | Out-File @OutFile
Get-ADDomain | Select-Object InfrastructureMaster, PDCEmulator, RIDMaster | Out-File -Append -FilePath $OutTxtFile
"########################################################################################################################`n" | Out-File @OutFile

## OU Structure
Write-ProgressHelper -Message 'Discovering OU Structure...' -StepNumber ($stepCounter++)
Get-ADOrganizationalUnit -Filter 'Name -like "*"' | Select-Object Name, DistinguishedName | Export-Csv -Path "$OutDirectory\OUStructure.csv" -NoTypeInformation

## AD Objects
Write-ProgressHelper -Message 'Discovering AD Objects...' -StepNumber ($stepCounter++)
"ADObjects" | Out-File @OutFile

## AD Users
Write-ProgressHelper -Message 'Discovering AD Users...' -StepNumber ($stepCounter++)
$UserCount = (Get-ADUser -Filter *).count
"Users (all): <$UserCount>" | Out-File @OutFile
"Users information has been exported to: ADUsers.csv" | Out-File @OutFile
Get-ADUser -Filter * -Properties * | Select-Object Name, CanonicalName, Description, DistinguishedName, EmailAddress, LastLogon, PasswordNeverExpires, ObjectGUID, ObjectSID, PasswordExpired, PasswordLastSet, WhenCreated | Export-Csv -Path "$OutDirectory\ADUsers.csv" -NoTypeInformation

## AD Groups
Write-ProgressHelper -Message 'Discovering AD Groups...' -StepNumber ($stepCounter++)
$GroupCount = (Get-ADGroup -Filter *).count
"Groups: <$GroupCount>" | Out-File @OutFile
"Groups information has been exported to: ADGroups.csv" | Out-File @OutFile
Get-ADGroup -Filter * -Properties * | Select-Object Name, CanonicalName, Description, DistinguishedName, EmailAddress, GroupCategory, GroupScope, ManagedBy, ObjectGUID, ObjectSID, WhenCreated | Export-Csv -Path "$OutDirectory\ADGroups.csv" -NoTypeInformation

## Privileged Groups
Write-ProgressHelper -Message 'Discovering Privileged Groups...' -StepNumber ($stepCounter++)
$PrivilegedGroupsCount = (Get-ADGroup -Filter 'AdminCount -eq 1').count
"Privileged Groups: <$PrivilegedGroupsCount>" | Out-File @OutFile
"Privileged Groups information has been exported to: PrivilegedGroups.csv" | Out-File @OutFile
Get-ADGroup -Filter 'AdminCount -eq 1' -Properties * | Select-Object Name, CanonicalName, DistinguishedName, MemberOf, Members, ObjectGUID, ObjectSID | Export-Csv -Path "$OutDirectory\ADGroups.csv" -NoTypeInformation

## Domain Admins
Write-ProgressHelper -Message 'Discovering Domain Admins...' -StepNumber ($stepCounter++)
$DomainAdminCount = (Get-ADGroupMember -Identity "Domain Admins").count
"Domain Admins: <$DomainAdminCount>" | Out-File @OutFile
"Domain Admin Group Members has been exported to: DomainAdmins.csv" | Out-File @OutFile
Get-ADGroupMember -Identity "Domain Admins" | Export-Csv -Path "$OutDirectory\DomainAdmins.csv" -NoTypeInformation

## AD Computer Objects
Write-ProgressHelper -Message 'Discovering AD Computer Objects...' -StepNumber ($stepCounter++)
$ComputerCount = (Get-ADComputer -Filter *).count
"Computer Objects: <$ComputerCount>" | Out-File @OutFile
"Computer Objects information has been exported to: $ADComputers.csv" | Out-File @OutFile
Get-ADComputer -Filter * -Properties * | Select-Object Name, CanonicalName, Description, DestinguishedName, DNSHostName, ObjectGUID, ObjectSID, OperatingSystem, OperatingSystemVersion, PrimaryGroup | Export-Csv -Path "$OutDirectory\ADComputers.csv" -NoTypeInformation

## AD Service Accounts - User
Write-ProgressHelper -Message 'Discovering AD Service Accounts (Users)...' -StepNumber ($stepCounter++)
$ServiceUserCount = (Get-ADUser -Filter { (Name -like "*svc*") -or (Name -like "*service*") }).count
"Service Accounts (User, filtered for svc or service in name): <$ServiceUserCount>" | Out-File @OutFile
"Service Accounts (User) information has been exported to: ServiceAccounts-User.csv" | Out-File @OutFile
Get-ADUser -Filter { (Name -like "*svc*") -or (Name -like "*service*") } -Properties * | Select-Object Name, CanonicalName, Description, DistinguishedName, EmailAddress, LastLogon, PasswordNeverExpires, ObjectGUID, ObjectSID, PasswordExpired, PasswordLastSet, WhenCreated | Export-Csv -Path "$OutDirectory\ServiceAccounts-User.csv" -NoTypeInformation

## AD Service Accounts - gMSA
Write-ProgressHelper -Message 'Discovering AD Service Accounts (gMSA)...' -StepNumber ($stepCounter++)
"Service Accounts (gMSA) information has been exported to: ServiceAccounts-gMSA.csv" | Out-File @OutFile
Get-ADServiceAccount -Filter * -Properties * | Select-Object Name, CanonicalName, Description, DistinguishedName, PrimaryGroup, PrincipalsAllowedToRetrieveManagedPassword, ObjectGUID, ObjectSID, WhenCreated | Export-Csv -Path "$OutDirectory\ServiceAccounts-gMSA.csv" -NoTypeInformation
"########################################################################################################################`n" | Out-File @OutFile

## GPOs
$GPOCount = (Get-GPO -All).Count
"Group Policy Objects: <$GPOCount>" | Out-File @OutFile
"GPOs have been backed up to the folder: $GPOFolder" | Out-File @OutFile
"GPOs have been exported to: DomainGPOs.csv" | Out-File @OutFile
"A full GPO HTML report has been generated to: FullGPOReport.html" | Out-File @OutFile
Write-ProgressHelper -Message 'Backing up GPOs...' -StepNumber ($stepCounter++)
Backup-Gpo -All -Path "$GPOFolder\"
Write-ProgressHelper -Message 'Discovering GPOs...' -StepNumber ($stepCounter++)
Get-GPO -All -Domain $ForestFQDN -InformationVariable * | Sort-Object DisplayName | Select-Object DisplayName, Id, Owner, GpoStatus | Export-Csv -Path "$OutDirectory\DomainGPOs.csv" -NoTypeInformation
Write-ProgressHelper -Message 'Generating HTML GPO Report...' -StepNumber ($stepCounter++)
Get-GPOReport -Domain $ForestFQDN -All -ReportType Html -Path "$OutDirectory\FullGPOReport.html"
"########################################################################################################################`n" | Out-File @OutFile

## Password Security Objects
Write-ProgressHelper -Message 'Discovering Password Security Objects...' -StepNumber ($stepCounter++)
$PSOCount = (Get-ADFineGrainedPasswordPolicy -Filter *).count
"Password Security Objects (PSOs): <$PSOCount>" | Out-File @OutFile
"PSOs have been exported to: DomainPSOs.csv"
Get-ADFineGrainedPasswordPolicy -Filter * -Properties * | Select-Object Name, Description, DistinguishedName, ObjectGUID, AppliesTo, Precedence, ComplexityEnabled, LockoutDuration, LockoutObservationWindow, LockoutThreshold, MaxPasswordAge, MinPasswordAge, MinPasswordLength, PasswordHistoryCount, ReversibleEncryptionEnabled | Export-Csv -Path "$OutDirectory\DomainPSOs.csv" -NoTypeInformation
"########################################################################################################################`n" | Out-File @OutFile

## Active Directory Features
Write-ProgressHelper -Message 'Discovering Optional AD Features...' -StepNumber ($stepCounter++)
Get-ADOptionalFeature -Filter * -Properties * | Select-Object Name, EnabledScopes | Export-Csv -Path "$OutDirectory\ADOptionalFeatures.csv" -NoTypeInformation

## DNS
Write-ProgressHelper -Message 'Discovering DNS...' -StepNumber ($stepCounter++)
"DNS Information" | Out-File @OutFile
"DNS Forwarders:" | Out-File @OutFile
Get-DnsServerForwarder -InformationVariable * | Select-Object IPAddress | Out-File @OutFile

#### DNS A Records
Write-ProgressHelper -Message 'Discovering DNS A-Records...' -StepNumber ($stepCounter++)
$DNSRecordCount = (Get-DnsServerResourceRecord -ZoneName $ForestFQDN -RRType "A").Count
"DNS A Records: <$DNSRecordCount>" | Out-File @OutFile
"DNS A Record information has been exported to: DNS-A-Records.csv`n" | Out-File @OutFile
Get-DnsServerResourceRecord -ZoneName $ForestFQDN -RRType "A" | Export-Csv -Path "$OutDirectory\DNS-A-Records.csv" -NoTypeInformation

#### DNS ServerZones
Write-ProgressHelper -Message 'Discovering DNS Server Zones...' -StepNumber ($stepCounter++)
$DNSZoneCount = (Get-DnsServerZone -InformationVariable *).Count
"DNS Zones: <$DNSZoneCount>" | Out-File @OutFile
"DNS Zones have been exported to: DNSZones.csv" | Out-File @OutFile
Get-DnsServerZone -InformationVariable * | Export-Csv -Path "$OutDirectory\DNSZones.csv" -NoTypeInformation

#### DNS Scavenging
Write-ProgressHelper -Message 'DNS Scavenging...' -StepNumber ($stepCounter++)
"DNS Scavenging:" | Out-File @OutFile
Get-DnsServerZoneAging -Name $ForestFQDN | Out-File @OutFile

#Compress GPO Backup folder and then delete the uncompressed folder
$ArchiveFileName = "GPOBackup.zip"
$ArchiveFullPath = Join-Path -Path $OutDirectory -ChildPath $ArchiveFileName
Write-ProgressHelper -Message 'Compressing GPO backup folder...' -StepNumber ($stepCounter++)
Compress-Archive -Path $GPOFolder -DestinationPath $ArchiveFullPath
Write-ProgressHelper -Message 'Cleaning up uncompressed GPO backup folder...' -StepNumber ($stepCounter++)
Remove-Item -Path $GPOFolder -Recurse -Force

Write-Progress -Activity "Discovering Domain" -Status "Complete" -Completed

"########################################################################################################################" | Out-File @OutFile
"######### End of file ######### End of file ######### End of file ######### End of file ######### End of file ##########" | Out-File @OutFile
"########################################################################################################################" | Out-File @OutFile