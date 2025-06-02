# Author: Chris Ibold
# Company: Sterling Consulting
# Version: 1.0.0

# Import required modules
Import-Module ActiveDirectory

# Initialize script variables
$company = 'CBIZ Technology Advisory'
$department = 'FS Consulting - Business Technology & Risk Solutions'
$description = 'Shared Mailbox'
$manager = 'CN=Thompson\, Paul,OU=Users,OU=PHI6,OU=East,DC=ad,DC=cbiz,DC=com'
$sourceFile = 'C:\scripts\CDIshared.csv'
$targetDomainController = 'CLE1VDC032.ad.cbiz.com'
$targetOU = 'OU=Service Accounts,OU=PHI6,OU=East,DC=ad,DC=cbiz,DC=com'

# Read in the data file
if ( Test-Path -Path $sourceFile ) {
    $csvData = Import-Csv -Path $sourceFile

    # Only continue if data is available
    if ( $null -ne $csvData ) {
        foreach ( $dataLine in $csvData ) {
            $cn = "{0}" -f $dataLine.DisplayName
            $dn = "CN={0},{1}" -f $cn, $targetOU
            $userPrincipalName = $dataLine.UserPrincipalName.ToLower()

            if ( $null -eq ( Get-ADObject -LDAPFilter ( "(distinguishedName={0})" -f $dn ) -Server $targetDomainController ) ) {
                $currentIncrement = 0

                do {
                    $sAMAccountName = ( $dataLine.UserPrincipalName.ToLower().Split( '@' ) )[ 0 ]

                    if ( $null -eq ( Get-ADObject -LDAPFilter ( "(sAMAccountName={0})" -f $sAMAccountName ) -Server $targetDomainController ) ) {
                        $targetsAMAccountName = $sAMAccountName
                    } else {
                        $targetsAMAccountName = $null
                        $currentIncrement++
                    }

                } until ( $null -ne $targetsAMAccountName )

                if ( $null -eq ( Get-ADObject -LDAPFilter ( "(userPrincipalName={0})" -f $userPrincipalName ) -Server $targetDomainController ) ) {
                    $newUserParams = @{
                        'Name'              = $dataLine.DisplayName
                        'Server'            = $targetDomainController
                        'SamAccountName'    = $targetsAMAccountName
                        'UserPrincipalName' = $userPrincipalName
                        'Path'              = $targetOU
                        'GivenName'         = $dataLine.DisplayName
                        'DisplayName'       = $dataLine.DisplayName
                        'Company'           = $company
                        'Department'        = $department
                        'Description'       = $description
                        'OtherAttributes'   = @{
                            'employeeType' = 's'
                            'info'         = 'Created as part of a batch process'
                            'manager'      = $manager
                        }
                    }
                    
                    New-ADUser @newUserParams

                }

            }
        }
        
    }
} else {
    Write-Warning -Message 'Please provide a valid file path for source information!'
}
