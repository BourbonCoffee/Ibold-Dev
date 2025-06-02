# Import required modules
Import-Module ActiveDirectory

# Initialize script variables
$company = 'CBIZ Technology Advisory'
$department = 'FS Consulting - Business Technology & Risk Solutions'
$description = 'Shared Mailbox'
$manager = 'CN=Chris Ibold,OU=Users,OU=Bourbon,DC=ibold,DC=dev'
$sourceFile = 'C:\Users\vvvibold\Desktop\Mailbox-Statistics.csv'
$targetDomainController = 'Coffee-DC-01.ibold.dev'
$targetOU = 'OU=Shared Mailboxes,OU=Bourbon,DC=ibold,DC=dev'
$date = Get-Date -Format MM/dd/yy

# Read in the data file
if ( Test-Path -Path $sourceFile ) {
    $csvData = Import-Csv -Path $sourceFile

    # Only continue if data is available
    if ( $null -ne $csvData ) {
        foreach ( $columnHeader in $csvData ) {
            $cn = "{0}" -f $columnHeader.DisplayName
            $dn = "CN={0},{1}" -f $cn, $targetOU
            $userPrincipalName = $columnHeader.EmailAddress.ToLower()

            if ( $null -eq ( Get-ADObject -LDAPFilter ( "(distinguishedName={0})" -f $dn ) -Server $targetDomainController ) ) {
                $currentIncrement = 0

                do {
                    $sAMAccountName = ( $columnHeader.EmailAddress.ToLower().Split( '@' ) )[ 0 ]

                    if ( $null -eq ( Get-ADObject -LDAPFilter ( "(sAMAccountName={0})" -f $sAMAccountName ) -Server $targetDomainController ) ) {
                        $targetsAMAccountName = $sAMAccountName
                    } else {
                        $targetsAMAccountName = $null
                        $currentIncrement++
                    }

                } until ( $null -ne $targetsAMAccountName )

                if ( $null -eq ( Get-ADObject -LDAPFilter ( "(userPrincipalName={0})" -f $userPrincipalName ) -Server $targetDomainController ) ) {
                    $newUserParams = @{
                        'Name'              = $columnHeader.DisplayName
                        'Server'            = $targetDomainController
                        'SamAccountName'    = $targetsAMAccountName
                        'UserPrincipalName' = $userPrincipalName
                        'Path'              = $targetOU
                        'GivenName'         = $columnHeader.DisplayName
                        'DisplayName'       = $columnHeader.DisplayName
                        'Company'           = $company
                        'Department'        = $department
                        'Description'       = $description
                        'OtherAttributes'   = @{
                            'employeeType' = 's'
                            'info'         = "Created on $date by Chris Ibold as part of a batch process"
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
