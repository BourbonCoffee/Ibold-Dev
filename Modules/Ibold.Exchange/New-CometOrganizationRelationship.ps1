function New-CometOrganizationRelationship {
    <#
    .SYNOPSIS
        Creates one or more one-way Organization Relationships in Exchange Online to allow Comet to view Free/Busy information.

    .DESCRIPTION
        This function is intended to be run on the client tenant side. It sets up an Organization Relationship that allows Comet to view the client's Free/Busy availability.
        It supports either a single domain setup or bulk setup via a CSV file. It can also explicitly pass federated info for hybrid/on-prem environments.

    .PARAMETER DomainName
        (Optional) The external domain to set up the organization relationship for. Defaults to "cometcg.com"

    .PARAMETER RelationshipName
        (Optional) The name to assign to the Organization Relationship. Defaults to "Comet MSP".

    .PARAMETER CsvPath
        (Optional) The path to a CSV file containing DomainName and RelationshipName columns for batch creation.

    .PARAMETER AccessLevel
        (Optional) Level of Free/Busy information shared. Defaults to 'AvailabilityOnly'. Options are 'AvailabilityOnly' or 'LimitedDetails'.

    .PARAMETER UPN
        (Optional) The UPN used to connect to Exchange Online. If omitted, will prompt or use existing session.

    .PARAMETER Force
        (Optional) If specified, will overwrite existing Organization Relationships with the same DomainName.

    .PARAMETER OnPrem
        (Optional) If specified, will gather federated info (Application URI and Autodiscover Endpoint) for hybrid or on-premises deployments.

    .EXAMPLE
        New-CometOrganizationRelationship

    .EXAMPLE
        New-CometOrganizationRelationship -CsvPath "C:\Input\Domains.csv" -Force -OnPrem

    .NOTES
        Author: Chris Ibold
        Comet Consulting Group
        Date: April 11, 2025
        Version: 1.0

        Requires ExchangeOnlineManagement module.
    #>

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $false)]
        [string]$DomainName = "cometcg.com",

        [Parameter(Mandatory = $false)]
        [string]$RelationshipName = "Comet MSP",

        [Parameter(Mandatory = $false)]
        [string]$CsvPath,

        [Parameter(Mandatory = $false)]
        [ValidateSet('AvailabilityOnly', 'LimitedDetails')]
        [string]$AccessLevel = "AvailabilityOnly",

        [Parameter(Mandatory = $false)]
        [string]$UPN,

        [Parameter(Mandatory = $false)]
        [switch]$Force,

        [Parameter(Mandatory = $false)]
        [switch]$OnPrem
    )

    try {
        # Load Exchange Module
        if (-not (Get-Module ExchangeOnlineManagement)) {
            Write-Verbose "Importing ExchangeOnlineManagement module..."
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
        }

        # Connect to Exchange Online
        if (-not (Get-ConnectionInformation)) {
            if ($UPN) {
                Write-Host "Connecting to Exchange Online as $UPN..." -ForegroundColor Cyan
                Connect-ExchangeOnline -UserPrincipalName $UPN -ShowProgress $true -ErrorAction Stop
            } else {
                Write-Host "Connecting to Exchange Online with default credentials..." -ForegroundColor Cyan
                Connect-ExchangeOnline -ShowProgress $true -ErrorAction Stop
            }
        } else {
            Write-Verbose "Already connected to Exchange Online."
        }

        # Build input list
        if ($CsvPath) {
            Write-Host "Importing CSV from $CsvPath..." -ForegroundColor Cyan

            $rawLines = Get-Content -Path $CsvPath | Where-Object { $_ -notmatch '^\s*#' -and $_.Trim() -ne "" }
            $inputList = $rawLines | ConvertFrom-Csv

            # Validate CSV structure
            if (-not ($inputList[0].PSObject.Properties.Name -contains "DomainName" -and
                    $inputList[0].PSObject.Properties.Name -contains "RelationshipName")) {
                throw "CSV file must contain headers 'DomainName' and 'RelationshipName'."
            }
        } else {
            $inputList = @([PSCustomObject]@{
                    DomainName       = $DomainName
                    RelationshipName = $RelationshipName
                })
        }

        foreach ($item in $inputList) {
            $domain = $item.DomainName
            $relationshipName = $item.RelationshipName

            try {
                Write-Host "`nProcessing domain: $domain..." -ForegroundColor Yellow

                $existingRelationship = Get-OrganizationRelationship | Where-Object { $_.DomainNames -contains $domain }
                
                if ($existingRelationship) {
                    Write-Host "⚠️  An organization relationship for domain '$domain' already exists: '$($existingRelationship.Name)'" -ForegroundColor Yellow

                    $confirmDelete = Read-Host "❓ Do you want to REMOVE and RECREATE the Organization Relationship? (Y/N)"

                    if ($confirmDelete.Trim().ToUpper() -eq 'Y') {
                        Write-Host "Removing existing Organization Relationship '$($existingRelationship.Name)'..." -ForegroundColor Red
                        Remove-OrganizationRelationship -Identity $existingRelationship.Identity -Confirm:$false -ErrorAction Stop
                    } else {
                        Write-Host "Skipping domain '$domain'. No changes made." -ForegroundColor DarkYellow
                        continue
                    }
                }


                if ($OnPrem) {
                    Write-Host "OnPrem switch activated, gathering federated information. Please wait as this may take a moment..." -ForegroundColor Magenta
                    $federationInfo = Get-FederationInformation -DomainName $domain -BypassAdditionalDomainValidation -ErrorAction Stop
                }

                if ($PSCmdlet.ShouldProcess("Domain '$domain'", "Create Organization Relationship '$relationshipName'")) {
                    if ($OnPrem -and $federationInfo) {
                        $federationInfo | New-OrganizationRelationship `
                            -Name $relationshipName `
                            -FreeBusyAccessEnabled $true `
                            -FreeBusyAccessLevel $AccessLevel `
                            -Enabled $true `
                            -TargetAutodiscoverEpr $federationInfo.TargetAutodiscoverEpr `
                            -Verbose
                    } else {
                        Write-Host "Gathering federated information. Please wait as this may take a moment..." -ForegroundColor Magenta
                        Get-FederationInformation -DomainName $domain -BypassAdditionalDomainValidation -Verbose -ErrorAction Stop | New-OrganizationRelationship `
                            -Name $relationshipName `
                            -FreeBusyAccessEnabled $true `
                            -FreeBusyAccessLevel $AccessLevel `
                            -Enabled $true `
                            -Verbose
                    }

                    Write-Host "✅ Successfully created Organization Relationship: '$relationshipName' for domain '$domain'." -ForegroundColor Green
                }
            } catch {
                Write-Host "❌ Failed to process domain '$domain'. Error: $_" -ForegroundColor Red
            }
        }

        Write-Host "`nAll requested Organization Relationships have been processed." -ForegroundColor Cyan
    } catch {
        Write-Error "Global function error encountered: $_"
    }
}