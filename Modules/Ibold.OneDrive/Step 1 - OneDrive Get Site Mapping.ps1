#requires -Module ShareGate

#region Parameters
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $false,
        HelpMessage = "Specify the distribution group domain name to be used in the output.")]
    [string]$SourceTenantName = "xcentricERICKS",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the filter to be used for the distribution groups to collect. It will default to ''*'' if not specified.")]
    [string]$TargetTenantName = "cbizcorp",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the source credentials XML file to be used. It will prompt for credentials, if not specified.")]
    [string]$SourceCredentialsPath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the target credentials XML file to be used. It will prompt for credentials, if not specified.")]
    [string]$TargetCredentialsPath,

    [Parameter(Mandatory = $true, ValueFromPipeline = $false,
        HelpMessage = "Specify the input CSV file to use used")]
    [string]$InputFilePath,

    [Parameter(Mandatory = $true, ValueFromPipeline = $false,
        HelpMessage = "Specify the output CSV file to be used")]
    [string]$OutputFilePath
)
#endregion

if ($SourceCredentialsPath) {
    $srcCredentials = Import-CLIXML $SourceCredentialsPath
}

if ($TargetCredentialsPath) {
    $dstCredentials = Import-Clixml $TargetCredentialsPath
}

if (Test-Path -Path $InputFilePath) {
    $table = Import-Csv $inputFilePath -Delimiter ","
}


Write-Host "Connecting to source tenant, $SourceTenantName..."
#$srcTenant = Connect-Site -Url https://$srcTenantAddress-admin.sharepoint.com -Browser -DisableSSO
$sourceTenant = Connect-Site -Url https://$SourceTenantName-admin.sharepoint.com -Credential $srcCredentials

Write-Host "Connecting to target tenant, $TargetTenantName..."
#$dstTenant = Connect-Site -Url https://$dstTenantAddress-admin.sharepoint.com -Browser -DisableSSO
$targetTenant = Connect-Site -Url https://$TargetTenantName-admin.sharepoint.com -Credential $dstCredentials

if ($sourceTenant -and $targetTenant) {
    $first = $true
    
    foreach ($row in $table) {
        try {

            $sourceEmail = $row.SourcePrimarySmtpAddress 
            $sourceUrl = Get-OneDriveUrl -Tenant $sourceTenant -Email $sourceEmail

            $targetEmail = $row.TargetUserPrincipalName 
            $targetUrl = Get-OneDriveUrl -Tenant $targetTenant -Email $targetEmail -ProvisionIfRequired

            if ($sourceUrl) {
                $item = [PSCustomObject]@{Username = ''; SourceEmail = $sourceEmail; SourceUrl = $sourceUrl; DestinationEmail = $targetEmail; DestinationUrl = $targetUrl }

                if ($first) {
                    $item | Export-Csv -Path $OutputFilePath -NoTypeInformation -Delimiter "," 
                    $first = $false
                }
                else {
                    $item | Export-Csv -Path $OutputFilePath -NoTypeInformation -Delimiter "," -Append
                }
            
                Write-Output $item
            } else {
                Write-Warning "No source URL found for $($sourceEmail)"
            }
        }
        catch {
            Write-Output "ERROR: $($_.Exception.Message)"
        }
    }
} else {
    Write-Error -Message "Not connected to source and/or target"
}
