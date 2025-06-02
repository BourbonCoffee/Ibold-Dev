#requires -Module ShareGate

#region Parameters
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $false,
        HelpMessage = "Specify the distribution group domain name to be used in the output.")]
    [string]$SourceTenantName = "constellation457",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the filter to be used for the distribution groups to collect. It will default to ''*'' if not specified.")]
    [string]$TargetTenantName = "mmicnc",

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the source credentials XML file to be used. It will prompt for credentials, if not specified.")]
    [string]$SourceCredentialsPath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the target credentials XML file to be used. It will prompt for credentials, if not specified.")]
    [string]$TargetCredentialsPath,

    [Parameter(Mandatory = $true, ValueFromPipeline = $false,
        HelpMessage = "Specify the input CSV file to use used")]
    [string]$InputFilePath
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

$dstFolder = "C:\Users\S.Office.Migrate\Desktop\Development\Testing"
$dstFolderName = "Testing"
$dstCreateFolder = $false

Set-Variable srcSite, dstSite, srcList, dstList, userName, srcUrl, dstUrl
$copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
$propertyTemplate = New-PropertyTemplate -AuthorsAndTimestamps -Permissions -VersionLimit 1 -From "4/1/2024"


Write-Host "Connecting to source tenant, $SourceTenantName..."
$sourceTenant = Connect-Site -Url https://$SourceTenantName-admin.sharepoint.com -Credential $srcCredentials

Write-Host "Connecting to target tenant, $TargetTenantName..."
$targetTenant = Connect-Site -Url https://$TargetTenantName-admin.sharepoint.com -Credential $dstCredentials

if ($sourceTenant -and $targetTenant) {
    if (($table | measure).Count -ge 1) {
        Write-Host "Processing $(($table | Measure).Count) users"

        foreach ($row in $table) {
            Clear-Variable srcSite
            Clear-Variable dstSite
            Clear-Variable srcList
            Clear-Variable dstList
            Clear-Variable userName
            Clear-Variable srcUrl 
            Clear-Variable dstUrl

            try {
                $startTime = (Get-Date).ToUniversalTime()
                Write-Host "Starting time           : $startTime " -ForegroundColor Green

                $userName = $row.SourceEmail
                $srcUrl = $row.SourceUrl
                $dstUrl = $row.DestinationUrl
                $taskName = "Copy OneDrive for $userName"

                Write-Host "Migrating OneDrive for  : $userName"
                Write-Host "Source URL              : $srcUrl"
                Write-Host "Destination URL         : $dstUrl"

                Add-SiteCollectionAdministrator -CentralAdmin $sourceTenant.Site -SiteUrl $row.SourceUrl -ErrorAction Stop | Out-Null
                Add-SiteCollectionAdministrator -CentralAdmin $targetTenant.Site -SiteUrl $row.DestinationUrl -ErrorAction Stop | Out-Null
                
                #Write-Host "Connecting to source site:" $srcUrl
                $srcSite = Connect-Site -Url $srcUrl -UseCredentialsFrom $sourceTenant -ErrorAction Stop
                #Write-Host "Connecting to destination site:" $dstUrl
                $dstSite = Connect-Site -Url $dstUrl -UseCredentialsFrom $targetTenant -ErrorAction Stop
                
                #Write-Host "Getting source list"
                $srcList = Get-List -Site $srcSite -Name "Documents" -ErrorAction Stop
                #Write-Host "Getting destination list"
                $dstList = Get-List -Site $dstSite -Name "Documents" -ErrorAction Stop

                Write-Host "Copying content from    : $srcSite"

                if (-not [string]::IsNullOrEmpty($dstFolder) -and $dstCreateFolder) {
                    Write-Host "Creating target folder  : $dstFolderName"
                    #Import-Document -SourceFilePath $dstFolder -DestinationList $dstList -ErrorAction SilentlyContinue
                } else {
                    #Copy-Content -SourceList $srcList -DestinationList $dstList -CopySettings $copysettings -TaskName $taskName -Template $propertyTemplate -DestinationFolder $dstFolderName -ErrorAction Stop
                    Copy-Content -SourceList $srcList -DestinationList $dstList -CopySettings $copysettings -TaskName $taskName -Template $propertyTemplate -ErrorAction Stop -WaitForImportCompletion:$false
                }#if/else
                #Remove-SiteCollectionAdministrator -CentralAdmin $srcTenant.Site -SiteUrl $row.SourceUrl
                #Remove-SiteCollectionAdministrator -CentralAdmin $dstTenant.Site -SiteUrl $row.DestinationUrl
                
                $EndTIme = (Get-Date).ToUniversalTime()
                $RunTime = ($EndTime - $startTime)
                $RunTime = '{0:00}:{1:00}:{2:00}:{3:00}.{4:00}' -f $RunTime.Days, $RunTime.Hours, $RunTime.Minutes, $RunTime.Seconds, $RunTime.Milliseconds
                Write-Host "End time was            : $EndTime"
                Write-Host "Run time was            : $RunTime"
                Write-Host "`r"
                Write-Host "`r"
                
                # Only do this once for testing
                #return
            } catch {
                Write-Error "Error migrating $username"
                Write-Host $_
            }#try/catch
        }#foreach
    } else {
        Write-Warning "No users found"
    }#if/else
}#if
