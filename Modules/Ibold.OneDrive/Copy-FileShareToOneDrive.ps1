Function Copy-FileShareToOneDrive {
    #requires -Module ShareGate
    <#
    .SYNOPSIS
        This function will utilize the ShareGate PowerShell module to copy a file share to OneDrive.

    .DESCRIPTION
        This function will utilize the ShareGate PowerShell module to copy a file share to OneDrive using either a mapped share or a UNC path.
        This function is currently built for bulk use using an input file. A future build will allow singular use.

    .PARAMETER TargetTenantName
        The target tenant name from either: https://<TargetTenantName>-admin.sharepoint.com or https://<TargetTenantName>.sharepoint.com
        
    .PARAMETER TargetCredentialsPath
        The path to the XML file where the target credentials are stored.

    .PARAMETER InputFilePath
        The path to a CSV file with at least the following, two headers:
            DestinationEmail - The primary email address of the user at the target
            DestinationUrl - The URL to the user's personal OneDrive
            SourcePath - The UNC path or mapped drive path to the source files.

    .PARAMETER TargetList
        The list where the folder will be created. Defaults to "Documents" if not specified. "Documents" is the root of the user's OneDrive.

    .PARAMETER DestinationFolder
        The target folder to migrate the file share content into.

    .PARAMETER VersionLimit
        The number of versions for Office docs to migrate. Defaults to 1 if not specified.

    .EXAMPLE
        Copy-FileShareToOneDrive -TargetTenantName "Sterling" -TargetCredentialsPath "C:\temp\creds.xml" -InputFilePath "C:\temp\users.csv" -DestinationFolder "Migrated" -TargetList "Documents" -VersionLimit 1 -Verbose
        This will copy only the single, main version of files from the file share to the user's OneDrive as specified in the input file into a folder called Migrated in the Documents list.
    
    .NOTES

    #>

    #region Parameters
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the destination tenant name")]
        [string]$TargetTenantName,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the path to the target credentials XML file to be used. It will prompt for credentials, if not specified")]
        [string]$TargetCredentialsPath,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the input CSV file to use used")]
        [string]$InputFilePath,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the list to get. Defaults to `"Documents`" is not specified")]
        [string]$TargetList = "Documents",

        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the name of the folder the content will migrate into")]
        [string]$DestinationFolder,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the number of versions to migrate. Defaults to 1, if not specified")]
        [int]$VersionLimit = 1
    )
    #endregion

    if ($TargetCredentialsPath) {
        $destinationCredentials = Import-Clixml $TargetCredentialsPath
    }

    if (Test-Path -Path $InputFilePath) {
        $table = Import-Csv $inputFilePath -Delimiter ","
    }

    Set-Variable destinationSite, destinationList, destinationURL, sourceUNC, userName
    $copySettings = New-CopySettings -OnContentItemExists IncrementalUpdate
    $propertyTemplate = New-PropertyTemplate -AuthorsAndTimestamps -VersionLimit $VersionLimit

    Write-Host "Connecting to target tenant, $TargetTenantName..." #| Out-Log
    $targetTenant = Connect-Site -Url https://$TargetTenantName-admin.sharepoint.com -Credential $destinationCredentials

    if ($targetTenant) {
        if (($table | Measure-Object).Count -ge 1) {
            Write-Host "Processing $(($table | Measure-Object).Count) users" -ForegroundColor Green #| Out-Log

            foreach ($row in $table) {
                Clear-Variable destinationSite
                Clear-Variable destinationList
                Clear-Variable destinationURL
                Clear-Variable sourceUNC
                Clear-Variable userName

                try {
                    $startTime = (Get-Date).ToUniversalTime()
                    Write-Host "Starting time: $startTime" -ForegroundColor Green #| Out-Log

                    $userName = $row.DestinationEmail
                    $sourceUNC = $row.SourceUNC
                    $destinationURL = $row.DestinationUrl
                    #$taskName = "Copy UNC to OneDrive for $userName"

                    Write-Host "Migrating UNC to OneDrive for   : $userName" #| Out-Log
                    Write-Host "Source UNC Path                 : $sourceUNC" #| Out-Log
                    Write-Host "Destination URL                 : $destinationURL" #| Out-Log

                    Write-Verbose "Adding $($destinationCredentials.UserName) as Site Collection Admin to $userName's OneDrive" #| Out-Log
                    #Add-SiteCollectionAdministrator -CentralAdmin $targetTenant.Site -SiteUrl $row.DestinationUrl -ErrorAction Stop | Out-Null
                    
                    Write-Verbose "Connecting to destination site: $destinationURL" #| Out-Log
                    $destinationSite = Connect-Site -Url $destinationURL -UseCredentialsFrom $targetTenant -ErrorAction Stop
                    
                    Write-Verbose "Getting destination list" #| Out-Log
                    $destinationList = Get-List -Site $destinationSite -Name "Documents" -ErrorAction Stop

                    # Splat Import-Document parameters
                    $importParameters = [ordered]@{
                        SourceFolder            = $row.SourceUNC
                        DestinationList         = $destinationList
                        CopySettings            = $copySettings
                        Template                = $propertyTemplate
                        DestinationFolder       = $DestinationFolder
                        TaskName                = "Copy FileShare to OneDrive for $userName"   
                        WaitForImportCompletion = $false
                    }

                    Write-Host "Copying content from: $sourceUNC" #| Out-Log
                    Import-Document @importParameters

                    $endTime = (Get-Date).ToUniversalTime()
                    $runTime = ($endTime - $startTime)
                    $runTime = '{0:00}:{1:00}:{2:00}:{3:00}.{4:00}' -f $runTime.Days, $runTime.Hours, $runTime.Minutes, $runTime.Seconds, $runTime.Milliseconds
                    Write-Host "End time was: $endTime" -ForegroundColor Green #| Out-Log
                    Write-Host "Run time was: $runTime`r`n" -ForegroundColor Green #| Out-Log
                    
                    # Only do this once for testing
                    #return
                } catch {
                    Write-Error "Error migrating $userName" #| Out-Log
                    Write-Host $_ #| Out-Log
                }#try/catch
            }#foreach
        } else {
            Write-Warning "No users found" #| Out-Log
        }#if/else
    }#if
}