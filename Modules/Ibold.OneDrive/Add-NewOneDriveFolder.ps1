Function Add-NewOneDriveFolder {
    #requires -Module ShareGate
    <#
    .SYNOPSIS
        This function will utilize the ShareGate PowerShell module to create a folder in a user's OneDrive.

    .DESCRIPTION
        This function will utilize the ShareGate PowerShell module to create a folder of your choosing at a destination of your choosing in a user's OneDrive. The ShareGate module needs
        a folder on the host to read from. The folder name you specify will be created at C:\temp\ if it does not already exist
        This function is currently built for bulk use using an input file. A future build will allow singular use.
        
    .PARAMETER TargetTenantName
        The target tenant name from either: https://<TargetTenantName>-admin.sharepoint.com or https://<TargetTenantName>.sharepoint.com
        
    .PARAMETER TargetCredentialsPath
        The path to the XML file where the target credentials are stored.

    .PARAMETER InputFilePath
        The path to a CSV file with at least the following, two headers:
            DestinationEmail - The primary email address of the user at the target
            DestinationUrl - The URL to the user's personal OneDrive

    .PARAMETER FolderToBeCreated
        The name of the folder to be created.

    .PARAMETER TargetList
        The list where the folder will be created. Defaults to "Documents" if not specified. "Documents" is the root of the user's OneDrive.

    .EXAMPLE
        Add-NewOneDriveFolder -TargetTenantName "Sterling" -TargetCredentialsPath "C:\temp\creds.xml" -InputFilePath "C:\temp\users.csv" -FolderToBeCreated "Migrated" -TargetList "Documents" -Verbose
        This example will add a folder called "Migrated" to the root of each user's OneDrive listed in the input file in the Sterling tenant.

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

        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the name of the folder to be created")]
        [string]$FolderToBeCreated,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false,
            HelpMessage = "Specify the list to get. Defaults to `"Documents`" is not specified")]
        [string]$TargetList = "Documents"
    )
    #endregion

    if ($TargetCredentialsPath) {
        $targetCredentials = Import-Clixml $TargetCredentialsPath
    }

    if ($FolderToBeCreated) {
        try {
            Write-Verbose "Testing if path of folder to import exists..." #| Out-Log
            $pathOfFolder = "C:\temp\$FolderToBeCreated"
            if (-not (Test-Path -Path $pathOfFolder)) {
                Write-Verbose "Folder does not exist. Creating folder $FolderToBeCreated at $pathOfFolder" #| Out-Log
                New-Item -Path $pathOfFolder -Type Directory
            } else {
                Write-Verbose "Folder already exists at $pathOfFolder. Continuing..." #| Out-Log
            }
        } catch {
            Write-Error "Error creating folder $FolderToBeCreated at: C:\temp\" #| Out-Log
            Write-Error "Please check local permissions." #| Out-Log
            Write-Host $_ #| Out-Log
        }
    }

    if (Test-Path -Path $InputFilePath) {
        $table = Import-Csv $inputFilePath -Delimiter ","
    }

    Set-Variable destinationList, userName, destinationSite, destinationURL
    Write-Host "Connecting to target tenant, $($TargetTenantName)..." #| Out-Log
    $targetTenant = Connect-Site -Url https://$TargetTenantName-admin.sharepoint.com -Credential $targetCredentials

    if ($targetTenant) {
        if (($table | Measure-Object).Count -ge 1) {
            Write-Host "Processing $(($table | Measure-Object).Count) users" -ForegroundColor Green #| Out-Log

            foreach ($row in $table) {
                Clear-Variable destinationSite
                Clear-Variable destinationList
                Clear-Variable userName
                Clear-Variable destinationURL

                try {
                    $startTime = (Get-Date).ToUniversalTime()
                    Write-Host "Starting time: $($startTime)" -ForegroundColor Green
                    Write-Verbose "Starting time: $($startTime)" #| Out-Log
                    Write-Verbose "`r" #| Out-Log

                    $userName = $row.DestinationEmail
                    $destinationURL = $row.DestinationUrl
                    $taskName = "Create Folder for $($userName)"

                    Write-Verbose "Preparing to create folder for: $($userName)" #| Out-Log
                    Write-Verbose "Destination URL: $($destinationURL)" #| Out-Log
                    Write-Verbose "`r" #| Out-Log

                    Write-Verbose "Adding $($targetCredentials.UserName) as Site Collection Admin to $userName's OneDrive" #| Out-Log
                    Add-SiteCollectionAdministrator -CentralAdmin $targetTenant.Site -SiteUrl $row.DestinationUrl -ErrorAction Stop | Out-Null
                    
                    Write-Host "Connecting to destination site: $($destinationURL)" #| Out-Log
                    $destinationSite = Connect-Site -Url $destinationURL -UseCredentialsFrom $targetTenant -ErrorAction Stop
                    
                    Write-Verbose "Getting destination list: $($TargetList)" #| Out-Log
                    $destinationList = Get-List -Site $destinationSite -Name $TargetList -ErrorAction Stop
                    
                    Write-Host "Creating target folder: $($FolderToBeCreated)" #| Out-Log
                    Import-Document -SourceFilePath $pathOfFolder -DestinationList $destinationList -TaskName $taskName -ErrorAction SilentlyContinue
                    
                    Write-Verbose "$($FolderToBeCreated) folder created in the OneDrive root directory" #| Out-Log
                    $EndTime = (Get-Date).ToUniversalTime()
                    $RunTime = ($EndTime - $startTime)
                    $RunTime = '{0:00}:{1:00}:{2:00}:{3:00}.{4:00}' -f $RunTime.Days, $RunTime.Hours, $RunTime.Minutes, $RunTime.Seconds, $RunTime.Milliseconds
                    Write-Host "End time was            : $EndTime" -ForegroundColor Green #| Out-Log
                    Write-Host "Run time was            : $RunTime`r`n" -ForegroundColor Green #| Out-Log
                    
                    # Only do this once for testing
                    #return
                } catch {
                    Write-Error "Error creating folder for: $userName" #| Out-Log
                    Write-Host $_ #| Out-Log
                }#try/catch
            }#foreach
        } else {
            Write-Warning "No users found" #| Out-Log
        }#if/else
    }#if
}