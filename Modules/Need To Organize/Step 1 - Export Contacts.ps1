#Requires -version 5.0
#Requires -Modules Sterling

#region Parameters
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $false,
        HelpMessage = "Specify the contact domain name to be used in the output.")]
    [string]$ContactDomain,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the filter to be used for the contact to collect. It will default to ''*'' if not specified.")]
    [string]$ContactFilter,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
        HelpMessage = "Specify the path to the credential XML file to be used")]
    [string]$CredentialsPath,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false,
    HelpMessage = "Specify the path to the export location to be used")]
    [string]$ExportLocation
)
#endregion

#region User Variables
#Very little reason to change these
$InformationPreference = "Continue"

if ($DebugPreference -eq "Confirm" -or $DebugPreference -eq "Inquire") {$DebugPreference = "Continue"}
#endregion

#region Static Variables
#Don't change these
Set-Variable -Name strBaseLocation -Option AllScope -Scope Script
Set-Variable -Name dateStartTimeStamp -Option AllScope -Scope Script -Value (Get-Date).ToUniversalTime()
Set-Variable -Name strLogTimeStamp -Option AllScope -Scope Script -Value $dateStartTimeStamp.ToString("MMddyyyy_HHmmss")
Set-Variable -Name strLogFile -Option ReadOnly -Scope Script
Set-Variable -Name htLoggingPreference -Option AllScope -Scope Script -Value @{"InformationPreference"=$InformationPreference; `
    "WarningPreference"=$WarningPreference;"ErrorActionPreference"=$ErrorActionPreference;"VerbosePreference"=$VerbosePreference;"DebugPreference"=$DebugPreference}
Set-Variable -Name verScript -Option AllScope -Scope Script -Value "5.1.2023.0817"

Set-Variable -Name StarterContactFilter -Option AllScope -Scope Script -Value "((RecipientType -eq 'MailContact') -or (RecipientType -eq 'MailUser'))"

Set-Variable -Name boolScriptIsModulesLoaded -Option AllScope -Scope Script -Value $false
Set-Variable -Name ExitCode -Option AllScope -Scope Script -Value 1

New-Object System.Data.DataTable | Set-Variable dtContacts -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtEmailAddresses -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute1 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute2 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute3 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute4 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtExtensionCustomAttribute5 -Option AllScope -Scope Script
New-Object System.Data.DataTable | Set-Variable dtManagedBy -Option AllScope -Scope Script

Set-Variable -Name arrContactAttribs -Option AllScope -Scope Script -Value 'Guid','Identity','Alias','ArchiveGuid', `
    'AuthenticationType','City','Notes','Company','CountryOrRegion','PostalCode','CustomAttribute1','CustomAttribute2', `
    'CustomAttribute3','CustomAttribute4','CustomAttribute5','CustomAttribute6','CustomAttribute7','CustomAttribute8', `
    'CustomAttribute9','CustomAttribute10','CustomAttribute11','CustomAttribute12','CustomAttribute13','CustomAttribute14', `
    'CustomAttribute15','ExtensionCustomAttribute1','ExtensionCustomAttribute2','ExtensionCustomAttribute3','ExtensionCustomAttribute4', `
    'ExtensionCustomAttribute5','Database','ArchiveDatabase','DatabaseName','Department','ExternalDirectoryObjectId', `
    'ManagedFolderMailboxPolicy','EmailAddresses','ExpansionServer','ExternalEmailAddress','DisplayName','FirstName', `
    'HiddenFromAddressListsEnabled','EmailAddressPolicyEnabled','IsDirSynced','LastName','ResourceType','ManagedBy', `
    'Manager','Name','Office','OrganizationalUnit','Phone','PrimarySmtpAddress','RecipientType','RecipientTypeDetails', `
    'SamAccountName','ServerLegacyDN','ServerName','StateOrProvince','StorageGroupName','Title','WindowsLiveID', `
    'OwaMailboxPolicy','AddressBookPolicy','InformationBarrierSegments','WhenIBSegmentChanged','SharingPolicy', `
    'RetentionPolicy','ShouldUseDefaultRetentionPolicy','ArchiveRelease','IsValidSecurityPrincipal','LitigationHoldEnabled', `
    'ArchiveState','SKUAssigned','WhenMailboxCreated','UsageLocation','ExchangeGuid','ArchiveStatus','WhenSoftDeleted', `
    'UnifiedGroupSKU','DistinguishedName','WhenChangedUTC','WhenCreatedUTC', 'legacyExchangeDN'
#endregion

#region Complete Functions
Function _ConfirmScriptRequirements
{
    <#
    .SYNOPSIS
        Verifies that all necessary requirements are present for the script and return true/false
    .EXAMPLE
        $valid = _ConfirmScriptRequirements

        This would check the script requirements and set $valid to true/false based on the results
    .NOTES
        Version:
        - 5.1.2023.0808:    New function
    #>
    [CmdletBinding()]
    Param()

    begin {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        Write-Debug -Message "Starting _ConfirmScriptRequirements"
        try {
            Write-Host "Loading Sterling PowerShell module`r"

            if (Get-Module -ListAvailable Sterling -Verbose:$false) {
                Import-Module Sterling -ErrorAction Stop -Verbose:$false
                $script:boolScriptIsModulesLoaded = $true
            } else {
                Write-Warning "Missing Sterling PowerShell module`r"
                $script:boolScriptIsModulesLoaded = $false
            }#if/else
        } catch {
            Write-Error "Unable to load Sterling PowerShell module`r"
            Write-Error $_

            $script:boolScriptIsModulesLoaded = $false
        }#try/catch

        try {
            Write-Host "Loading ExchangeOnlineManagement PowerShell module`r"

            if (Get-Module -ListAvailable ExchangeOnlineManagement -Verbose:$false) {
                $global:VerbosePreference = "SilentlyContinue"

                Import-Module ExchangeOnlineManagement -ErrorAction Stop -Verbose:$false
                $script:boolScriptIsModulesLoaded = $true

                if($htLoggingPreference['VerbosePreference'] -eq "Continue"){$global:VerbosePreference = "Continue"}#if
            } else {
                Write-Warning "Missing ExchangeOnlineManagement PowerShell module`r"
                $script:boolScriptIsModulesLoaded = $false
            }#if/else
        } catch {
            Write-Error "Unable to load ExchangeOnlineManagement PowerShell module`r"
            Write-Error $_

            $script:boolScriptIsModulesLoaded = $false
        }#try/catch

        Set-Variable -Name strBaseLocation -Option AllScope -Scope Script -Value $(_GetScriptDirectory -Path)
        Set-Variable -Name strLogFile -Option ReadOnly -Force -Scope Script -Value "$script:strBaseLocation\Logging\$script:strLogTimeStamp-$((_GetScriptDirectory -Leaf).Replace(".ps1",'')).log"

        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Debug -WriteBackToHost -Message "Starting _ConfirmScriptRequirements"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Script version $verScript starting"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: InformationPreference = $InformationPreference"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ErrorActionPreference = $ErrorActionPreference"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: VerbosePreference = $VerbosePreference"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: DebugPreference = $DebugPreference"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ContactDomain = $ContactDomain"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ContactFilter = $ContactFilter"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: CredentialsPath = $CredentialsPath"
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Parameter: ExportLocation = $ExportLocation"
    }#begin
    
    process{
        if ($script:boolScriptIsModulesLoaded) {
            try{
                $global:VerbosePreference = "SilentlyContinue"

                $ConnectSplat = @{
                    "ShowBanner" = $False
                }

                if ($CredentialsPath -and (Test-Path -Path $CredentialsPath)) {
                    $ConnectSplat.Add("Credential", $(Import-Clixml $CredentialsPath))
                } else {
                    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "CredentialsPath specified but does not exist"
                    $script:boolScriptIsFilesExist = $false
                }#if/else

                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Connecting to Exchange Online"
                Connect-ExchangeOnline @ConnectSplat
                
                if($htLoggingPreference['VerbosePreference'] -eq "Continue"){$global:VerbosePreference = "Continue"}#if                
            } catch {
                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error verifying script requirements"
                Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
                return $false
            }#try/catch
        }#if

        #Final check
        if ($script:boolScriptIsModulesLoaded){return $true}
        else {return $false}
    }#process

    end {
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Debug -WriteBackToHost -Message "Finishing _ConfirmScriptRequirements"
    }#end
}#function _ConfirmScriptRequirements

Function _CreateMVATable
{
    <#
    .SYNOPSIS
        Create a blank MVA datatable
    .PARAMETER Attribute
        Specify the MVA attribute build the MVA DataTable.
    .EXAMPLE
        $dtGroups = _CreateMVATable Attribute "ManagedBy"
    
        This would create a blank MVA datatable for "ManagedBy" attribute
    .NOTES
        Version:
            - 5.1.2023.0808:    New function
    #>
 
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true,
        ValueFromPipeline = $false,
        HelpMessage = "Specify the attribute for the MVA datatable")]
        [string]$Attribute
    )
    
    begin {
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _CreateMVATable"
    }#begin
    
    process	{
        try	{
            $dtMVA = New-Object System.Data.DataTable
            $col1 = New-Object System.Data.DataColumn ContactDomain,([string])
            $col2 = New-Object System.Data.DataColumn Guid,([Guid])
            $col3 = New-Object System.Data.DataColumn $Attribute,([string])
            $dtMVA.Columns.Add($col1)
            $dtMVA.Columns.Add($col2)
            $dtMVA.Columns.Add($col3)
            [System.Data.DataColumn[]]$KeyColumn = ($dtMVA.Columns["ContactDomain"],$dtMVA.Columns["Guid"],$dtMVA.Columns[$Attribute])
            $dtMVA.PrimaryKey = $KeyColumn
            return @(,$dtMVA)
        } catch {
            Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to create MVA datatable"
            Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $false
        }#try/catch
    }#process
    
    end {
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _CreateMVATable"
    }#end
}#function _CreateMVATable

function _GetScriptDirectory
{
    <#
    .SYNOPSIS
        _GetScriptDirectory returns the proper location of the script.
 
    .OUTPUTS
        System.String
   
    .NOTES
        Returns the correct path within a packaged executable.
    #>
    [OutputType([string])]
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [switch]$Path,

        [Parameter(Mandatory = $false)]
        [switch]$Leaf,

        [Parameter(Mandatory = $false)]
        [switch]$LeafBase
    )

    if ($null -ne $hostinvocation) {
        if($Leaf) {
            Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf
        } elseif($LeafBase) {
            (Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf).Split(".")[0]
        } elseif($Path) {
            Split-Path $hostinvocation.MyCommand.path
        } else {
            Split-Path $hostinvocation.MyCommand.path
        }#if/else
    } elseif ($null -ne $script:MyInvocation.MyCommand.Path) {
        if($Leaf) {
            Split-Path $script:MyInvocation.MyCommand.Path -Leaf
        } elseif($LeafBase) {
            (Split-Path $script:MyInvocation.MyCommand.Path -Leaf).Split(".")[0]
        } elseif($Path) {
            Split-Path $script:MyInvocation.MyCommand.Path
        } else {
            (Get-Location).Path + "\" + (Split-Path $script:MyInvocation.MyCommand.Path -Leaf)
        }#if/else
    } else {
        if($Leaf) {
            Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf
        } elseif($LeafBase) {
            (Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf).Split(".")[0]
        } elseif($Path) {
            (Get-Location).Path
        } else {
            (Get-Location).Path + "\" + (Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Leaf)
        }#if/else
    }#if/else
}#function _GetScriptDirectory

#endregion

#region Active Development
Function _GetContactInfo 
{
    <#
    .SYNOPSIS
        Collects the necessary contact cache and returns a datatable with the results
    .PARAMETER Filter
        The optional parameter for a filter to be used when querying contacts
   .PARAMETER ContactAttributes
        Specify the array of contacts attributes to return with the DataTable.
    .EXAMPLE
        $dtContacts = _GetContactInfo -RecipientAttributes $RecAttribs
    
        This would get all contacts and return attributes $RecAttribs to $dtContacts
    .NOTES
        Version:
            - 5.1.2023.0808:    New function
    #>
 
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false,
        ValueFromPipeline = $false,
        HelpMessage = "Specify the filter to use for the contact.")]
        [string]$Filter = "*",
        
        [Parameter(Mandatory = $true,
        ValueFromPipeline = $false,
        HelpMessage = "Specify the contact datatable to update with found information")]
        [System.Data.DataTable]$Contacts,
        
        [Parameter(Mandatory = $true,
        ValueFromPipeline = $false,
        HelpMessage = "Specify the array of contact attributes to return with the DataTable.")]
        [array]$ContactAttributes,
        
        [Parameter(Mandatory = $true,
        ValueFromPipeline = $false,
        HelpMessage = "Specify the EmailAddresses datatable to update with found information")]
        [System.Data.DataTable]$EmailAddresses,
        
        [Parameter(Mandatory = $true,
        ValueFromPipeline = $false,
        HelpMessage = "Specify the ExtensionCustomAttribute1 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute1,
        
        [Parameter(Mandatory = $true,
        ValueFromPipeline = $false,
        HelpMessage = "Specify the ExtensionCustomAttribute2 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute2,
        
        [Parameter(Mandatory = $true,
        ValueFromPipeline = $false,
        HelpMessage = "Specify the ExtensionCustomAttribute3 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute3,
        
        [Parameter(Mandatory = $true,
        ValueFromPipeline = $false,
        HelpMessage = "Specify the ExtensionCustomAttribute4 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute4,
        
        [Parameter(Mandatory = $true,
        ValueFromPipeline = $false,
        HelpMessage = "Specify the ExtensionCustomAttribute5 datatable to update with found information")]
        [System.Data.DataTable]$ExtensionCustomAttribute5,

        [Parameter(Mandatory = $true,
        ValueFromPipeline = $false,
        HelpMessage = "Specify the ManagedBy datatable to update with found information")]
        [System.Data.DataTable]$ManagedBy
    )
    
    begin {
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting _GetContactInfo"
    }#begin
    
    process	{
        try	{
            Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Building base datatable"
            $dtContact = Get-MailContact -ResultSize 1 -Verbose:$false -Filter $Filter -WarningAction SilentlyContinue | Select-Object -Property $ContactAttributes | ConvertTo-DataTable
            
            foreach($column in $dtContact.Columns) {if(-not $Contacts.Columns.Contains($column.ColumnName)){[void]$Contacts.Columns.Add($column.ColumnName, $column.DataType)}}
            if(-not $Contacts.Columns.Contains("ContactDomain")){[void]$Contacts.Columns.Add("ContactDomain", "string")}

            Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Populating datatable with contact information"
            Get-MailContact -ResultSize Unlimited -Verbose:$false -Filter $Filter | Select-Object -Property $ContactAttributes | foreach {
                $drNewRow = $Contacts.NewRow()
                ForEach($element in $_.PSObject.Properties) {
                    $columnName = $element.Name
                    $columnValue = $element.Value
                    
                    if ([string]::IsNullorEmpty($columnValue) -or $columnValue.ToString() -eq "Unlimited") {
                        $columnValue = [DBNull]::Value
                    } else {
                        switch ($columnName) {
                            "EmailAddresses" {
                                ForEach($entry in $columnValue){
                                    $drNewAddressRow = $EmailAddresses.NewRow()
                                    $drNewAddressRow["ContactDomain"] = [string]$ContactDomain
                                    $drNewAddressRow["Guid"] = [guid]($drNewRow["Guid"]).Guid
                                    $drNewAddressRow["EmailAddresses"] = [string]$entry
                                    [void]$EmailAddresses.Rows.Add($drNewAddressRow)
                                }#foreach
                            }#EmailAddresses
                            "ExtensionCustomAttribute1" {
                                ForEach($entry in $columnValue){
                                    $drExtCustomAttribute1Row = $ExtensionCustomAttribute1.NewRow()
                                    $drExtCustomAttribute1Row["ContactDomain"] = [string]$ContactDomain
                                    $drExtCustomAttribute1Row["Guid"] = [guid]($drNewRow["Guid"]).Guid
                                    $drExtCustomAttribute1Row["ExtensionCustomAttribute1"] = [string]$entry
                                    [void]$ExtensionCustomAttribute1.Rows.Add($drExtCustomAttribute1Row)
                                }#foreach
                            }#ExtensionCustomAttribute1
                            "ExtensionCustomAttribute2" {
                                ForEach($entry in $columnValue){
                                    $drExtCustomAttribute2Row = $ExtensionCustomAttribute2.NewRow()
                                    $drExtCustomAttribute2Row["ContactDomain"] = [string]$ContactDomain
                                    $drExtCustomAttribute2Row["Guid"] = [guid]($drNewRow["Guid"]).Guid
                                    $drExtCustomAttribute2Row["ExtensionCustomAttribute2"] = [string]$entry
                                    [void]$ExtensionCustomAttribute2.Rows.Add($drExtCustomAttribute2Row)
                                }#foreach
                            }#ExtensionCustomAttribute2
                            "ExtensionCustomAttribute3" {
                                ForEach($entry in $columnValue){
                                    $drExtCustomAttribute3Row = $ExtensionCustomAttribute3.NewRow()
                                    $drExtCustomAttribute3Row["ContactDomain"] = [string]$ContactDomain
                                    $drExtCustomAttribute3Row["Guid"] = [guid]($drNewRow["Guid"]).Guid
                                    $drExtCustomAttribute3Row["ExtensionCustomAttribute3"] = [string]$entry
                                    [void]$ExtensionCustomAttribute3.Rows.Add($drExtCustomAttribute3Row)
                                }#foreach
                            }#ExtensionCustomAttribute3
                            "ExtensionCustomAttribute4" {
                                ForEach($entry in $columnValue){
                                    $drExtCustomAttribute4Row = $ExtensionCustomAttribute4.NewRow()
                                    $drExtCustomAttribute4Row["ContactDomain"] = [string]$ContactDomain
                                    $drExtCustomAttribute4Row["Guid"] = [guid]($drNewRow["Guid"]).Guid
                                    $drExtCustomAttribute4Row["ExtensionCustomAttribute4"] = [string]$entry
                                    [void]$ExtensionCustomAttribute4.Rows.Add($drExtCustomAttribute4Row)
                                }#foreach
                            }#ExtensionCustomAttribute4
                            "ExtensionCustomAttribute5" {
                                ForEach($entry in $columnValue){
                                    $drExtCustomAttribute5Row = $ExtensionCustomAttribute5.NewRow()
                                    $drExtCustomAttribute5Row["ContactDomain"] = [string]$ContactDomain
                                    $drExtCustomAttribute5Row["Guid"] = [guid]($drNewRow["Guid"]).Guid
                                    $drExtCustomAttribute5Row["ExtensionCustomAttribute5"] = [string]$entry
                                    [void]$ExtensionCustomAttribute5.Rows.Add($drExtCustomAttribute5Row)
                                }#foreach
                            }#ExtensionCustomAttribute5
                            "ManagedBy" {
                                ForEach($entry in $columnValue){
                                    $drManagedByRow = $ManagedBy.NewRow()
                                    $drManagedByRow["ContactDomain"] = [string]$ContactDomain
                                    $drManagedByRow["Guid"] = [guid]($drNewRow["Guid"]).Guid
                                    $drManagedByRow["ManagedBy"] = [string]$entry
                                    [void]$ManagedBy.Rows.Add($drManagedByRow)
                                }#foreach
                            }#ManagedBy
                            default {
                                if ($columnValue.gettype().Name -eq "ArrayList") {
                                    $drNewRow["$columnName"] = $columnValue.Clone()
                                } else {
                                    $drNewRow["$columnName"] = $columnValue
                                }#if/else
                            }#default
                        }#switch
                    }#if/else
                }#loop through each property
                $drNewRow["ContactDomain"] = [string]$ContactDomain

                [void]$Contacts.Rows.Add($drNewRow)
            }#get-mailcontact/foreach
            
            return @(,$Contacts)
        } catch {
            Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Error while trying to gather contact information"
            Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message $_
            return $null
        }#try/catch
    }#process
    
    end {
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finishing _GetContactInfo"
    }#end
}#function _GetContactInfo
#endregion

#region Main Program
Write-Host "`r"
Write-Host "Script Written by Sterling Consulting`r"
Write-Host "All rights reserved. Proprietary and Confidential Material`r"
Write-Host "Exchange Contact Inventory Script`r"
Write-Host "`r"

Write-Host "Script starting`r"

if (_ConfirmScriptRequirements) {
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Script requirements met"

    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Creating MVA datatables"
    $dtEmailAddresses = _CreateMVATable -Attribute "EmailAddresses"
    $dtExtensionCustomAttribute1 = _CreateMVATable -Attribute "ExtensionCustomAttribute1"
    $dtExtensionCustomAttribute2 = _CreateMVATable -Attribute "ExtensionCustomAttribute2"
    $dtExtensionCustomAttribute3 = _CreateMVATable -Attribute "ExtensionCustomAttribute3"
    $dtExtensionCustomAttribute4 = _CreateMVATable -Attribute "ExtensionCustomAttribute4"
    $dtExtensionCustomAttribute5 = _CreateMVATable -Attribute "ExtensionCustomAttribute5"
    $dtManagedBy = _CreateMVATable -Attribute "ManagedBy"
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Finished creating MVA datatables"

    #Get recipients from system based on supplied filter
    if ($ContactFilter -ne "") {
        $ContactFilter = $StarterContactFilter + " -and (" + $ContactFilter + ")"
    } else {
        $ContactFilter = $StarterContactFilter
    }#if/else

    $ContactInfoParamsSplat = @{
        "Filter" = $ContactFilter
        "Contacts" = $dtContacts
        "ContactAttributes" = $arrContactAttribs
        "EmailAddresses" = $dtEmailAddresses
        "ExtensionCustomAttribute1" = $dtExtensionCustomAttribute1
        "ExtensionCustomAttribute2" = $dtExtensionCustomAttribute2
        "ExtensionCustomAttribute3" = $dtExtensionCustomAttribute3
        "ExtensionCustomAttribute4" = $dtExtensionCustomAttribute4
        "ExtensionCustomAttribute5" = $dtExtensionCustomAttribute5
        "ManagedBy" = $dtManagedBy
    }

    #Get recipients
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Retrieving contact information"
    _GetContactInfo @ContactInfoParamsSplat | Out-Null

    if ($dtContacts.DefaultView.Count -le 0){
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "No contact information found. Unable to continue without contact information. Exiting script"
        Exit $ExitCode
    }#if
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished retrieving contact information"
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "$($dtContacts.Rows.Count) Contact entries collected"

    if ($ExportLocation -eq "") {
        $ExportLocation = $script:strBaseLocation + "\Exchange"
    }
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exporting CSV to $ExportLocation with , delimiter"
        
    #Check for path/folder
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Checking for $ExportLocation"
    if (-not (Test-Path -Path $ExportLocation)) {
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Creating folder structure for $ExportLocation"
        New-Item -Path $ExportLocation -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    }
    
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Starting export of CSV"
    if($dtContacts.DefaultView.Count -ge 1){
        $dtContacts | Export-Csv -Path "$ExportLocation\ContactInfo_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null
        
        if($dtEmailAddresses.DefaultView.Count -ge 1){$dtEmailAddresses | Export-Csv -Path "$ExportLocation\ContactInfo_EmailAddresses_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null}
        if($dtExtensionCustomAttribute1.DefaultView.Count -ge 1){$dtExtensionCustomAttribute1 | Export-Csv -Path "$ExportLocation\ContactInfo_ExtensionCustomAttribute1_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null}
        if($dtExtensionCustomAttribute2.DefaultView.Count -ge 1){$dtExtensionCustomAttribute2 | Export-Csv -Path "$ExportLocation\ContactInfo_ExtensionCustomAttribute2_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null}
        if($dtExtensionCustomAttribute3.DefaultView.Count -ge 1){$dtExtensionCustomAttribute3 | Export-Csv -Path "$ExportLocation\ContactInfo_ExtensionCustomAttribute3_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null}
        if($dtExtensionCustomAttribute4.DefaultView.Count -ge 1){$dtExtensionCustomAttribute4 | Export-Csv -Path "$ExportLocation\ContactInfo_ExtensionCustomAttribute4_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null}
        if($dtExtensionCustomAttribute5.DefaultView.Count -ge 1){$dtExtensionCustomAttribute5 | Export-Csv -Path "$ExportLocation\ContactInfo_ExtensionCustomAttribute5_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null}
        if($dtManagedBy.DefaultView.Count -ge 1){$dtManagedBy | Export-Csv -Path "$ExportLocation\ContactInfo_ManagedBy_$strLogTimeStamp.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop | Out-Null}
        
        $ExitCode = 0
    }
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Verbose -WriteBackToHost -Message "Finished export of CSV"

    $RunTime = ((get-date).ToUniversalTime() - $dateStartTimeStamp)
    $RunTime = '{0:00}:{1:00}:{2:00}:{3:00}.{4:00}' -f $RunTime.Days,$RunTime.Hours,$RunTime.Minutes,$RunTime.Seconds,$RunTime.Milliseconds
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Run time was $RunTime"
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Information -WriteBackToHost -Message "Exit code is $ExitCode"
} else {
    Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Script requirements not met:"
    
    if (-not $script:boolScriptIsModulesLoaded) {
        Out-CMTraceLog -LoggingPreference $htLoggingPreference -Logfile $strLogFile -Type Error -WriteBackToHost -Message "Missing required PowerShell module(s) or could not load modules"
    }#if
}#if/else

Get-ConnectionInformation -ErrorAction SilentlyContinue -Verbose:$fasle | Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
Exit $ExitCode
#endregion