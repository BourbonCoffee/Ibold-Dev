# Dell Warranty Information Retrieval Script
# Original reference script: https://www.hull1.com/scriptit/2020/08/28/dell-api-warranty-lookup.html (2020)
# Chris Ibold 3/8/2023
# Generate XML credential file with your API key and secret and place it somewhere.
# Adjust $CredentialPath as needed
# Use: Get-DellWarrantyInfo -ServiceTags TAG1, TAG2, TAG3, TAG4
# If no service tags are supplied, it will use the device you are running it from.

param(
    [Parameter(Mandatory = $false)]
    [string[]]$ServiceTags,
    
    [Parameter(Mandatory = $false)]
    [string]$CredentialsPath = "$env:OneDrive\Documents\DellTDM.xml"
)

# Function to Get Credentials from XML
function Get-StoredCredential {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CredentialsPath
    )

    # Check if credentials file exists
    if (-not (Test-Path $CredentialsPath)) {
        Write-Host "Credentials file not found. Please create it first." -ForegroundColor Red
        
        # Credential Creation Guidance
        Write-Host "`nTo create credentials, run these commands:" -ForegroundColor Yellow
        Write-Host "`$apiKey = ConvertTo-SecureString 'YOUR_API_KEY' -AsPlainText -Force" -ForegroundColor Cyan
        Write-Host "`$apiSecret = ConvertTo-SecureString 'YOUR_API_SECRET' -AsPlainText -Force" -ForegroundColor Cyan
        Write-Host "`$credentials = @{ ApiKey = `$apiKey; ApiSecret = `$apiSecret }" -ForegroundColor Cyan
        Write-Host "`$credentials | Export-Clixml -Path '$CredentialsPath'" -ForegroundColor Cyan
        
        throw "Credentials file is missing"
    }

    try {
        # Import the encrypted credentials
        $storedCredentials = Import-Clixml -Path $CredentialsPath

        # Convert secure strings back to plain text for API use
        $ApiKey = [System.Net.NetworkCredential]::new("", $storedCredentials.ApiKey).Password
        $ApiSecret = [System.Net.NetworkCredential]::new("", $storedCredentials.ApiSecret).Password

        return @{
            ApiKey    = $ApiKey
            ApiSecret = $ApiSecret
        }
    } catch {
        Write-Host "Error reading credentials: $_" -ForegroundColor Red
        throw
    }
}

# Logging Function
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "Info" { Write-Host $logMessage -ForegroundColor Green }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error" { Write-Host $logMessage -ForegroundColor Red }
    }
    
    # Optional: Add logging to file
    # Add-Content -Path "$PSScriptRoot\warranty_log.txt" -Value $logMessage
}

# Validate Service Tag - Thanks GPT for impossible to understand Regex formulas
function Validate-ServiceTag {
    param([string]$ServiceTag)
    
    if ($ServiceTag -notmatch '^[A-Z0-9]{7}$') {
        Write-Log -Message "Invalid Service Tag: $ServiceTag" -Level "Error"
        return $false
    }
    return $true
}

# Get Warranty Information
function Get-DellWarrantyInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ServiceTags,
        
        [Parameter(Mandatory = $true)]
        [string]$ApiKey,
        
        [Parameter(Mandatory = $true)]
        [string]$KeySecret
    )

    # Validate API Credentials
    if ([string]::IsNullOrWhiteSpace($ApiKey) -or [string]::IsNullOrWhiteSpace($KeySecret)) {
        Write-Log -Message "Missing API Credentials" -Level "Error"
        throw "API Key or Secret is missing"
    }

    # Validate Service Tags
    $validServiceTags = $ServiceTags | Where-Object { Validate-ServiceTag $_ }
    
    if ($validServiceTags.Count -eq 0) {
        Write-Log -Message "No valid Service Tags provided" -Level "Error"
        return
    }

    # Authentication
    try {
        $AuthURI = "https://apigtwb2c.us.dell.com/auth/oauth/v2/token"
        $OAuth = "$ApiKey`:$KeySecret"
        $Bytes = [System.Text.Encoding]::ASCII.GetBytes($OAuth)
        $EncodedOAuth = [Convert]::ToBase64String($Bytes)
        
        $Headers = @{ 
            "authorization" = "Basic $EncodedOAuth"
            "Accept"        = "application/json"
        }
        
        $AuthBody = 'grant_type=client_credentials'
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $AuthResult = Invoke-RestMethod -Method Post -Uri $AuthURI -Body $AuthBody -Headers $Headers
        $token = $AuthResult.access_token
        
        Write-Log -Message "Successfully obtained authentication token"
    } catch {
        Write-Log -Message "Authentication Failed: $_" -Level "Error"
        throw
    }

    # Fetch Warranty Information
    try {
        $Headers = @{
            "Accept"        = "application/json"
            "Authorization" = "Bearer $token"
        }
        
        $Params = @{
            servicetags = $validServiceTags -join ","
            Method      = "GET"
        }
        
        $Response = Invoke-RestMethod -Uri "https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements" `
            -Headers $Headers -Body $Params -Method Get -ContentType "application/json"
        
        Write-Log -Message "Successfully retrieved warranty information"
        
        # Process and display results
        foreach ($Record in $Response) {
            $servicetag = $Record.servicetag
            $Device = $Record.productLineDescription
            $ShipDate = ($Record.shipDate | Get-Date -f "MM-dd-yyyy")
            $EndDate = (($Record.entitlements | Select-Object -Last 1).endDate | Get-Date -f "MM-dd-yyyy")
            $Support = ($Record.entitlements | Select-Object -Last 1).serviceLevelDescription
            $today = Get-Date
            
            $DeviceType = switch -Wildcard ($Record) {
                # Check $Record.ProductID for Desktop identifiers
                { $_.ProductID -like '*desktop*' } { 'Desktop' }

                # Check $Record.ProductLineDescription for Laptop identifiers because ProductID is empty for some reason...
                { $_.ProductLineDescription -like '*Latitude*' -or $_.ProductLineDescription -like '*Precision*' } { 'Laptop' }

                # Default case
                default { 'Unknown' }
            }

            
            Write-Host "`nService Tag Details for $servicetag" -ForegroundColor Cyan
            Write-Host "Model         : $Device"
            Write-Host "Type          : $DeviceType"
            Write-Host "Ship Date     : $ShipDate"
            
            if ($today -ge $EndDate) { 
                Write-Host "Warranty Exp. : $EndDate" -NoNewline
                Write-Host "  [WARRANTY EXPIRED]" -ForegroundColor Yellow 
            } else { 
                Write-Host "Warranty Exp. : $EndDate" 
            }
            
            Write-Host "Service Levels:" -ForegroundColor Green
            $Record.entitlements.serviceLevelDescription | Select-Object -Unique | ForEach-Object {
                Write-Host "  - $_"
            }
        }
    } catch {
        Write-Log -Message "Warranty Retrieval Failed: $_" -Level "Error"
        throw
    }
}

# Main Execution
try {
    # Retrieve stored credentials
    $Credentials = Get-StoredCredential -CredentialsPath $CredentialsPath

    # If no service tags provided, prompt user
    if (-not $ServiceTags) {
        $ServiceTagInput = Read-Host "Enter Service Tags (comma-separated)"
        
        # Split input and trim whitespace
        $ServiceTags = $ServiceTagInput -split ',' | ForEach-Object { $_.Trim() }
        
        # If still no tags, use local machine's service tag
        if (-not $ServiceTags) {
            $LocalServiceTag = (Get-WmiObject Win32_SystemEnclosure).SerialNumber
            $ServiceTags = @($LocalServiceTag)
            Write-Log -Message "Using local machine's service tag: $LocalServiceTag"
        }
    }

    # Call warranty info function with retrieved credentials
    Get-DellWarrantyInfo -ServiceTags $ServiceTags -ApiKey $Credentials.ApiKey -KeySecret $Credentials.ApiSecret
} catch {
    Write-Log -Message "Script execution failed: $_" -Level "Error"
}