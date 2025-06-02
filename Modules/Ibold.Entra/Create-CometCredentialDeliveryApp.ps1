<#
.SYNOPSIS
    Creates an app registration in Entra ID for Comet Credential Delivery with minimum required permissions.

.DESCRIPTION
    This script creates an app registration named "Comet Credential Delivery" in the target tenant
    with the minimum necessary Microsoft Graph API permissions to manage Temporary Access Passes (TAPs),
    read user attributes, and send emails. It supports authentication via certificate (default) or client secret.

    NOTE: This script uses Mail.Send permission but operationally restricts sending to a single 
    dedicated service account (tap-delivery@yourdomain.com).

.PARAMETER UseCertificate
    Creates an app registration that uses a certificate for authentication (default).

.PARAMETER UseClientSecret
    Creates an app registration that uses a client secret for authentication.

.PARAMETER AppName
    The name for the app registration. Defaults to "Comet Credential Delivery".

.PARAMETER CertificateValidityYears
    The number of years the certificate should be valid. Defaults to 2 years.

.PARAMETER SecretValidityDays
    The number of days the client secret should be valid. Defaults to 365 days.

.EXAMPLE
    .\Create-CometCredentialDeliveryApp.ps1

.EXAMPLE
    .\Create-CometCredentialDeliveryApp.ps1 -UseClientSecret

.EXAMPLE
    .\Create-CometCredentialDeliveryApp.ps1 -AppName "Custom Credential Delivery" -CertificateValidityYears 3

.NOTES
    Author: Chris Ibold
    Company: Comet Consulting Group
    Version: 1.2
    Date: 2025-05-14
#>

#Requires -Version 5.1

param(
    [switch]$UseClientSecret,
    [switch]$UseCertificate = (-not $UseClientSecret), # Certificate is default if neither is specified
    [string]$AppName = "Comet Credential Delivery",
    [int]$CertificateValidityYears = 2,
    [int]$SecretValidityDays = 365
)

# Check PowerShell version
function Test-PowerShellVersion {
    $isPSCore = $PSVersionTable.PSEdition -eq 'Core'
    $minVersion = [Version]'5.1'
    $currentVersion = $PSVersionTable.PSVersion
    
    if ($isPSCore) {
        Write-Error "This script requires Windows PowerShell 5.1 or higher. PowerShell Core is not supported due to AzureAD module compatibility issues."
        return $false
    }
    
    if ($currentVersion -lt $minVersion) {
        Write-Error "This script requires Windows PowerShell 5.1 or higher. Current version: $currentVersion"
        return $false
    }
    
    return $true
}

# Check and install required modules
function Ensure-RequiredModules {
    $requiredModules = @(
        @{
            Name           = "Microsoft.Graph.Applications"
            MinimumVersion = "1.15.0"
        },
        @{
            Name           = "Microsoft.Graph.Authentication" 
            MinimumVersion = "1.15.0"
        },
        @{
            Name           = "AzureAD"
            MinimumVersion = "2.0.2.130"
        }
    )
    
    $allModulesPresent = $true
    
    foreach ($module in $requiredModules) {
        $installedModule = Get-Module -Name $module.Name -ListAvailable | 
            Where-Object { $_.Version -ge [Version]$module.MinimumVersion } | 
            Sort-Object -Property Version -Descending | 
            Select-Object -First 1
        
        if (-not $installedModule) {
            $allModulesPresent = $false
            Write-Host "Required module $($module.Name) (version $($module.MinimumVersion) or higher) is not installed." -ForegroundColor Yellow
            
            $installConsent = Read-Host "Do you want to install $($module.Name) module now? (Y/N)"
            if ($installConsent -eq 'Y' -or $installConsent -eq 'y') {
                try {
                    Write-Host "Installing $($module.Name) module..." -ForegroundColor Cyan
                    Install-Module -Name $module.Name -MinimumVersion $module.MinimumVersion -Scope CurrentUser -Force -AllowClobber
                    Write-Host "$($module.Name) module installed successfully." -ForegroundColor Green
                } catch {
                    Write-Error "Failed to install $($module.Name) module: $_"
                    return $false
                }
            } else {
                Write-Host "Module installation declined. This module is required to run the script." -ForegroundColor Red
                return $false
            }
        } else {
            Write-Host "Module $($module.Name) version $($installedModule.Version) is already installed." -ForegroundColor Green
        }
    }
    
    # Import modules to ensure they're loaded
    try {
        Import-Module Microsoft.Graph.Applications -MinimumVersion "1.15.0" -ErrorAction Stop
        Import-Module Microsoft.Graph.Authentication -MinimumVersion "1.15.0" -ErrorAction Stop
        Import-Module AzureAD -MinimumVersion "2.0.2.130" -ErrorAction Stop
        return $true
    } catch {
        Write-Error "Failed to import required modules: $_"
        return $false
    }
}

# Validate authentication parameters
if ($UseClientSecret -and $UseCertificate) {
    Write-Error "You cannot specify both -UseClientSecret and -UseCertificate. Choose one authentication method."
    exit 1
}

# If neither is specified, default to certificate
if (-not $UseClientSecret -and -not $UseCertificate) {
    $UseCertificate = $true
}

function New-ApplicationCertificate {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$CertificateName,
        
        [Parameter(Mandatory = $true)]
        [int]$ValidityYears
    )
    
    Write-Host "Creating self-signed certificate for application authentication..." -ForegroundColor Cyan
    
    # Set certificate expiration date
    $notAfter = (Get-Date).AddYears($ValidityYears)
    
    try {
        # Create self-signed certificate
        $cert = New-SelfSignedCertificate -Subject "CN=$CertificateName" `
            -CertStoreLocation "cert:\CurrentUser\My" `
            -KeyExportPolicy Exportable `
            -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
            -NotAfter $notAfter `
            -KeySpec Signature `
            -KeyLength 2048
        
        Write-Host "Certificate created successfully." -ForegroundColor Green
        Write-Host "Certificate Thumbprint: $($cert.Thumbprint)" -ForegroundColor Cyan
        
        return $cert
    } catch {
        Write-Error "Failed to create certificate: $_"
        Write-Host "Note: You may need to run this script as Administrator." -ForegroundColor Yellow
        return $null
    }
}

function Export-CertificateToPfx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $true)]
        [string]$Password
    )
    
    try {
        # Convert plain password to secure string
        $securePassword = ConvertTo-SecureString -String $Password -Force -AsPlainText
        
        # Export certificate with private key
        $Certificate | Export-PfxCertificate -FilePath $FilePath -Password $securePassword | Out-Null
        
        Write-Host "Certificate exported to: $FilePath" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to export certificate as PFX: $_"
        return $false
    }
}

function Export-CertificateToCer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        # Export certificate without private key (public key only)
        $Certificate | Export-Certificate -FilePath $FilePath -Type CERT | Out-Null
        
        Write-Host "Public certificate exported to: $FilePath" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to export certificate as CER: $_"
        return $false
    }
}

function Add-CertificateToApplication {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ApplicationObjectId,
        
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )
    
    try {
        Write-Host "Adding certificate to application using AzureAD module..." -ForegroundColor Cyan
        
        # Convert certificate to base64 string format as required by AzureAD
        $keyValue = [System.Convert]::ToBase64String($Certificate.GetRawCertData())
        $startDate = Get-Date
        $endDate = $Certificate.NotAfter
        
        # Get current credentials if any
        $azureApp = Get-AzureADApplication -ObjectId $ApplicationObjectId
        $keyCredentials = @()
        if ($null -ne $azureApp.KeyCredentials) {
            $keyCredentials = $azureApp.KeyCredentials
        }
        
        # Create the credential object
        $keyCredential = New-Object Microsoft.Open.AzureAD.Model.KeyCredential
        $keyCredential.StartDate = $startDate
        $keyCredential.EndDate = $endDate
        $keyCredential.Type = "AsymmetricX509Cert" 
        $keyCredential.Usage = "Verify"
        $keyCredential.Value = [System.Text.Encoding]::ASCII.GetBytes($keyValue)
        
        # Add to existing credentials
        $keyCredentials += $keyCredential
        
        # Update the application
        Set-AzureADApplication -ObjectId $ApplicationObjectId -KeyCredentials $keyCredentials
        
        Write-Host "Certificate successfully added to application." -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Error adding certificate to application: $_" -ForegroundColor Red
        
        # More detailed error information
        if ($_.Exception.InnerException) {
            Write-Host "Inner exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
        }
        
        Write-Host "`nFalling back to manual certificate upload instructions." -ForegroundColor Yellow
        return $false
    }
}

function Connect-ToAzureServices {
    try {
        # Check if already connected to Microsoft Graph
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        $azureADContext = $null
        
        try {
            # Try to get AzureAD context (will throw if not connected)
            $azureADContext = Get-AzureADCurrentSessionInfo -ErrorAction Stop
        } catch {
            # Not connected to AzureAD yet
            $azureADContext = $null
        }
        
        # If both are already connected, check if they're for the same tenant
        if ($null -ne $graphContext -and $null -ne $azureADContext) {
            if ($graphContext.TenantId -eq $azureADContext.TenantId.ToString()) {
                Write-Host "Already connected to Microsoft Graph and AzureAD as: $($graphContext.Account)" -ForegroundColor Green
                
                # Get tenant details for confirmation
                $tenant = Get-MgOrganization
                Write-Host "Target tenant: $($tenant.DisplayName) ($($tenant.Id))" -ForegroundColor Cyan
                
                # Confirm with user
                $confirm = Read-Host "Is this the correct target tenant? (Y/N)"
                if ($confirm -ne "Y" -and $confirm -ne "y") {
                    Write-Host "Operation cancelled by user. Please reconnect to the correct tenant." -ForegroundColor Yellow
                    # Disconnect both services
                    Disconnect-MgGraph | Out-Null
                    Disconnect-AzureAD | Out-Null
                    return Connect-ToAzureServices # Recursive call to reconnect
                }
                
                return @{
                    GraphContext   = $graphContext
                    AzureADContext = $azureADContext
                    TenantDetails  = $tenant
                }
            } else {
                Write-Host "Connected to different tenants in Graph and AzureAD. Reconnecting to ensure consistency." -ForegroundColor Yellow
                Disconnect-MgGraph | Out-Null
                Disconnect-AzureAD | Out-Null
                $graphContext = $null
                $azureADContext = $null
            }
        }
        
        # Connect to Microsoft Graph if not already connected
        if ($null -eq $graphContext) {
            # Connect with only the permission needed to create/modify applications
            Connect-MgGraph -Scopes "Application.ReadWrite.All" -ErrorAction Stop
            Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            $graphContext = Get-MgContext
        }
        
        # Connect to AzureAD if not already connected
        if ($null -eq $azureADContext) {
            Connect-AzureAD -TenantId $graphContext.TenantId -ErrorAction Stop
            Write-Host "Successfully connected to AzureAD" -ForegroundColor Green
            $azureADContext = Get-AzureADCurrentSessionInfo
        }
        
        # Get tenant details for confirmation
        $tenant = Get-MgOrganization
        Write-Host "Target tenant: $($tenant.DisplayName) ($($tenant.Id))" -ForegroundColor Cyan
        
        # Confirm with user
        $confirm = Read-Host "Is this the correct target tenant? (Y/N)"
        if ($confirm -ne "Y" -and $confirm -ne "y") {
            Write-Host "Operation cancelled by user. Please reconnect to the correct tenant." -ForegroundColor Yellow
            Disconnect-MgGraph | Out-Null
            Disconnect-AzureAD | Out-Null
            return Connect-ToAzureServices # Recursive call to reconnect
        }
        
        return @{
            GraphContext   = $graphContext
            AzureADContext = $azureADContext
            TenantDetails  = $tenant
        }
    } catch {
        Write-Host "Failed to connect to Microsoft services: $_" -ForegroundColor Red
        return $null
    }
}

function Get-DomainFromTenant {
    try {
        $domains = Get-MgDomain
        
        # Find the primary verified domain
        $primaryDomain = $domains | Where-Object { $_.IsDefault -eq $true -and $_.IsVerified -eq $true } | 
            Select-Object -First 1
        
        if ($null -eq $primaryDomain) {
            # If no primary domain, take the first verified domain
            $primaryDomain = $domains | Where-Object { $_.IsVerified -eq $true } | 
                Select-Object -First 1
        }
        
        if ($null -eq $primaryDomain) {
            Write-Host "Unable to find a verified domain in the tenant." -ForegroundColor Yellow
            return "tap-delivery@unknown-domain.com"
        }
        
        # Use a dedicated service account name that clearly indicates its purpose
        return "tap-delivery@$($primaryDomain.Id)"
    } catch {
        Write-Host "Error retrieving domains: $_" -ForegroundColor Red
        return "tap-delivery@unknown-domain.com"
    }
}

function New-CometCredentialDeliveryApp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppName,
        
        [Parameter(Mandatory = $false)]
        [bool]$UseSecret,
        
        [Parameter(Mandatory = $false)]
        [bool]$UseCert,
        
        [Parameter(Mandatory = $false)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        
        [Parameter(Mandatory = $false)]
        [int]$SecretValidityDays
    )
    
    try {
        # Check if app already exists
        $existingApp = Get-MgApplication -Filter "DisplayName eq '$AppName'"
        
        if ($null -ne $existingApp) {
            Write-Host "Application '$AppName' already exists with ID: $($existingApp.Id)" -ForegroundColor Yellow
            
            # Confirm overwrite
            $confirmOverwrite = Read-Host "Do you want to update the existing application? (Y/N)"
            if ($confirmOverwrite -ne "Y" -and $confirmOverwrite -ne "y") {
                Write-Host "Operation cancelled. The existing app will not be modified." -ForegroundColor Yellow
                return $existingApp
            }
            
            # Will modify the existing app below
            $appId = $existingApp.Id
            $app = $existingApp
        } else {
            # Create the required permissions for Microsoft Graph API with least privilege
            $permissionsList = @(
                # User.Read.All (for reading user properties including extensionAttribute15)
                @{
                    Id   = "df021288-bdef-4463-88db-98f22de89214"
                    Type = "Role"
                },
                # UserAuthenticationMethod.ReadWrite.All (for TAP management)
                @{
                    Id   = "50483e42-d915-4231-9639-7fdb7fd190e5"
                    Type = "Role"
                },
                # Mail.Send (for sending emails)
                # Note: While this permission allows sending as any user, 
                # operationally we restrict it to a single service account
                @{
                    Id   = "b633e1c5-b582-4048-a93e-9f11b44c7e96"
                    Type = "Role"
                },
                # Domain.Read.All (For gathering the default domain)
                @{
                    Id   = "dbb9058a-0e50-45d7-ae91-66909b5d4664"
                    Type = "Role"
                },
                # Group.Read.All (For getting security group with members)
                @{
                    Id   = "5b567255-7703-4780-807c-7be8301ae99b"
                    Type = "Role"
                }
            )
            
            $msGraphAccess = @{
                ResourceAppId  = "00000003-0000-0000-c000-000000000000" # Microsoft Graph
                ResourceAccess = $permissionsList
            }
            
            # Create new application
            Write-Host "Creating new application registration: '$AppName'..." -ForegroundColor Cyan
            
            $params = @{
                DisplayName            = $AppName
                Web                    = @{
                    RedirectUris = @("https://localhost")
                }
                RequiredResourceAccess = @($msGraphAccess)
                SignInAudience         = "AzureADMyOrg"
                Notes                  = "Created by Comet Consulting Group for TAP credential delivery"
            }
            
            # Create the application
            $app = New-MgApplication @params
            Write-Host "Application created with ID: $($app.Id)" -ForegroundColor Green
            
            # Create service principal
            Write-Host "Creating service principal..." -ForegroundColor Cyan
            $sp = New-MgServicePrincipal -AppId $app.AppId
            Write-Host "Service principal created with ID: $($sp.Id)" -ForegroundColor Green
            
            # Wait for service principal propagation
            Write-Host "Waiting for service principal propagation (30 seconds)..." -ForegroundColor Cyan
            Start-Sleep -Seconds 30
        }
        
        # Get the AzureAD application with the same AppId
        $azureADApp = Get-AzureADApplication -Filter "AppId eq '$($app.AppId)'"
        
        # Add authentication credentials (secret or certificate)
        if ($UseSecret) {
            # Add client secret
            Write-Host "Creating client secret valid for $SecretValidityDays days..." -ForegroundColor Cyan
            $endDateTime = (Get-Date).AddDays($SecretValidityDays)
            
            $passwordCred = @{
                DisplayName = "Created by script on $(Get-Date -Format 'yyyy-MM-dd')"
                EndDateTime = $endDateTime
            }
            
            $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $passwordCred
            
            # Store authentication details for return
            $authDetails = @{
                Type           = "Secret"
                ClientId       = $app.AppId
                ClientSecret   = $secret.SecretText
                ExpirationDate = $endDateTime
            }
        } elseif ($UseCert -and $null -ne $Certificate) {
            # Add certificate to the application using AzureAD module
            $certificateAdded = Add-CertificateToApplication -ApplicationObjectId $azureADApp.ObjectId -Certificate $Certificate
            
            # Store authentication details for return
            $authDetails = @{
                Type                         = "Certificate"
                ClientId                     = $app.AppId
                CertificateThumbprint        = $Certificate.Thumbprint
                ExpirationDate               = $Certificate.NotAfter
                CertificateAddedSuccessfully = $certificateAdded
            }
        } else {
            Write-Error "Neither client secret nor certificate authentication was specified or certificate is missing."
            return $null
        }
        
        return @{
            Application = $app
            AuthDetails = $authDetails
        }
    } catch {
        Write-Error "Error creating application: $_"
        throw
    }
}

# Main script execution
try {
    # Clear the screen to start fresh
    Clear-Host
    
    Write-Host "=====================================================" -ForegroundColor Green
    Write-Host "  COMET CREDENTIAL DELIVERY APP REGISTRATION SCRIPT" -ForegroundColor Green
    Write-Host "=====================================================" -ForegroundColor Green
    Write-Host "This script will create an application registration in Entra ID"
    Write-Host "with minimum required permissions to create TAPs and send emails."
    Write-Host
    
    # Check PowerShell version
    if (-not (Test-PowerShellVersion)) {
        exit 1
    }
    
    # Check and install required modules
    if (-not (Ensure-RequiredModules)) {
        exit 1
    }
    
    # Connect to Microsoft Graph and AzureAD
    $connectionContext = Connect-ToAzureServices
    if ($null -eq $connectionContext) {
        Write-Error "Failed to connect to Microsoft services. Exiting."
        exit 1
    }
    
    # Get default email for this tenant
    $fromEmailAddress = Get-DomainFromTenant
    Write-Host "Using sender email address: $fromEmailAddress" -ForegroundColor Cyan
    Write-Host "NOTE: This is a dedicated service account for sending TAP credentials." -ForegroundColor Yellow
    Write-Host "      You will need to create this mailbox in your tenant before using the script." -ForegroundColor Yellow
    
    # Create secure folder for certificate/credential export
    $secureFolderPath = Join-Path -Path $env:USERPROFILE -ChildPath "CometCredentialDelivery"
    if (-not (Test-Path -Path $secureFolderPath)) {
        New-Item -ItemType Directory -Path $secureFolderPath | Out-Null
    }
    
    # Define authentication parameters
    if ($UseCertificate) {
        Write-Host "Creating application with certificate authentication..." -ForegroundColor Cyan
        
        # Create self-signed certificate
        $cert = New-ApplicationCertificate -CertificateName "$AppName Certificate" -ValidityYears $CertificateValidityYears
        if ($null -eq $cert) {
            Write-Error "Failed to create certificate. Exiting."
            exit 1
        }
        
        # Generate a secure password for the certificate
        $certPassword = [System.Guid]::NewGuid().ToString()
        
        # Create file paths for both formats
        $pfxFileName = "CometCredentialDelivery_$(Get-Date -Format 'yyyyMMdd_HHmmss').pfx"
        $cerFileName = "CometCredentialDelivery_$(Get-Date -Format 'yyyyMMdd_HHmmss').cer"
        $pfxFilePath = Join-Path -Path $secureFolderPath -ChildPath $pfxFileName
        $cerFilePath = Join-Path -Path $secureFolderPath -ChildPath $cerFileName
        
        # Export the certificate in both formats
        Export-CertificateToPfx -Certificate $cert -FilePath $pfxFilePath -Password $certPassword
        Export-CertificateToCer -Certificate $cert -FilePath $cerFilePath
        
        # Create the application with certificate
        $appResult = New-CometCredentialDeliveryApp -AppName $AppName -UseSecret $false -UseCert $true -Certificate $cert
    } elseif ($UseClientSecret) {
        Write-Host "Creating application with client secret authentication..." -ForegroundColor Cyan
        
        # Create the application with client secret
        $appResult = New-CometCredentialDeliveryApp -AppName $AppName -UseSecret $true -UseCert $false -SecretValidityDays $SecretValidityDays
    } else {
        Write-Error "You must specify either -UseClientSecret or -UseCertificate."
        exit 1
    }
    
    if ($null -eq $appResult) {
        Write-Error "Failed to create the application. Exiting."
        exit 1
    }
    
    # Extract application details
    $app = $appResult.Application
    $authDetails = $appResult.AuthDetails
    $tenantId = $connectionContext.GraphContext.TenantId
    
    # Output application details
    Write-Host "`n===== APPLICATION DETAILS =====" -ForegroundColor Yellow
    Write-Host "Application Name: $AppName" -ForegroundColor White
    Write-Host "Application (Client) ID: $($app.AppId)" -ForegroundColor White
    Write-Host "Directory (Tenant) ID: $tenantId" -ForegroundColor White
    
    if ($authDetails.Type -eq "Secret") {
        Write-Host "Authentication Type: Client Secret" -ForegroundColor White
        Write-Host "Client Secret: $($authDetails.ClientSecret)" -ForegroundColor White
        Write-Host "Secret Expiration: $($authDetails.ExpirationDate)" -ForegroundColor White
    } else {
        Write-Host "Authentication Type: Certificate" -ForegroundColor White
        Write-Host "Certificate Thumbprint: $($authDetails.CertificateThumbprint)" -ForegroundColor White
        Write-Host "Certificate Expiration: $($authDetails.ExpirationDate)" -ForegroundColor White
        Write-Host "Certificate File (PFX with private key): $pfxFilePath" -ForegroundColor White
        Write-Host "Certificate Password: $certPassword" -ForegroundColor White
        Write-Host "Certificate File (CER public key): $cerFilePath" -ForegroundColor White
        
        if ($authDetails.CertificateAddedSuccessfully -ne $true) {
            Write-Host "`nWARNING: The certificate was not automatically added to the application." -ForegroundColor Yellow
            Write-Host "You will need to manually upload the certificate. See instructions below." -ForegroundColor Yellow
        }
    }
    
    Write-Host "From Email Address: $fromEmailAddress" -ForegroundColor White
    Write-Host "NOTE: This is a dedicated service account for sending TAP credentials." -ForegroundColor White
    Write-Host "      Ensure this mailbox exists in your tenant before using the script." -ForegroundColor White
    Write-Host "--------------------------------" -ForegroundColor Yellow
    Write-Host "Permissions Assigned (Least Privilege):" -ForegroundColor Yellow
    Write-Host "- User.Read.All (read user properties including extensionAttribute15)" -ForegroundColor White
    Write-Host "- UserAuthenticationMethod.ReadWrite.All (manage TAPs)" -ForegroundColor White
    Write-Host "- Mail.Send (send emails from dedicated service account)" -ForegroundColor White
    Write-Host "- Domain.Read.All (read domain information)" -ForegroundColor White
    Write-Host "=================================" -ForegroundColor Yellow
    
    # Save to file
    $credentialFilePath = Join-Path -Path $secureFolderPath -ChildPath "AppCredentials_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    
    $credentialInfo = @"
===== COMET CREDENTIAL DELIVERY APP INFO =====
Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Application Name: $AppName
Application (Client) ID: $($app.AppId)
Directory (Tenant) ID: $tenantId
Authentication Type: $($authDetails.Type)
"@
    
    if ($authDetails.Type -eq "Secret") {
        $credentialInfo += @"

Client Secret: $($authDetails.ClientSecret)
Secret Expiration: $($authDetails.ExpirationDate)
"@
    } else {
        $credentialInfo += @"

Certificate Thumbprint: $($authDetails.CertificateThumbprint)
Certificate Expiration: $($authDetails.ExpirationDate)
Certificate File (PFX with private key): $pfxFilePath
Certificate Password: $certPassword
Certificate File (CER public key): $cerFilePath
"@
        
        if ($authDetails.CertificateAddedSuccessfully -ne $true) {
            $credentialInfo += @"

IMPORTANT: The certificate must be manually uploaded to the application.
"@
        }
    }
    
    $credentialInfo += @"

From Email Address: $fromEmailAddress
NOTE: This is a dedicated service account that should be created in your tenant

Permissions Assigned (Least Privilege):
- User.Read.All (read user properties including extensionAttribute15)
- UserAuthenticationMethod.ReadWrite.All (manage TAPs)
- Mail.Send (send emails from dedicated service account)
- Domain.Read.All (read domain information)
============================================
"@
    
    # Write credential info to file
    $credentialInfo | Out-File -FilePath $credentialFilePath -Encoding utf8
    Write-Host "Credential information saved to: $credentialFilePath" -ForegroundColor Green
    
    # If certificate was not automatically added, show manual upload instructions
    if ($UseCertificate -and $authDetails.CertificateAddedSuccessfully -ne $true) {
        Write-Host "`n===== CERTIFICATE UPLOAD INSTRUCTIONS =====" -ForegroundColor Yellow
        Write-Host "The certificate could not be added automatically. Please upload it manually:" -ForegroundColor White
        Write-Host "1. Go to Azure Portal > App registrations > $AppName" -ForegroundColor White
        Write-Host "2. Navigate to 'Certificates & secrets'" -ForegroundColor White
        Write-Host "3. Select 'Certificates' tab" -ForegroundColor White
        Write-Host "4. Click '+ Upload certificate'" -ForegroundColor White
        Write-Host "5. Browse and select this file: $cerFilePath" -ForegroundColor White
        Write-Host "6. Click 'Add'" -ForegroundColor White
        Write-Host "=============================================" -ForegroundColor Yellow
    }
    
    # Generate admin consent URL
    $adminConsentUrl = "https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$($app.AppId)"
    
    # Prompt for admin consent
    Write-Host "`nNOTE: An administrator needs to grant consent to the required permissions." -ForegroundColor Yellow
    $grantConsent = Read-Host "Would you like to open the admin consent URL now? (Y/N)"
    
    if ($grantConsent -eq "Y" -or $grantConsent -eq "y") {
        Write-Host "Opening the following URL in your default browser: $adminConsentUrl" -ForegroundColor Cyan
        Start-Process $adminConsentUrl
    } else {
        Write-Host "You can grant admin consent later using this URL:" -ForegroundColor Cyan
        Write-Host $adminConsentUrl -ForegroundColor White
    }
    
    # Show service account creation instructions
    Write-Host "`n===== SERVICE ACCOUNT SETUP INSTRUCTIONS =====" -ForegroundColor Yellow
    Write-Host "You must create the dedicated service account for sending TAP credentials:" -ForegroundColor White
    Write-Host "1. Go to Microsoft 365 Admin Center > Users > Active users > Add a user" -ForegroundColor White
    Write-Host "2. Create a new user with the following details:" -ForegroundColor White
    Write-Host "   - Display name: TAP Credential Delivery" -ForegroundColor White
    Write-Host "   - Email/UPN: $fromEmailAddress" -ForegroundColor White
    Write-Host "   - Password: Set a strong, complex password" -ForegroundColor White
    Write-Host "3. Configure the mailbox:" -ForegroundColor White
    Write-Host "   - Set a profile picture that clearly identifies this as an automated system" -ForegroundColor White
    Write-Host "   - Configure an auto-reply indicating this is an unmonitored mailbox" -ForegroundColor White
    Write-Host "   - Set forwarding rules to direct replies to your IT support team" -ForegroundColor White
    Write-Host "4. Apply security settings:" -ForegroundColor White
    Write-Host "   - Enable MFA for the account" -ForegroundColor White
    Write-Host "   - Add the account to a security group with restricted permissions" -ForegroundColor White
    Write-Host "=============================================" -ForegroundColor Yellow
    
    # Show sample command for running the TAP script
    Write-Host "`n===== SAMPLE COMMAND TO RUN THE TAP SCRIPT =====" -ForegroundColor Yellow
    
    if ($authDetails.Type -eq "Secret") {
        Write-Host ".\New-TapCredentialAndDelivery.ps1 -TenantId '$tenantId' -SecurityGroupId '<group-id>' -ClientId '$($app.AppId)' -ClientSecret '$($authDetails.ClientSecret)' -SupportEmail $fromEmailAddress" -ForegroundColor Cyan
    } else {
        Write-Host ".\New-TapCredentialAndDelivery.ps1 -TenantId '$tenantId' -SecurityGroupId '<group-id>' -ClientId '$($app.AppId)' -CertificateThumbprint '$($authDetails.CertificateThumbprint)' -SupportEmail $fromEmailAddress" -ForegroundColor Cyan
    }
    
    Write-Host "=============================================" -ForegroundColor Yellow
    
    Write-Host "`nApplication setup completed successfully!" -ForegroundColor Green
    Write-Host "IMPORTANT: Store your credential information securely and delete the files after use!" -ForegroundColor Yellow
} catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    
    # More detailed error information if available
    if ($_.Exception.InnerException) {
        Write-Host "Inner exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    
    exit 1
} finally {
    # Disconnect from Graph API and AzureAD
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Disconnect-AzureAD -ErrorAction SilentlyContinue | Out-Null
}