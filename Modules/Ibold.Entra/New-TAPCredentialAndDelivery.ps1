<#
.SYNOPSIS
    Creates one-time-use TAP credentials and sends them via encrypted email.

.DESCRIPTION
    This script generates Temporary Access Pass (TAP) credentials for users in an Entra security group
    in the target tenant. It then sends these credentials via encrypted email to the users' 
    corresponding email addresses in the source tenant, as specified in extensionAttribute15.

    NOTE: This script uses a dedicated service account (tap-delivery@yourdomain.com) for sending
    emails to ensure proper tracking and audit capabilities.

.PARAMETER TenantId
    The ID of the Entra ID tenant where the operations will be performed.

.PARAMETER SecurityGroupId
    The ID of the Entra security group containing users who need TAP credentials.

.PARAMETER TapLifetimeMinutes
    The lifetime of the Temporary Access Pass in minutes. Default is 1440 (24 hours).

.PARAMETER ClientId
    The client ID of the Comet Credential Delivery app registration.

.PARAMETER ClientSecret
    The client secret of the Comet Credential Delivery app registration.
    Required if not using certificate authentication.

.PARAMETER CertificateThumbprint
    The thumbprint of the certificate to use for authentication.
    Required if not using client secret authentication.

.PARAMETER EmailSubject
    The subject line for the email that will be sent. Default is "Your Temporary Access Credentials".

.PARAMETER EmailTemplatePath
    Path to a custom HTML email template. If not provided, a default template will be used.

.PARAMETER SupportEmail
    Email address for support that will be included in the email. Default is "support@cometcg.com".

.PARAMETER FromEmailAddress
    Optional email address to send from. If not provided, will use tap-delivery@[tenant-domain].com.

.PARAMETER Force
    If specified, the script will overwrite existing TAPs for users.

.EXAMPLE
    .\New-TapCredentialAndDelivery.ps1 -TenantId "12345678-1234-1234-1234-123456789012" -SecurityGroupId "87654321-4321-4321-4321-210987654321" -ClientId "app-id-here" -ClientSecret "secret-here"

.EXAMPLE
    .\New-TapCredentialAndDelivery.ps1 -TenantId "12345678-1234-1234-1234-123456789012" -SecurityGroupId "87654321-4321-4321-4321-210987654321" -ClientId "app-id-here" -CertificateThumbprint "certificate-thumbprint" -SupportEmail "helpdesk@contoso.com"

.NOTES
    Author: Chris Ibold
    Company: Comet Consulting Group
    Version: 1.1
    Date: 2025-05-14
#>

#Requires -Version 5.1

param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $true)]
    [string]$SecurityGroupId,
    
    [Parameter(Mandatory = $false)]
    [int]$TapLifetimeMinutes = 1440,
    
    [Parameter(Mandatory = $true)]
    [string]$ClientId,
    
    [Parameter(ParameterSetName = 'Secret')]
    [string]$ClientSecret,
    
    [Parameter(ParameterSetName = 'Certificate')]
    [string]$CertificateThumbprint,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailSubject = "Your Temporary Access Credentials",
    
    [Parameter(Mandatory = $false)]
    [string]$EmailTemplatePath,
    
    [Parameter(Mandatory = $true)]
    [string]$SupportEmail,
    
    [Parameter(Mandatory = $false)]
    [string]$FromEmailAddress,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# Check PowerShell version
function Test-PowerShellVersion {
    $isPSCore = $PSVersionTable.PSEdition -eq 'Core'
    $minVersion = [Version]'5.1'
    $currentVersion = $PSVersionTable.PSVersion
    
    if ($isPSCore) {
        Write-Error "This script requires Windows PowerShell 5.1 or higher. PowerShell Core is not supported due to module compatibility issues."
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
    # Define only the specific Microsoft Graph sub-modules we need
    $requiredModules = @(
        @{
            Name           = "Microsoft.Graph.Authentication"
            MinimumVersion = "1.15.0"
        },
        @{
            Name           = "Microsoft.Graph.Users"
            MinimumVersion = "1.15.0"
        },
        @{
            Name           = "Microsoft.Graph.Groups"
            MinimumVersion = "1.15.0"
        },
        @{
            Name           = "Microsoft.Graph.Identity.DirectoryManagement"
            MinimumVersion = "1.15.0"
        }
    )
    
    $allModulesPresent = $true
    
    Write-Host "Checking required modules for app authentication..." -ForegroundColor Cyan
    
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
                    Write-Host "Failed to install $($module.Name) module: $_" -ForegroundColor Red
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
    
    # Import only the specific modules we need instead of the entire Microsoft.Graph module
    try {
        Write-Host "Importing required Microsoft Graph modules..." -ForegroundColor Cyan
        
        # Import modules one by one to avoid the function capacity issue
        foreach ($module in $requiredModules) {
            Import-Module $module.Name -MinimumVersion $module.MinimumVersion -ErrorAction Stop
            Write-Host "Successfully imported $($module.Name)" -ForegroundColor Green
        }
        
        return $true
    } catch {
        Write-Host "Failed to import required modules: $_" -ForegroundColor Red
        return $false
    }
}

# Validate authentication parameters
if (-not $ClientSecret -and -not $CertificateThumbprint) {
    Write-Error "You must provide either -ClientSecret or -CertificateThumbprint for authentication."
    exit 1
}

# Script variables
$script:tapCredentials = @{}
$script:logFolder = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath "CometTAPLogs"
$script:logFile = Join-Path -Path $script:logFolder -ChildPath "TapCredentialDelivery_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:errorCount = 0
$script:successCount = 0
$script:userCount = 0

# Ensure log folder exists
if (-not (Test-Path -Path $script:logFolder)) {
    New-Item -ItemType Directory -Path $script:logFolder -Force | Out-Null
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "INFO" { Write-Host $logMessage -ForegroundColor Cyan }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
    }
    
    $logMessage | Out-File -FilePath $script:logFile -Append
}

function Connect-ToMsGraphWithAuth {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret,
        
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint
    )
    
    try {
        Write-Log "Connecting to Microsoft Graph API for tenant ID: $TenantId" -Level "INFO"
        
        # Check if already connected
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -ne $context) {
            if ($context.TenantId -eq $TenantId) {
                Write-Log "Already connected to Microsoft Graph for tenant $TenantId" -Level "INFO"
                
                # Get tenant information for verification
                $tenant = Get-MgOrganization
                $connectedTenantId = $tenant.Id
                $tenantDisplayName = $tenant.DisplayName
                
                Write-Log "Connected tenant: $tenantDisplayName ($connectedTenantId)" -Level "INFO"
                
                # Check if authenticated with the correct app
                if ($context.ClientId -eq $ClientId) {
                    Write-Log "Authenticated with the correct application ID: $ClientId" -Level "INFO"
                    return $connectedTenantId
                } else {
                    Write-Log "Connected with application ID $($context.ClientId), but requested application ID is $ClientId. Reconnecting..." -Level "WARNING"
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                }
            } else {
                Write-Log "Connected to tenant $($context.TenantId), but requested tenant ID is $TenantId. Reconnecting..." -Level "WARNING"
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            }
        }
        
        if ($CertificateThumbprint) {
            # Certificate-based authentication (application auth flow)
            Write-Log "Using certificate authentication with thumbprint: $CertificateThumbprint" -Level "INFO"
            
            # Find the certificate in the certificate store
            $certificate = Get-Item -Path "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction Stop
            
            if ($null -eq $certificate) {
                Write-Log "Certificate with thumbprint $CertificateThumbprint not found in the certificate store." -Level "ERROR"
                return $false
            }
            
            # Connect to Microsoft Graph using certificate authentication (app-only flow)
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop
            Write-Log "Successfully connected using certificate authentication" -Level "SUCCESS"
        } else {
            # Client secret authentication (application auth flow)
            Write-Log "Using client secret authentication" -Level "INFO"
            
            # Convert the client secret to a secure string
            $secureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
            
            # Connect to Microsoft Graph using client secret (app-only flow)
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -ClientSecret $secureClientSecret -ErrorAction Stop
            Write-Log "Successfully connected using client secret authentication" -Level "SUCCESS"
        }
        
        # Get tenant information for verification
        try {
            $tenant = Get-MgOrganization
            $connectedTenantId = $tenant.Id
            $tenantDisplayName = $tenant.DisplayName
            
            if ($connectedTenantId -ne $TenantId) {
                Write-Log "WARNING: Connected to tenant ID $connectedTenantId but requested tenant ID was $TenantId" -Level "WARNING"
            }
            
            Write-Log "Connected to tenant: $tenantDisplayName ($connectedTenantId)" -Level "INFO"
            return $connectedTenantId
        } catch {
            Write-Log "Failed to retrieve tenant information: $_" -Level "ERROR"
            return $false
        }
    } catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Level "ERROR"
        
        if ($_.Exception.Message -like "*AADSTS*") {
            Write-Log "Authentication error detected. This may be due to insufficient permissions or incorrect credentials." -Level "ERROR"
            Write-Log "Ensure that the app registration has the required API permissions and they've been admin-consented." -Level "ERROR"
        }
        
        return $false
    }
}

function Get-SecurityGroupMembers {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )
    
    try {
        Write-Log "Retrieving members of security group $GroupId..." -Level "INFO"
        
        $members = @()
        $pageSize = 100
        $groupMembers = $null
        
        # Use paging to handle large groups efficiently
        $groupMembers = Get-MgGroupMember -GroupId $GroupId -All -PageSize $pageSize
        
        foreach ($member in $groupMembers) {
            # Get detailed user information only if it's a user (not a group or service principal)
            if ($member.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.user") {
                try {
                    $userId = $member.Id
                    # Only retrieve the properties we need
                    $user = Get-MgUser -UserId $userId -Property "id,displayName,userPrincipalName" -ErrorAction Stop
                    
                    if ($null -ne $user) {
                        $members += $user
                    }
                } catch {
                    Write-Log "Error retrieving user details for member ID $($member.Id): $_" -Level "ERROR"
                }
            }
        }
        
        if ($members.Count -eq 0) {
            Write-Log "No user members found in security group $GroupId." -Level "WARNING"
        } else {
            Write-Log "Retrieved $($members.Count) users from security group" -Level "SUCCESS"
        }
        
        return $members
    } catch {
        $errorMessage = $_
        
        # Check for specific error conditions
        if ($errorMessage -like "*Resource '$GroupId' does not exist*") {
            Write-Log "Security group with ID '$GroupId' does not exist in the tenant." -Level "ERROR"
        } elseif ($errorMessage -like "*Authorization_RequestDenied*") {
            Write-Log "Access denied. The authenticated application does not have permission to read group members." -Level "ERROR"
            Write-Log "Ensure the app has the Group.Read.All or Directory.Read.All permission and it has been admin-consented." -Level "ERROR"
        } else {
            Write-Log "Error retrieving security group members: $errorMessage" -Level "ERROR"
        }
        
        return @()
    }
}

function Get-SourceEmailFromExtensionAttribute {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )
    
    try {
        # Get the user's extensionAttribute15 which should contain the source email address
        # Using specific property selection for least privilege
        $user = Get-MgUser -UserId $UserId -Property "id,displayName,userPrincipalName,onPremisesExtensionAttributes" -ErrorAction Stop
        
        if ($null -ne $user.OnPremisesExtensionAttributes -and 
            $null -ne $user.OnPremisesExtensionAttributes.ExtensionAttribute15 -and 
            $user.OnPremisesExtensionAttributes.ExtensionAttribute15 -like "*@*") {
            
            $sourceEmail = $user.OnPremisesExtensionAttributes.ExtensionAttribute15.Trim()
            
            # Validate email format
            if ($sourceEmail -match "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") {
                return $sourceEmail
            } else {
                Write-Log "User $($user.DisplayName) ($($user.UserPrincipalName)) has an invalid email format in extensionAttribute15: $sourceEmail" -Level "WARNING"
                return $null
            }
        }
        
        # If no valid email in extensionAttribute15, log a warning and return null
        Write-Log "User $($user.DisplayName) ($($user.UserPrincipalName)) does not have a valid source email address in extensionAttribute15" -Level "WARNING"
        return $null
    } catch {
        Write-Log "Error retrieving source email address for user ID $($UserId): $_" -Level "ERROR"
        return $null
    }
}

function New-TemporaryAccessPass {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [int]$LifetimeInMinutes,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )
    
    try {
        # Check if user already has an active TAP
        $uri = "https://graph.microsoft.com/beta/users/$UserId/authentication/temporaryAccessPassMethods"
        $existingTaps = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        
        if ($existingTaps.value.Count -gt 0 -and -not $Force) {
            $tapId = $existingTaps.value[0].id
            $tapState = $existingTaps.value[0].methodUsabilityInformation.state
            
            if ($tapState -eq "Enabled") {
                Write-Log "User $UserId already has an active Temporary Access Pass (ID: $tapId). Use -Force to overwrite." -Level "WARNING"
            } else {
                Write-Log "User $UserId has a disabled Temporary Access Pass (ID: $tapId). It will be replaced." -Level "INFO"
                $Force = $true
            }
            
            if (-not $Force) {
                return $null
            }
        }
        
        if ($existingTaps.value.Count -gt 0 -and $Force) {
            # Delete existing TAPs if Force is specified or TAP is disabled
            foreach ($tap in $existingTaps.value) {
                $deleteUri = "https://graph.microsoft.com/beta/users/$UserId/authentication/temporaryAccessPassMethods/$($tap.id)"
                Invoke-MgGraphRequest -Method DELETE -Uri $deleteUri -ErrorAction Stop
                Write-Log "Deleted existing Temporary Access Pass (ID: $($tap.id)) for user $UserId" -Level "INFO"
            }
        }
        
        # Create new Temporary Access Pass
        $tapCreationBody = @{
            isUsableOnce      = $true
            lifetimeInMinutes = $LifetimeInMinutes
        } | ConvertTo-Json
        
        $newTap = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $tapCreationBody -ContentType "application/json" -ErrorAction Stop
        
        if ($null -ne $newTap.temporaryAccessPass) {
            Write-Log "Successfully created one-time-use Temporary Access Pass for user $UserId (Expires: $($newTap.lifetimeInMinutes) minutes)" -Level "SUCCESS"
            return $newTap.temporaryAccessPass
        } else {
            Write-Log "Created TAP but did not receive a valid TAP value from API." -Level "ERROR"
            return $null
        }
    } catch {
        $errorMsg = $_
        
        if ($errorMsg -like "*Authorization_RequestDenied*") {
            Write-Log "Access denied when creating Temporary Access Pass. The app may not have the UserAuthenticationMethod.ReadWrite.All permission." -Level "ERROR"
        } elseif ($errorMsg -like "*Request_ResourceNotFound*") {
            Write-Log "User with ID $UserId not found when creating TAP. The user may no longer exist." -Level "ERROR"
        } else {
            Write-Log "Error creating Temporary Access Pass for user $($UserId): $errorMsg" -Level "ERROR"
        }
        
        return $null
    }
}

function Get-EmailTemplate {
    param(
        [Parameter(Mandatory = $false)]
        [string]$TemplatePath,
        
        [Parameter(Mandatory = $true)]
        [string]$UserDisplayName,
        
        [Parameter(Mandatory = $true)]
        [string]$TemporaryAccessPass,
        
        [Parameter(Mandatory = $true)]
        [string]$SupportEmailAddress,
        
        [Parameter(Mandatory = $false)]
        [string]$ExpirationMinutes = "1440"
    )
    
    if ($TemplatePath -and (Test-Path -Path $TemplatePath)) {
        try {
            $template = Get-Content -Path $TemplatePath -Raw -ErrorAction Stop
            Write-Log "Using custom email template from: $TemplatePath" -Level "INFO"
        } catch {
            Write-Log "Error reading custom template: $_. Falling back to default template." -Level "WARNING"
            $template = Get-DefaultEmailTemplate
        }
    } else {
        $template = Get-DefaultEmailTemplate
    }
    
    # Calculate expiration time
    $expirationTime = (Get-Date).AddMinutes([int]$ExpirationMinutes)
    $expirationTimeStr = $expirationTime.ToString("dddd, MMMM d, yyyy h:mm tt")
    
    # Replace template variables
    $template = $template.Replace('${UserDisplayName}', $UserDisplayName)
    $template = $template.Replace('${TemporaryAccessPass}', $TemporaryAccessPass)
    $template = $template.Replace('${SupportEmailAddress}', $SupportEmailAddress)
    $template = $template.Replace('${ExpirationTime}', $expirationTimeStr)
    $template = $template.Replace('${ExpirationMinutes}', $ExpirationMinutes)
    
    return $template
}

function Get-DefaultEmailTemplate {
    # Default email template
    $template = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Your Temporary Access Credentials</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
        }
        .header {
            background-color: #0078d4;
            color: white;
            padding: 20px;
            border-radius: 5px 5px 0 0;
        }
        .content {
            background-color: #f9f9f9;
            padding: 20px;
            border-radius: 0 0 5px 5px;
            border: 1px solid #ddd;
        }
        .credentials {
            background-color: #eff6fc;
            border: 1px solid #0078d4;
            border-radius: 5px;
            padding: 15px;
            margin: 20px 0;
        }
        .credentials code {
            font-family: Consolas, monospace;
            font-size: 18px;
            font-weight: bold;
            color: #0078d4;
            display: block;
            text-align: center;
            padding: 10px;
        }
        .footer {
            margin-top: 20px;
            font-size: 12px;
            color: #666;
        }
        .button {
            display: inline-block;
            background-color: #0078d4;
            color: white;
            text-decoration: none;
            padding: 12px 20px;
            border-radius: 4px;
            font-weight: bold;
            margin: 20px 0;
        }
        .warning {
            background-color: #fff4ce;
            border-left: 4px solid #ffd335;
            padding: 10px;
            margin: 20px 0;
        }
        .expiration {
            background-color: #e1dfdd;
            padding: 10px;
            border-radius: 4px;
            margin-top: 15px;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="header">
        <h2>Welcome to Your New Account</h2>
    </div>
    <div class="content">
        <p>Hello ${UserDisplayName},</p>
        
        <p>Your new account has been created and is ready for you to access. Below is your temporary access credential:</p>
        
        <div class="credentials">
            <p><strong>Temporary Access Pass:</strong></p>
            <code>${TemporaryAccessPass}</code>
            <div class="expiration">
                <strong>Expires:</strong> ${ExpirationTime} (valid for ${ExpirationMinutes} minutes)
            </div>
        </div>
        
        <div class="warning">
            <p><strong>Important:</strong> This Temporary Access Pass can only be used once. After using it, you will need to set up your own authentication methods.</p>
        </div>
        
        <p>Please follow these steps to sign in:</p>
        <ol>
            <li>Go to <a href="https://aka.ms/mysecurityinfo">https://aka.ms/mysecurityinfo</a></li>
            <li>Enter your email address when prompted</li>
            <li>When asked for a password, enter the Temporary Access Pass provided above</li>
            <li>Follow the prompts to set up your authentication methods</li>
        </ol>
        
        <a href="https://aka.ms/mysecurityinfo" class="button">Sign In Now</a>
        
        <p>If you experience any issues or need a new Temporary Access Pass, please contact our support team at <a href="mailto:${SupportEmailAddress}">${SupportEmailAddress}</a>.</p>
        
        <p>Thank you,<br>
        IT Support Team</p>
        
        <div class="footer">
            <p>This is an automated message. Please do not reply to this email.</p>
            <p>This email was sent encrypted for security purposes.</p>
        </div>
    </div>
</body>
</html>
"@

    return $template
}

function Send-EmailWithOME {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FromAddress,
        
        [Parameter(Mandatory = $true)]
        [string]$ToAddress,
        
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        
        [Parameter(Mandatory = $true)]
        [string]$HtmlBody,
        
        [Parameter(Mandatory = $true)]
        [string]$UserDisplayName
    )
    
    try {
        # Validate that we're using a dedicated service account
        if (-not ($FromAddress -like "tap-delivery@*")) {
            Write-Log "WARNING: Using a non-standard service account ($FromAddress). Recommended format is tap-delivery@yourdomain.com" -Level "WARNING"
        }
        
        # Check if OME is available in the tenant
        $omeCapabilities = $null
        try {
            # Try to check OME capabilities
            $omeCapabilities = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/admin/serviceAnnouncement/healthOverviews/Exchange" -ErrorAction SilentlyContinue
        } catch {
            Write-Log "Unable to check OME capabilities: $_" -Level "WARNING"
        }
        
        $useOme = $true
        if ($null -eq $omeCapabilities -or $omeCapabilities.service -ne "Exchange" -or $omeCapabilities.status -ne "ServiceOperational") {
            Write-Log "Office 365 Message Encryption may not be available or could not be verified. Falling back to plain text email." -Level "WARNING"
            $useOme = $false
        }
        
        # Prepare the email message
        $messageParams = @{
            Message = @{
                Subject      = $Subject
                Body         = @{
                    ContentType = "HTML"
                    Content     = $HtmlBody
                }
                ToRecipients = @(
                    @{
                        EmailAddress = @{
                            Address = $ToAddress
                        }
                    }
                )
                # Explicitly set the sender account
                From         = @{
                    EmailAddress = @{
                        Address = $FromAddress
                    }
                }
                Importance   = "High"
            }
        }
        
        # Add encryption properties if OME is available
        if ($useOme) {
            $messageParams.Message.SingleValueExtendedProperties = @(
                @{
                    Id    = "String 0x001F0138"
                    Value = "Encrypt"
                }
            )
        }
        
        # Construct the URI for sending mail
        $sendMailUri = "https://graph.microsoft.com/v1.0/users/$FromAddress/sendMail"
        
        # Send the email
        Invoke-MgGraphRequest -Method POST -Uri $sendMailUri -Body ($messageParams | ConvertTo-Json -Depth 10) -ContentType "application/json" -ErrorAction Stop
        
        if ($useOme) {
            Write-Log "Successfully sent encrypted email to $ToAddress for user $UserDisplayName" -Level "SUCCESS"
        } else {
            Write-Log "Successfully sent email to $ToAddress for user $UserDisplayName (without encryption)" -Level "SUCCESS"
        }
        
        return $true
    } catch {
        $errorMsg = $_
        
        if ($errorMsg -like "*Authorization_RequestDenied*") {
            Write-Log "Access denied when sending email. The app may not have the Mail.Send permission." -Level "ERROR"
            Write-Log "Ensure the app has the Mail.Send permission and it has been admin-consented." -Level "ERROR"
        } elseif ($errorMsg -like "*Request_ResourceNotFound*") {
            Write-Log "Email address '$FromAddress' not found. The mailbox may not exist." -Level "ERROR"
            Write-Log "Create the service account mailbox or specify a valid FromEmailAddress parameter." -Level "ERROR"
        } else {
            Write-Log "Error sending email to $($ToAddress): $errorMsg" -Level "ERROR"
        }
        
        return $false
    }
}

function Process-User {
    param(
        [Parameter(Mandatory = $true)]
        [object]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$FromAddress,
        
        [Parameter(Mandatory = $true)]
        [int]$TapLifetime
    )
    
    $userId = $User.Id
    $displayName = $User.DisplayName
    $upn = $User.UserPrincipalName
    
    try {
        Write-Log "Processing user: $displayName ($upn)" -Level "INFO"
        
        # Get the source email address from extensionAttribute15
        $sourceEmail = Get-SourceEmailFromExtensionAttribute -UserId $userId
        
        if (-not $sourceEmail) {
            Write-Log "Skipping user $displayName ($upn): No valid source email found in extensionAttribute15" -Level "WARNING"
            $script:errorCount++
            return
        }
        
        # Create Temporary Access Pass
        $tap = New-TemporaryAccessPass -UserId $userId -LifetimeInMinutes $TapLifetime -Force:$Force
        
        if (-not $tap) {
            Write-Log "Failed to create Temporary Access Pass for user $displayName ($upn)" -Level "ERROR"
            $script:errorCount++
            return
        }
        
        # Store credentials (for potential logging/reporting)
        $script:tapCredentials[$upn] = @{
            "DisplayName" = $displayName
            "TAP"         = $tap
            "SourceEmail" = $sourceEmail
            "Created"     = Get-Date
            "Expires"     = (Get-Date).AddMinutes($TapLifetime)
        }
        
        # Generate email content
        $emailBody = Get-EmailTemplate -UserDisplayName $displayName -TemporaryAccessPass $tap -SupportEmailAddress $SupportEmail -ExpirationMinutes $TapLifetime
        
        # Send email
        $emailSent = Send-EmailWithOME -FromAddress $FromAddress -ToAddress $sourceEmail -Subject $EmailSubject -HtmlBody $emailBody -UserDisplayName $displayName
        
        if ($emailSent) {
            Write-Log "Successfully processed user $displayName ($upn) - TAP created and email sent to $sourceEmail" -Level "SUCCESS"
            $script:successCount++
        } else {
            Write-Log "TAP created for $displayName ($upn) but failed to send email to $sourceEmail" -Level "ERROR"
            $script:errorCount++
        }
    } catch {
        Write-Log "Error processing user $displayName ($upn): $_" -Level "ERROR"
        $script:errorCount++
    }
}

# Main script execution
try {
    # Clear the screen to start fresh
    Clear-Host
    
    Write-Host "=====================================================" -ForegroundColor Green
    Write-Host "  COMET TAP CREDENTIAL GENERATION AND DELIVERY SCRIPT" -ForegroundColor Green
    Write-Host "=====================================================" -ForegroundColor Green
    Write-Host "This script will create TAPs for users in a security group"
    Write-Host "and send them via email to addresses in extensionAttribute15."
    Write-Host
    
    # Check PowerShell version
    if (-not (Test-PowerShellVersion)) {
        exit 1
    }
    
    # Check and install required modules
    if (-not (Ensure-RequiredModules)) {
        exit 1
    }
    
    Write-Log "Starting TAP credential generation and delivery process" -Level "INFO"
    Write-Log "Log file: $($script:logFile)" -Level "INFO"
    Write-Log "Using least privileged approach with minimum required permissions" -Level "INFO"
    
    # Connect to Microsoft Graph with appropriate authentication
    $connectedTenantId = $null
    if ($ClientSecret) {
        $connectedTenantId = Connect-ToMsGraphWithAuth -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret
    } else {
        $connectedTenantId = Connect-ToMsGraphWithAuth -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
    }
    
    if (-not $connectedTenantId) {
        Write-Log "Failed to connect to Microsoft Graph. Exiting." -Level "ERROR"
        exit 1
    }
    
    # If no FromEmailAddress is provided, get it from the tenant
    if (-not $FromEmailAddress) {
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
                Write-Log "Unable to find a verified domain in the tenant. Using default address." -Level "WARNING"
                $FromEmailAddress = "tap-delivery@unknown-domain.com"
            } else {
                # Use a dedicated service account name that clearly indicates its purpose
                $FromEmailAddress = "tap-delivery@$($primaryDomain.Id)"
            }
        } catch {
            Write-Log "Error retrieving domains. Using default address: $_" -Level "WARNING"
            $FromEmailAddress = "tap-delivery@unknown-domain.com"
        }
    }
    
    # Verify the service account exists
    Write-Log "Using sender email address: $FromEmailAddress" -Level "INFO"
    try {
        $serviceAccount = Get-MgUser -Filter "mail eq '$FromEmailAddress' or userPrincipalName eq '$FromEmailAddress'" -ErrorAction SilentlyContinue
        
        if ($null -eq $serviceAccount) {
            Write-Log "WARNING: The service account '$FromEmailAddress' does not appear to exist in the tenant." -Level "WARNING"
            Write-Log "Create this account or specify a different FromEmailAddress parameter." -Level "WARNING"
            
            $continue = Read-Host "The service account doesn't seem to exist. Continue anyway? (Y/N)"
            if ($continue -ne "Y" -and $continue -ne "y") {
                Write-Log "Operation cancelled by user. Please create the service account and try again." -Level "WARNING"
                exit 0
            }
        } else {
            Write-Log "Verified service account exists: $FromEmailAddress" -Level "SUCCESS"
        }
    } catch {
        Write-Log "Could not verify if service account exists: $_" -Level "WARNING"
    }
    
    # Get members of the specified security group
    $members = Get-SecurityGroupMembers -GroupId $SecurityGroupId
    
    if ($members.Count -eq 0) {
        Write-Log "No members found in security group $SecurityGroupId or security group does not exist. Exiting." -Level "ERROR"
        exit 1
    }
    
    $script:userCount = $members.Count
    Write-Log "Found $($members.Count) users in security group to process" -Level "INFO"
    
    # Process each user
    $counter = 0
    $batchSize = [Math]::Min($members.Count, 10)  # Process in batches of 10 for progress reporting
    $batchCount = [Math]::Ceiling($members.Count / $batchSize)
    
    for ($batchIndex = 0; $batchIndex -lt $batchCount; $batchIndex++) {
        $batchStart = $batchIndex * $batchSize
        $batchEnd = [Math]::Min($batchStart + $batchSize - 1, $members.Count - 1)
        
        Write-Progress -Activity "Processing Users" -Status "Batch $($batchIndex+1) of $batchCount" -PercentComplete (($batchIndex / $batchCount) * 100)
        
        # Process users in the current batch
        for ($i = $batchStart; $i -le $batchEnd; $i++) {
            $member = $members[$i]
            $counter++
            
            Write-Progress -Id 1 -ParentId 0 -Activity "Processing User" -Status "$counter of $($members.Count): $($member.DisplayName)" -PercentComplete (($counter / $members.Count) * 100)
            
            Process-User -User $member -FromAddress $FromEmailAddress -TapLifetime $TapLifetimeMinutes
            
            # Small delay to prevent throttling
            Start-Sleep -Milliseconds 500
        }
    }
    
    # Summary
    Write-Progress -Activity "Processing Users" -Completed
    Write-Log "-------------------------------------------" -Level "INFO"
    Write-Log "TAP Credential Delivery Summary:" -Level "INFO"
    Write-Log "Total users processed: $script:userCount" -Level "INFO"
    Write-Log "Successful: $script:successCount" -Level "SUCCESS"
    Write-Log "Failed: $script:errorCount" -Level "ERROR"
    Write-Log "Log file: $script:logFile" -Level "INFO"
    Write-Log "Service account used: $FromEmailAddress" -Level "INFO"
    Write-Log "-------------------------------------------" -Level "INFO"
    
    # Export results to CSV if there are successful TAPs created
    if ($script:tapCredentials.Count -gt 0) {
        $csvFilePath = Join-Path -Path $script:logFolder -ChildPath "TapCredentials_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        $csvData = $script:tapCredentials.GetEnumerator() | ForEach-Object {
            [PSCustomObject]@{
                UserPrincipalName   = $_.Key
                DisplayName         = $_.Value.DisplayName
                SourceEmail         = $_.Value.SourceEmail
                TemporaryAccessPass = $_.Value.TAP
                Created             = $_.Value.Created
                Expires             = $_.Value.Expires
            }
        }
        
        $csvData | Export-Csv -Path $csvFilePath -NoTypeInformation
        Write-Log "Credential information exported to: $csvFilePath" -Level "SUCCESS"
        Write-Log "WARNING: This file contains sensitive information. Delete it after use!" -Level "WARNING"
    }
    
    Write-Host "`nTAP credential generation and delivery completed successfully!" -ForegroundColor Green
    Write-Host "See the log file for details: $($script:logFile)" -ForegroundColor Cyan
    Write-Host "`nPermissions used (least privilege):" -ForegroundColor Yellow
    Write-Host "- User.Read.All (Reading user properties including extensionAttribute15)" -ForegroundColor White
    Write-Host "- UserAuthenticationMethod.ReadWrite.All (Managing TAPs)" -ForegroundColor White  
    Write-Host "- Mail.Send (From dedicated service account: $FromEmailAddress)" -ForegroundColor White
} catch {
    Write-Log "An unhandled error occurred: $_" -Level "ERROR"
    
    # More detailed error information
    if ($_.Exception.Response) {
        try {
            $errorResponseStream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponseStream)
            $errorResponseBody = $reader.ReadToEnd()
            Write-Log "Error details: $errorResponseBody" -Level "ERROR"
        } catch {
            Write-Log "Could not retrieve detailed error information." -Level "ERROR"
        }
    }
    
    exit 1
} finally {
    # Disconnect from Graph API
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}