<#
.SYNOPSIS
    Performs a comprehensive health check of a Microsoft 365 tenant.
.DESCRIPTION
    This script checks tenant security settings, license usage, service health,
    and other critical areas of a Microsoft 365 tenant using Microsoft Graph.
.PARAMETER TenantId
    The tenant ID to perform the health check against.
.NOTES
    File Name      : Invoke-M365TenantHealthCheck.ps1
    Author         : MSP Administrator
    Prerequisite   : Microsoft Graph PowerShell SDK
#>

function Connect-ToMicrosoftGraph {
    [CmdletBinding()]
    param()
    
    try {
        # Connect to Microsoft Graph with required scopes
        Connect-MgGraph -Scopes "Directory.Read.All", "Organization.Read.All", "Reports.Read.All", "SecurityEvents.Read.All"
        $orgInfo = Get-MgOrganization
        Write-Host "Successfully connected to tenant: $($orgInfo.DisplayName)" -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit 1
    }
}

function Get-LicenseUtilization {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "Checking license utilization..." -ForegroundColor Cyan
        
        $subscribedSkus = Get-MgSubscribedSku
        
        foreach ($sku in $subscribedSkus) {
            $usedLicenses = $sku.ConsumedUnits
            $totalLicenses = $sku.PrepaidUnits.Enabled
            $percentUsed = [math]::Round(($usedLicenses / $totalLicenses) * 100, 2)
            
            $skuPartNumber = $sku.SkuPartNumber
            
            # Color code based on utilization percentage
            $color = "Green"
            if ($percentUsed -gt 90) { $color = "Red" }
            elseif ($percentUsed -gt 80) { $color = "Yellow" }
            
            Write-Host "License: $skuPartNumber - Used: $usedLicenses/$totalLicenses ($percentUsed%)" -ForegroundColor $color
        }
    } catch {
        Write-Error "Failed to retrieve license information: $_"
    }
}

function Get-SecureScoreDetails {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "Checking Microsoft Secure Score..." -ForegroundColor Cyan
        
        # Get secure score using Microsoft Graph
        $secureScore = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/security/secureScores?`$top=1"
        
        if ($secureScore.value.Count -gt 0) {
            $currentScore = $secureScore.value[0].currentScore
            $maxScore = $secureScore.value[0].maxScore
            $percentage = [math]::Round(($currentScore / $maxScore) * 100, 2)
            
            # Color code based on secure score percentage
            $color = "Green"
            if ($percentage -lt 50) { $color = "Red" }
            elseif ($percentage -lt 70) { $color = "Yellow" }
            
            Write-Host "Current Secure Score: $currentScore out of $maxScore ($percentage%)" -ForegroundColor $color
            
            # Get improvement actions
            $improvementActions = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles"
            
            Write-Host "`nTop 5 Improvement Actions:" -ForegroundColor Cyan
            $improvementActions.value | 
                Where-Object { $_.implementationStatus -ne "Implemented" } | 
                Sort-Object -Property { $_.contributionToSecureScore } -Descending | 
                Select-Object -First 5 | 
                ForEach-Object {
                    Write-Host "- $($_.title): +$($_.contributionToSecureScore) points" -ForegroundColor Yellow
                }
        } else {
            Write-Host "No Secure Score data available." -ForegroundColor Yellow
        }
    } catch {
        Write-Error "Failed to retrieve secure score information: $_"
    }
}

function Get-ConditionalAccessPolicySummary {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "Checking Conditional Access Policies..." -ForegroundColor Cyan
        
        $policies = Get-MgIdentityConditionalAccessPolicy
        
        $enabledPolicies = $policies | Where-Object { $_.State -eq "enabled" }
        $reportOnlyPolicies = $policies | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }
        $disabledPolicies = $policies | Where-Object { $_.State -eq "disabled" }
        
        Write-Host "Total Policies: $($policies.Count)" -ForegroundColor White
        Write-Host "- Enabled: $($enabledPolicies.Count)" -ForegroundColor Green
        Write-Host "- Report Only: $($reportOnlyPolicies.Count)" -ForegroundColor Yellow
        Write-Host "- Disabled: $($disabledPolicies.Count)" -ForegroundColor Red
        
        # Check for baseline policies
        $mfaPolicy = $enabledPolicies | Where-Object { $_.DisplayName -like "*MFA*" -or $_.DisplayName -like "*Multi-Factor*" }
        if (-not $mfaPolicy) {
            Write-Host "WARNING: No enabled MFA policies detected!" -ForegroundColor Red
        }
    } catch {
        Write-Error "Failed to retrieve Conditional Access Policies: $_"
    }
}

# Main execution
try {
    # Verify required modules
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Error "Required module Microsoft.Graph not installed. Please install with: Install-Module Microsoft.Graph -Scope CurrentUser"
        exit 1
    }
    
    # Connect to Microsoft Graph
    Connect-ToMicrosoftGraph
    
    # Run the health checks
    Get-LicenseUtilization
    Get-SecureScoreDetails
    Get-ConditionalAccessPolicySummary
    
    # Additional checks can be added here
    
    Write-Host "`nTenant health check completed successfully." -ForegroundColor Green
} catch {
    Write-Error "An error occurred during the tenant health check: $_"
} finally {
    # Disconnect from Microsoft Graph
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}