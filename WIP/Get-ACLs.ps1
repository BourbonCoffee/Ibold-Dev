<#
.SYNOPSIS
    Enumerate directories (and optionally files) in a given path, retrieve ACLs,
    separate them into success/failure, and also record any "non-default" permissions
    (blocked inheritance or explicit ACEs) into a separate CSV.

.PARAMETER Path
    The root path to begin enumerating (can be UNC, e.g., \\Server\Share).

.PARAMETER SuccessCsv
    CSV file for successful ACL entries.

.PARAMETER FailureCsv
    CSV file for failures (where ACL retrieval fails).

.PARAMETER NonDefaultCsv
    CSV file for items that have blocked inheritance or any explicit ACEs.

.PARAMETER IncludeFiles
    Switch to also enumerate files (not just directories).

.PARAMETER LogFile
    Optional. If provided, logs are appended to this file in addition to the console.

.EXAMPLE
    .\Get-ACLs-NonDefault.ps1 -Path "M:\" `
        -SuccessCsv "C:\Temp\ACLs_Success.csv" `
        -FailureCsv "C:\Temp\ACLs_Failure.csv" `
        -NonDefaultCsv "C:\Temp\ACLs_NonDefault.csv" `
        -IncludeFiles `
        -LogFile "C:\Temp\ACLs.log"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$SuccessCsv,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$FailureCsv,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$NonDefaultCsv,

    [switch]$IncludeFiles,

    [Parameter(Mandatory = $false)]
    [string]$LogFile
)

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "$timestamp [$Level] $Message"

    Write-Host $logEntry

    if ($LogFile) {
        Add-Content -Path $LogFile -Value $logEntry
    }
}

Write-Log "Starting ACL enumeration on '$Path'..."

try {
    if ($IncludeFiles) {
        Write-Log "Including both directories AND files."
        # Enumerate directories & files
        $itemList = Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue
    } else {
        Write-Log "Enumerating directories only."
        # Enumerate directories only
        $itemList = Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction SilentlyContinue
    }
} catch {
    Write-Log "Failed to retrieve items from '$Path': $($_.Exception.Message)" "ERROR"
    return
}

$total = $itemList.Count
Write-Log "Found $total items in '$Path'."

if ($total -eq 0) {
    Write-Log "No items found. Exiting."
    return
}

# Prepare three collections:
#   1) Successes (ALL successful ACL retrievals)
#   2) Failures (ACL retrieval failed)
#   3) NonDefault (subset of successes with either blocked inheritance or explicit ACEs)
$successes = New-Object System.Collections.Generic.List[System.Object]
$failures = New-Object System.Collections.Generic.List[System.Object]
$nondefault = New-Object System.Collections.Generic.List[System.Object]

$counter = 0

Write-Log "Collecting ACLs..."

foreach ($item in $itemList) {
    $counter++
    $percentComplete = [int](($counter / $total) * 100)

    Write-Progress `
        -Activity "Collecting ACLs" `
        -Status "Processing $($item.FullName) ($counter of $total)" `
        -PercentComplete $percentComplete

    try {
        # Get the ACL (Note: "Stop" so we can catch any exception)
        $acl = Get-Acl -Path $item.FullName -ErrorAction Stop

        # For each Access Control Entry in this item's ACL
        foreach ($ace in $acl.Access) {
            # Build a record
            $successObj = [pscustomobject]@{
                Path               = $item.FullName
                IdentityReference  = $ace.IdentityReference.ToString()
                IsInherited        = $ace.IsInherited
                InheritanceBlocked = $acl.AreAccessRulesProtected
            }
            $successes.Add($successObj) | Out-Null

            # Check if "non-default"
            #   i.e. the ACL is blocking inheritance (AreAccessRulesProtected = $true)
            #   OR this ACE is explicitly defined (IsInherited = $false)
            if ($acl.AreAccessRulesProtected -eq $true -or $ace.IsInherited -eq $false) {
                $nondefault.Add($successObj) | Out-Null
            }
        }
    } catch {
        # If ACL retrieval fails, store in the failures list
        $errorMsg = $_.Exception.Message
        Write-Log "Cannot get ACL for '$($item.FullName)': $errorMsg" "WARN"

        $failureObj = [pscustomobject]@{
            Path  = $item.FullName
            Error = $errorMsg
        }
        $failures.Add($failureObj) | Out-Null
    }
}

Write-Log "Finished collecting ACLs. Writing CSV outputs..."

try {
    # --- Success CSV (all retrieved ACEs) ---
    $successes | Sort-Object Path, IdentityReference -Unique |
        Export-Csv -Path $SuccessCsv -NoTypeInformation -Force
    Write-Log "Success CSV saved to '$SuccessCsv'."

    # --- Failure CSV (couldn't retrieve ACL) ---
    $failures | Sort-Object Path -Unique |
        Export-Csv -Path $FailureCsv -NoTypeInformation -Force
    Write-Log "Failure CSV saved to '$FailureCsv'."

    # --- NonDefault CSV (subset of successes) ---
    $nondefault | Sort-Object Path, IdentityReference -Unique |
        Export-Csv -Path $NonDefaultCsv -NoTypeInformation -Force
    Write-Log "Non-default CSV saved to '$NonDefaultCsv'."
} catch {
    Write-Log "Failed to write CSV files: $($_.Exception.Message)" "ERROR"
    return
}

Write-Log "Done!"
