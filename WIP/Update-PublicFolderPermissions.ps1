#############################################################################################################
#.SYNOPSIS
#   Updates client permissions of several users to a public folder
#
#.DESCRIPTION
#   Updates the client permissions of a public folder (and its children if -recurse
#	is provided) clearing the permissions a set of users have on the folder and setting
#	the provided access rights
#
#	Copyright (c) 2014 Microsoft Corporation. All rights reserved.
#
#	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
#	OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
#.PARAMETER  IncludeFolders
#   Identities of the Public Folders that will be updated
#
#.PARAMETER  Users
#   List of users whose current access rights to the folder will be overriten
#
#.PARAMETER  AccessRights
#   List of permissions to assign to the users
#
#.PARAMETER  Recurse
#   If provided the permission changes will also be applied to the children of the folders.
#
#.PARAMETER  ExcludeFolderEntryIds
#   List of EntryIds of the folders that should be ignored from the update. Notice however
#   that if you use the Recurse option the children of these folders won't be ignored unless
#   their EntryIds are also provided in this list.
#
#.PARAMETER  SkipCurrentAccessCheck
#   If provided the right access updates will be performed in the folder regardless of whether
#   the current folder has the same permissions already applied.
#
#.PARAMETER  Confirm
#   If this switch parameter is set to $false all operations on the public folder will be 
#   performed without requesting confirmation from the user.
#
#.PARAMETER  WhatIf
#   If this switch parameter is present the operations on the public folder will not be 
#   performed but information on what task would be performed are printed to the console.
#
#.PARAMETER  ProgressLogFile
#   File to log EntryIds of folders that were successfully updated. The content of this file may
#   become handy to save time if the previous execution of the script was aborted and you want to restart
#   from the point the script stopped. To do this simply get the contents of the file (get-content) and 
#   provide the data to the ExcludeFolderEntryIds parameter.
#
#   The default path is UpdatePublicFolderPermission.[yyyyMMdd_HHmm].log where the portion in square brackets
#   gets replaced with the current date and time at the moment of execution of the script.
#
#.EXAMPLE
#    .\Update-PublicFolderPermissions.ps1 -IncludeFolders "\MyFolder" -AccessRights "Owner" -Users "John", "Administrator" -Recurse -Confirm:$false
#
#	This command replaces the current client permissions for users "John" and "Administrator" on the "\MyFolder"
#   Public Folder and all its children. The users will be granted "Owner" access rights. These actions will be 
#	performed without requesting confirmation to the user.
#
#.EXAMPLE
#    $foldersProcessed = get-content .\UpdatePublicFolderPermission.20141031_1820.log
#    .\Update-PublicFolderPermissions.ps1 -IncludeFolders "\MyFolder" -AccessRights "Owner" -Users "John", "Administrator" -Recurse -ExcludeFolderEntryIds $foldersProcessed -Confirm:$false
#
#	These commands replace the current client permissions for users "John" and "Administrator" on the "\MyFolder"
#   Public Folder and all its children but skips those folders that were completd in the execution of Oct 30th 2014 at 6:20 pm.
#   The users will be granted "Owner" access rights. These actions will be performed without requesting confirmation to the user.
#############################################################################################################

param (
    [Parameter(Mandatory=$True)]
    [string[]]$IncludeFolders,
    [Parameter(Mandatory=$True)]
    [string[]]$Users,
    [Parameter(Mandatory=$True)]
    [string[]]$AccessRights,
    [switch]$Recurse,
    [string[]]$ExcludeFolderEntryIds = @(),
    [switch]$SkipCurrentAccessCheck,
    [string]$ProgressLogFile = ".\UpdatePublicFolderPermission.$((Get-Date).ToString('yyyyMMdd_HHmm')).log",
    [switch]$confirm,
    [switch]$whatIf
)

#############################################################################################################
#   Returns the list of public folders to process ignoring duplicates and folders in the exclude list
#############################################################################################################
function FindFoldersToUpdate([string[]]$includeFolders, [bool]$recurseOnFolders, [string[]]$excludeFolderEntryIds)
{
    $folderToSkip = new-object 'System.Collections.Generic.HashSet[string]' -ArgumentList @(,$excludeFolderEntryIds)
    $currentIncludeFolder=0;
    foreach($includeFolder in $includeFolders)
    {
        $progress = 100 * $currentIncludeFolder / $includeFolders.Count;
        Write-Progress -Activity "Retrieving folders to update" -Status $includeFolder -PercentComplete $progress

        $foldersFound = Get-PublicFolder -Recurse:$recurseOnFolders $includeFolder

        if ($foldersFound -eq $null)
        {
            continue;
        }

        foreach($foundFolder in $foldersFound)
        {
            if ($foundFolder -eq $null)
            {
                continue;
            }

            if ($folderToSkip -notContains $foundFolder.EntryId)
            {
                #Return found folder
                $foundFolder;
            }

            $folderToSkip.Add($foundFolder.EntryId) > $null;
        }

        $currentIncludeFolder++;
    }
}

#############################################################################################################
#   Returns the Identity of the users that need processing.
#############################################################################################################
function GetUserIdentities([string[]]$Users)
{
    $userIdentities = new-object 'System.Collections.Generic.HashSet[object]' 
    $currentUserNumber=0;
    foreach($user in $Users)
    {
        $progress = 100 * $currentUserNumber / $Users.Count;
        Write-Progress -Activity "Retrieving users" -Status $user -PercentComplete $progress
        $id = (Get-Recipient $user).Identity

        if ($id -ne $null)
        {
            $userIdentities.Add($id) > $null
        }

        $currentUserNumber++;
    }

    $userIdentities
}

#############################################################################################################
#   Returns whether all the elements of a collection are present in a reference collection.
#############################################################################################################
function CollectionContains($referenceCollection, $otherCollection)
{
    foreach($item in $otherCollection)
    {
        if ($referenceCollection -notcontains $item)
        {
            return $false
        }
    }

    return $true
}

#############################################################################################################
#   Verifies whether there is a mismatch between the desired and found permissions.
#############################################################################################################
function IsUpdateRequired ($currentAccessRights, $desiredAccessRights)
{
    $allDesiredPermissionsWhereFound = CollectionContains $currentAccessRights $desiredAccessRights
    $allFoundPermissionsAreDesired = CollectionContains $desiredAccessRights $currentAccessRights

    return -not ($allDesiredPermissionsWhereFound -and $allFoundPermissionsAreDesired)
}

#############################################################################################################
#   Gets the list of users whose access right to a folder don't match the desired ones.
#############################################################################################################
function GetUsersToUpdate($currentFolder, [Array]$usersToUpdate, [string[]]$accessRights)
{
    Write-Progress -Id 1 -Activity "Querying current permissions" -Status "Processing";

    $existingPermissions = [Array](Get-PublicFolderClientPermission $currentFolder.Identity);
    $existingPermissionsPerUser = @{}

    $permissionCount = 0;
    foreach($permission in $existingPermissions)
    {
        $progress = 100 * $permissionCount / $existingPermissions.Count;
        Write-Progress -Id 1 -Activity "Processing current permissions" -PercentComplete $progress -Status "Processing";

        $adIdentity = $permission.User.ADRecipient.Identity;

        if ($adIdentity -ne $null)
        {
            $existingPermissionsPerUser[$adIdentity] = $permission;
        }
    }

    $permissionCount = 0;
    foreach($user in $usersToUpdate)
    {
        $progress = 100 * $permissionCount / $usersToUpdate.Count;
        Write-Progress -Id 1 -Activity "Comparing permissions" -PercentComplete $progress -Status "Processing";

        if (-not $existingPermissionsPerUser.ContainsKey($user))
        {
            $user;
        }
        else
        {
            if (IsUpdateRequired $existingPermissionsPerUser[$user].AccessRights $AccessRights)
            {
                $user;
            }
        }

        $permissionCount++;
    }
}

#############################################################################################################
#   Script logic.
#############################################################################################################

$foldersToUpdate=[Array](FindFoldersToUpdate $IncludeFolders $Recurse $ExcludeFolderEntryIds);
$usersToUpdate=[Array](GetUserIdentities $Users)

$foldersProcessed=0;
foreach($currentFolder in $foldersToUpdate)
{
    $percentFoldersProcessed = 100 * $foldersProcessed/($foldersToUpdate.Count);
    Write-Progress -Activity "Processing folders" -Status $currentFolder.Identity -PercentComplete $percentFoldersProcessed

    $usersToUpdateForFolder = @()
    if (-not $SkipCurrentAccessCheck)
    {
        $usersToUpdateForFolder =  [Array](GetUsersToUpdate $currentFolder $usersToUpdate $AccessRights)
    }
    else
    {
        $usersToUpdateForFolder = $usersToUpdate;
    }

    $folderOperationFailed=$false;
    $usersProcessed=0;

    if (($usersToUpdateForFolder -eq $null) -or ($usersToUpdateForFolder.Count -eq 0))
    {
        Write-Warning "Couldn't find any changes to perform for folder $($currentFolder.Identity)"
        continue;
    }

    foreach($user in $usersToUpdateForFolder)
    {
        $percentUsersProcessed = 100 * $usersProcessed/($usersToUpdateForFolder.Count)

        Write-Progress -Id 1 -Activity "Processing User" -Status $user -CurrentOperation "Removing exisitng permission" -PercentComplete $percentUsersProcessed
        Remove-PublicFolderClientPermission -User $user $currentFolder.Identity -ErrorAction SilentlyContinue -Confirm:$confirm -WhatIf:$whatIf

        Write-Progress -Id 1 -Activity "Processing User" -Status $user -CurrentOperation "Adding permission" -PercentComplete $percentUsersProcessed

        try
        {
            Add-PublicFolderClientPermission -AccessRights $accessRights -User $user $currentFolder.Identity -ErrorAction Stop -Confirm:$confirm -WhatIf:$whatIf
        }
        catch
        {
            Write-Error $_
            $folderOperationFailed=$true;
        }

        $usersProcessed++;
    }

    if (-not $folderOperationFailed)
    {
        Add-Content $ProgressLogFile "$($currentFolder.EntryId)`n" -Confirm:$confirm -WhatIf:$whatIf
    }

    $foldersProcessed++;
}


# SIG # Begin signature block
# MIIjhgYJKoZIhvcNAQcCoIIjdzCCI3MCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDtsliZT3G6Lqhn
# +nxFJGp2TAyMeBSLmO1UHzeTg7v3+KCCDYEwggX/MIID56ADAgECAhMzAAAB32vw
# LpKnSrTQAAAAAAHfMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjAxMjE1MjEzMTQ1WhcNMjExMjAyMjEzMTQ1WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC2uxlZEACjqfHkuFyoCwfL25ofI9DZWKt4wEj3JBQ48GPt1UsDv834CcoUUPMn
# s/6CtPoaQ4Thy/kbOOg/zJAnrJeiMQqRe2Lsdb/NSI2gXXX9lad1/yPUDOXo4GNw
# PjXq1JZi+HZV91bUr6ZjzePj1g+bepsqd/HC1XScj0fT3aAxLRykJSzExEBmU9eS
# yuOwUuq+CriudQtWGMdJU650v/KmzfM46Y6lo/MCnnpvz3zEL7PMdUdwqj/nYhGG
# 3UVILxX7tAdMbz7LN+6WOIpT1A41rwaoOVnv+8Ua94HwhjZmu1S73yeV7RZZNxoh
# EegJi9YYssXa7UZUUkCCA+KnAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUOPbML8IdkNGtCfMmVPtvI6VZ8+Mw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDYzMDA5MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAnnqH
# tDyYUFaVAkvAK0eqq6nhoL95SZQu3RnpZ7tdQ89QR3++7A+4hrr7V4xxmkB5BObS
# 0YK+MALE02atjwWgPdpYQ68WdLGroJZHkbZdgERG+7tETFl3aKF4KpoSaGOskZXp
# TPnCaMo2PXoAMVMGpsQEQswimZq3IQ3nRQfBlJ0PoMMcN/+Pks8ZTL1BoPYsJpok
# t6cql59q6CypZYIwgyJ892HpttybHKg1ZtQLUlSXccRMlugPgEcNZJagPEgPYni4
# b11snjRAgf0dyQ0zI9aLXqTxWUU5pCIFiPT0b2wsxzRqCtyGqpkGM8P9GazO8eao
# mVItCYBcJSByBx/pS0cSYwBBHAZxJODUqxSXoSGDvmTfqUJXntnWkL4okok1FiCD
# Z4jpyXOQunb6egIXvkgQ7jb2uO26Ow0m8RwleDvhOMrnHsupiOPbozKroSa6paFt
# VSh89abUSooR8QdZciemmoFhcWkEwFg4spzvYNP4nIs193261WyTaRMZoceGun7G
# CT2Rl653uUj+F+g94c63AhzSq4khdL4HlFIP2ePv29smfUnHtGq6yYFDLnT0q/Y+
# Di3jwloF8EWkkHRtSuXlFUbTmwr/lDDgbpZiKhLS7CBTDj32I0L5i532+uHczw82
# oZDmYmYmIUSMbZOgS65h797rj5JJ6OkeEUJoAVwwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVWzCCFVcCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAd9r8C6Sp0q00AAAAAAB3zAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgTwb+BQf8
# 1nWr35ixUA2tKzO47rZDZFJ52e8r2vdddzowQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQBhraRlMECMS++k8fr8u9wYqyZd8d1SYnuxZuyLanWe
# XhwzW4rHRmz83T8866DEyIl2H0GjYjp9cg828TGkaOXKLVqoVsA1W2lNrWl8HmsN
# KBMRc/uL6mWmORP0s+WoO0ihiC4XGnvxrdp4yBb7q1694kDX7X+D8pRiJFBhHVrN
# ZkDBqFvhSjKF1tr8pTwyVhWdOdhoUxWgB6YBYkM7L1l5BvrCvTqPDueJTv33IL/i
# bcRRGMVW4BMqRuNEsa/yZ28t8jdham7oPqzSSj2KczKNcroc9RVeWP2p9y7D70mI
# hZzf3344KzORpcnAnsuHSPqQhLIAypBlvZR3tQ5dDyp6oYIS5TCCEuEGCisGAQQB
# gjcDAwExghLRMIISzQYJKoZIhvcNAQcCoIISvjCCEroCAQMxDzANBglghkgBZQME
# AgEFADCCAVEGCyqGSIb3DQEJEAEEoIIBQASCATwwggE4AgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEIF6Lu9z/In55fQtXWVTB7zzN4dDYNjkWiO10VuM/
# yK+GAgZg+ZxHt/4YEzIwMjEwODA1MDIxOTE1LjM5NVowBIACAfSggdCkgc0wgcox
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1p
# Y3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlvbnMxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
# RVNOOjNFN0EtRTM1OS1BMjVEMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloIIOPDCCBPEwggPZoAMCAQICEzMAAAFSMEtdiazmcEcAAAAAAVIw
# DQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcN
# MjAxMTEyMTgyNjA1WhcNMjIwMjExMTgyNjA1WjCByjELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2Eg
# T3BlcmF0aW9uczEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046M0U3QS1FMzU5LUEy
# NUQxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCuzG6EiZh0taCSbswMiupMTYnbboFz
# jj1DuDbbvT0RXKBCVl/umA+Uy214DmHiFhkeuRdlLB0ya5S9um5aKr7lBBqZzvtK
# gGNgCRbDTG9Yu6kzDzPTzQRulVIvoWVy0gITnEyoJ1O3m5IPpsLBNQCdXsh+3TZF
# 73JAcub21bnxm/4sxe4zTdbdttBrqX8/JJF2VEnAP+MBvF2UQSo6XUAaTKC/HPDP
# Cce/IsNoAxxLDI1wHhIlqjRBnt4HM5HcKHrZrvH+vHnihikdlEzh3fjQFowk1fG7
# PVhmO60O5vVdqA+H9314hHENQI0cbo+SkSi8SSJSLNixgj0eWePTh7pbAgMBAAGj
# ggEbMIIBFzAdBgNVHQ4EFgQUhN2u2qwj1l2c2h/kULDuBRJsexQwHwYDVR0jBBgw
# FoAU1WM6XIoxkPNDe3xGG8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
# L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljVGltU3RhUENB
# XzIwMTAtMDctMDEuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0
# cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNUaW1TdGFQQ0FfMjAx
# MC0wNy0wMS5jcnQwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDCDAN
# BgkqhkiG9w0BAQsFAAOCAQEAVcUncfFqSazQbDEXf3d10/upiWQU5HdTbwG9v9be
# VIDaG4oELyIcNE6e6CbOBMlPU+smpYYcnK3jucNqChwquLmxdi2iPy4iQ6vjAdBp
# 9+VFWlrBqUsNXZzjCpgMCZj6bu8Xq0Nndl4WyBbI0Jku68vUNG4wsMdKP3dz+1Mz
# k9SUma3j7HyNA559do9nhKmoZMn5dtf03QvxlaEwMAaPk9xuUv9BN8cNvFnpWk4m
# LERQW6tA3rXK0soEISKTYG7Ose7oMXZDYPWxf9oFhYKzZw/SwnhdBoj2S5eyYE3A
# uF/ZXzR3hdp3/XGzZeOdERfFy1rC7ZBwhDIajeFMi53GnzCCBnEwggRZoAMCAQIC
# CmEJgSoAAAAAAAIwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRp
# ZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTEwMDcwMTIxMzY1NVoXDTI1MDcwMTIx
# NDY1NVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCpHQ28dxGKOiDs/BOX9fp/aZRrdFQQ1aUKAIKF
# ++18aEssX8XD5WHCdrc+Zitb8BVTJwQxH0EbGpUdzgkTjnxhMFmxMEQP8WCIhFRD
# DNdNuDgIs0Ldk6zWczBXJoKjRQ3Q6vVHgc2/JGAyWGBG8lhHhjKEHnRhZ5FfgVSx
# z5NMksHEpl3RYRNuKMYa+YaAu99h/EbBJx0kZxJyGiGKr0tkiVBisV39dx898Fd1
# rL2KQk1AUdEPnAY+Z3/1ZsADlkR+79BL/W7lmsqxqPJ6Kgox8NpOBpG2iAg16Hgc
# sOmZzTznL0S6p/TcZL2kAcEgCZN4zfy8wMlEXV4WnAEFTyJNAgMBAAGjggHmMIIB
# 4jAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQU1WM6XIoxkPNDe3xGG8UzaFqF
# bVUwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1Ud
# EwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYD
# VR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwv
# cHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEB
# BE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwgaAGA1UdIAEB/wSBlTCB
# kjCBjwYJKwYBBAGCNy4DMIGBMD0GCCsGAQUFBwIBFjFodHRwOi8vd3d3Lm1pY3Jv
# c29mdC5jb20vUEtJL2RvY3MvQ1BTL2RlZmF1bHQuaHRtMEAGCCsGAQUFBwICMDQe
# MiAdAEwAZQBnAGEAbABfAFAAbwBsAGkAYwB5AF8AUwB0AGEAdABlAG0AZQBuAHQA
# LiAdMA0GCSqGSIb3DQEBCwUAA4ICAQAH5ohRDeLG4Jg/gXEDPZ2joSFvs+umzPUx
# vs8F4qn++ldtGTCzwsVmyWrf9efweL3HqJ4l4/m87WtUVwgrUYJEEvu5U4zM9GAS
# inbMQEBBm9xcF/9c+V4XNZgkVkt070IQyK+/f8Z/8jd9Wj8c8pl5SpFSAK84Dxf1
# L3mBZdmptWvkx872ynoAb0swRCQiPM/tA6WWj1kpvLb9BOFwnzJKJ/1Vry/+tuWO
# M7tiX5rbV0Dp8c6ZZpCM/2pif93FSguRJuI57BlKcWOdeyFtw5yjojz6f32WapB4
# pm3S4Zz5Hfw42JT0xqUKloakvZ4argRCg7i1gJsiOCC1JeVk7Pf0v35jWSUPei45
# V3aicaoGig+JFrphpxHLmtgOR5qAxdDNp9DvfYPw4TtxCd9ddJgiCGHasFAeb73x
# 4QDf5zEHpJM692VHeOj4qEir995yfmFrb3epgcunCaw5u+zGy9iCtHLNHfS4hQEe
# gPsbiSpUObJb2sgNVZl6h3M7COaYLeqN4DMuEin1wC9UJyH3yKxO2ii4sanblrKn
# QqLJzxlBTeCG+SqaoxFmMNO7dDJL32N79ZmKLxvHIa9Zta7cRDyXUHHXodLFVeNp
# 3lfB0d4wwP3M5k37Db9dT+mdHhk4L7zPWAUu7w2gUDXa7wknHNWzfjUeCLraNtvT
# X4/edIhJEqGCAs4wggI3AgEBMIH4oYHQpIHNMIHKMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBP
# cGVyYXRpb25zMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjozRTdBLUUzNTktQTI1
# RDElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcG
# BSsOAwIaAxUAv26eVJaumcmTchd6hqayQMNDXluggYMwgYCkfjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIFAOS1PecwIhgPMjAy
# MTA4MDUwMDI1MTFaGA8yMDIxMDgwNjAwMjUxMVowdzA9BgorBgEEAYRZCgQBMS8w
# LTAKAgUA5LU95wIBADAKAgEAAgIEiAIB/zAHAgEAAgIRPDAKAgUA5LaPZwIBADA2
# BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIB
# AAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBAMy3Jt2KPkPKeGanZVQ8SxS7JEvARLh4
# b/mc9nVTM2B28aYRx514M5hvS+r0bR6UTr9aSyPoIWSNqlMeaIeT97Dgx5L3cJgu
# LIzYJ0wnlCKPH09C9K7Ko8gWcT7z2kdFR8ERq/iZOPaw2vZfpTL5TyoOv1YaStg5
# aLKoG0ny78BxMYIDDTCCAwkCAQEwgZMwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
# IDIwMTACEzMAAAFSMEtdiazmcEcAAAAAAVIwDQYJYIZIAWUDBAIBBQCgggFKMBoG
# CSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQxIgQgi9/0KNMa
# Bt+ns591rm+L32s1NfKTacMvzWzk0UL7vLEwgfoGCyqGSIb3DQEJEAIvMYHqMIHn
# MIHkMIG9BCCT7lzHo4slUIxfEGp8LXQNik/ecK6vuuGWIcmBrrsnpjCBmDCBgKR+
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAABUjBLXYms5nBHAAAA
# AAFSMCIEIL1iNonPmoRBH/Ro9dZ/Cd99ledsQqqVXCjBDskxLD5eMA0GCSqGSIb3
# DQEBCwUABIIBADaLfhZUg14vn4IsfDfgxMGLw6P6BVLFi3WzYdKSz1m+oTqyH68g
# fnLFZX9pVN4UPsXCadway5S7YRDPCmW8H7hTpKTVODQv9gECb0jzhPQxM3c8/1Ei
# suY0uZEcr0fubZEJ0g2qyVfg/HvXecm70eHbywMo4TYY1qRiI2wH2iBNQwGCCaqp
# WNP/jgPPfTZu1NJEpFrarNx2mv6vOuZTDz62HdzgL36FkdLi3HPCYKg0BuGGRVfe
# YdTOsK76YiqM94Jhz13qo1WUJjKi8DPmwoQnoF9fLwr902UBsOio5/oM7TO6jq9G
# x2lAEltpTmxhnaI1xhuizPZzSbuUHhizIMk=
# SIG # End signature block
