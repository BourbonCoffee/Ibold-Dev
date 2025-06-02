function Import-MailFlowRules {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$XmlFilePath,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force,
        
        [Parameter(Mandatory = $false)]
        [switch]$WhatIf
    )

    # Verify the XML file exists
    if (-not (Test-Path $XmlFilePath)) {
        Write-Error "The specified XML file does not exist: $XmlFilePath"
        return
    }

    # Check if there are existing rules
    $existingRules = Get-TransportRule
    if ($existingRules.Count -ne 0 -and -not $Force) {
        Write-Warning "There are $($existingRules.Count) existing mail flow rules."
        Write-Warning "Use -Force parameter to proceed with import despite existing rules."
        return
    }

    # Load the XML file
    try {
        [xml]$xml = Get-Content $XmlFilePath -ErrorAction Stop
        $rulesToImport = $xml.SelectNodes("//rules/rule")
    } catch {
        Write-Error "Failed to parse the XML file: $_"
        return
    }

    # Check if there are rules to import
    if ($rulesToImport.Count -eq 0) {
        Write-Warning "There are no mail flow rules to be imported from the XML file."
        return
    }

    Write-Host "Importing $($rulesToImport.Count) mail flow rules." -ForegroundColor Cyan
    $successCount = 0
    $failedRules = @()

    # Import each rule
    foreach ($rule in $rulesToImport) {
        Write-Progress -Activity "Importing Mail Flow Rules" -Status "Processing rule: $($rule.Name)" `
            -PercentComplete (($successCount + $failedRules.Count) * 100 / $rulesToImport.Count)
        
        Write-Verbose "Importing rule '$($rule.Name)' $($successCount + $failedRules.Count + 1)/$($rulesToImport.Count)."
        
        try {
            if ($WhatIf) {
                Write-Host "WhatIf: Would import rule '$($rule.Name)'" -ForegroundColor Yellow
                $successCount++
                continue
            }
            
            # Execute the command block from the XML
            $commandBlock = $rule.version.commandBlock.InnerText
            if ([string]::IsNullOrWhiteSpace($commandBlock)) {
                Write-Warning "Rule '$($rule.Name)' has an empty command block. Skipping."
                $failedRules += $rule.Name
                continue
            }
            
            Invoke-Expression $commandBlock -ErrorAction Stop | Out-Null
            Write-Host "Successfully imported rule '$($rule.Name)'" -ForegroundColor Green
            $successCount++
        } catch {
            Write-Error "Failed to import rule '$($rule.Name)': $_"
            $failedRules += $rule.Name
        }
    }

    Write-Progress -Activity "Importing Mail Flow Rules" -Completed

    # Output summary
    Write-Host "`nImport summary:" -ForegroundColor Cyan
    Write-Host "- Successfully imported: $successCount rules" -ForegroundColor $(if ($successCount -gt 0) { "Green" } else { "Red" })
    if ($failedRules.Count -gt 0) {
        Write-Host "- Failed to import: $($failedRules.Count) rules" -ForegroundColor Red
        Write-Host "  Failed rules:" -ForegroundColor Red
        $failedRules | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    }
}