function Import-BrowserBookmarks {
    <#
        .SYNOPSIS
            This cmdlet will import bookmarks into Chrome and Edge browsers.

        .INPUTS
            You can pipe objects to this function

        .OUTPUTS
            None

        .EXAMPLE
            Import-Bookmarks

            This would import bookmarks into both Chrome and Edge browsers.

        .NOTES
            - 5.1.2024.0220:    New function
    #>
    [CmdletBinding()]
    Param(
        [string]$chromeBasePath = "$env:LOCALAPPDATA\Google\Chrome\User Data",
        [string]$edgeBasePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
    )

    $desktop = Get-KnownFolderPath "Desktop"
    $folderPath = Join-Path $desktop 'Bookmarks and Favorites'

    # Get list of HTML files
    $htmlFiles = Get-ChildItem -Path $folderPath -Filter "*.html"
    
    foreach ($htmlFile in $htmlFiles) {
        # Determine browser and profile from file name
        $fileName = $htmlFile.Name
        $browser = if ($fileName -like "*Chrome*") { "Chrome" } elseif ($fileName -like "*Edge*") { "Edge" } else { "Unknown" }
        $profile = if ($fileName -match "(?<=_$browser_).*?(?=_Bookmarks)") { $Matches[0] } else { "Default" }
        
        # Determine base path
        $basePath = if ($browser -eq "Chrome") { $chromeBasePath } elseif ($browser -eq "Edge") { $edgeBasePath } else { "" }
        
        # Skip unknown files or unsupported browsers
        if ($browser -eq "Unknown" -or $basePath -eq "") {
            Write-Warning "Unsupported file: $($htmlFile.FullName)"
            continue
        }
        
        # Import bookmarks
        $bookmarkPath = Join-Path $basePath $profile
        $bookmarkFile = Join-Path $bookmarkPath "Bookmarks"
        
        # Backup existing bookmarks file
        $backupFile = Join-Path $bookmarkPath "Bookmarks_backup"
        if (Test-Path $bookmarkFile) {
            Copy-Item -Path $bookmarkFile -Destination $backupFile -Force
        }
        
        # Copy HTML content to bookmarks file
        Copy-Item -Path $htmlFile.FullName -Destination $bookmarkFile -Force
        Write-Host "Bookmarks imported into $browser - Profile: $profile"
    }
}
