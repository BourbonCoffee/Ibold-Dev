# Limitations:
#   Cannot yet run at the system level
#   Only grabs Edge/Chrome because Firefox Bookmarks are stored in a SQLLite DB


# Create folder if doesn't exist
$folderPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Bookmarks and Favorites'
if (-not (Test-Path $folderPath)) { New-Item -Path $folderPath -ItemType Directory }

# Function to export bookmarks
# To-DO: Clean this up...
function Export-Bookmarks {
    param(
        [string]$browserBasePath,
        [string]$fileName,
        [string]$titlePrefix
    )
    
    # Get list of profile directories (including 'Default')
    $profileDirs = Get-ChildItem -Path $browserBasePath -Directory

    foreach ($profileDir in $profileDirs) {
        $browserFilePath = Join-Path $browserBasePath (Join-Path $profileDir.Name $fileName)
        if (Test-Path $browserFilePath) {
            $content = Get-Content -Raw -Path $browserFilePath | ConvertFrom-Json
            $htmlHeader = @"
            <!DOCTYPE NETSCAPE-Bookmark-file-1>
            <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
            <Title>$titlePrefix - $($profileDir.Name)</Title>
            <H1>Bookmarks</H1>
            <DL><p>
"@
            $htmlFooter = @"
            </DL><p>
"@
            $htmlContent = $htmlHeader
            foreach ($bookmark in $content.roots.bookmark_bar.children) {
                $url = $bookmark.url
                $name = $bookmark.name
                $htmlContent += "<DT><A HREF='$url'>$name</A>"
            }
            $htmlContent += $htmlFooter
            $htmlFilePath = Join-Path $folderPath "$($titlePrefix)_$($profileDir.Name)_Bookmarks.html"
            $htmlContent | Set-Content -Path $htmlFilePath
            Write-Host "$titlePrefix - $($profileDir.Name) bookmarks exported to: $htmlFilePath"
        }
    }
}

$chromeBasePath = "$env:LOCALAPPDATA\Google\Chrome\User Data"
$edgeBasePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"

# Export Chrome and Edge bookmarks
Export-Bookmarks -browserBasePath $chromeBasePath -fileName 'Bookmarks' -titlePrefix 'Chrome'
Export-Bookmarks -browserBasePath $edgeBasePath -fileName 'Bookmarks' -titlePrefix 'Edge'