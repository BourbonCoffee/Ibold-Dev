Import-Module -Name ImportExcel

# Import the necessary module to handle Excel file operations.

# Define the directory to scan
$SharedDrivePath = "C:\Path\To\ShareDrive"
# Define the output Excel file path
$OutputExcelFile = "C:\Path\To\Output\FileReport.xlsx"

# Create a hashtable to store file hashes
$FileHashes = @{}

# Function to calculate the hash of a file
function Get-FileHash {
    param (
        [string]$FilePath  # Path of the file to hash
    )

    try {
        if (Test-Path $FilePath) {
            # Check if the file exists
            Write-Host "Calculating hash for file: $FilePath"
            $HashAlgorithm = [System.Security.Cryptography.SHA256]::Create()  # Create a SHA256 hash algorithm instance
            try {
                $Stream = [System.IO.File]::OpenRead($FilePath)  # Open the file stream for reading
                try {
                    $Hash = $HashAlgorithm.ComputeHash($Stream)  # Compute the hash of the file
                } finally {
                    $Stream.Close()  # Ensure the file stream is closed
                }
                Write-Host "Hash calculated successfully for file: $FilePath"
                return [BitConverter]::ToString($Hash) -replace "-"  # Return the hash as a string
            } finally {
                $HashAlgorithm.Dispose()  # Dispose of the hash algorithm instance to free resources
            }
        } else {
            Write-Host "File not found: $FilePath"
            return $null  # Return null if the file does not exist
        }
    } catch {
        Write-Host "An error occurred while reading the file: $FilePath. Error: $_"  # Log any errors
        return $null  # Return null on error
    }
}

# Scan files and compute hashes
Write-Host "Scanning files in $SharedDrivePath..."
$Files = Get-ChildItem -Path $SharedDrivePath -Recurse -File  # Get all files in the directory and subdirectories
Write-Host "Total files found: $($Files.Count)"
$TotalFiles = $Files.Count  # Total number of files found
$Progress = 0  # Initialize progress counter

foreach ($File in $Files) {
    $Progress++  # Increment the progress counter
    Write-Progress -Activity "Scanning Files" -Status "Processing $Progress of $TotalFiles" -PercentComplete (($Progress / $TotalFiles) * 100)  # Display progress
    Write-Host "Processing file $Progress of $($TotalFiles): $($File.FullName)"
    $Hash = Get-FileHash -FilePath $File.FullName  # Calculate the hash for the current file
    if ($Hash) {
        if (-not $FileHashes.ContainsKey($Hash)) {
            # Check if the hash is new
            Write-Host "New hash encountered: $Hash"
            $FileHashes[$Hash] = @()  # Initialize an array for the new hash
        } else {
            Write-Host "Duplicate hash found: $Hash"
        }
        $FileHashes[$Hash] += $File.FullName  # Add the file path to the hash entry
    }
}

# Prepare data for the Excel report
$UniqueFiles = @()  # Initialize the array for unique files
$DuplicateFiles = @()  # Initialize the array for duplicate files

# Optimize iteration by precomputing reverse mappings
$HashToFileMap = @{}  # Create a mapping from file paths to their hashes
foreach ($Key in $FileHashes.Keys) {
    foreach ($File in $FileHashes[$Key]) {
        $HashToFileMap[$File] = $Key  # Map each file path to its hash
    }
}

Write-Host "Preparing data for Excel report..."
foreach ($FilesWithHash in $FileHashes.Values) {
    if ($FilesWithHash.Count -eq 1) {
        # Check if the hash is unique
        Write-Host "Unique file identified: $($FilesWithHash[0])"
        $UniqueFiles += [PSCustomObject]@{
            FilePath = $FilesWithHash[0]  # Add the unique file's path
            Hash     = $HashToFileMap[$FilesWithHash[0]]  # Add the hash of the unique file
        }
    } else {
        Write-Host "Duplicate files identified: $($FilesWithHash -join ", ")"
        foreach ($FilePath in $FilesWithHash) {
            # Add each duplicate file
            $DuplicateFiles += [PSCustomObject]@{
                FilePath = $FilePath  # Add the file path
                Hash     = $HashToFileMap[$FilePath]  # Add the hash
            }
        }
    }
}

# Export to Excel
Write-Host "Generating Excel report at $OutputExcelFile..."
if (Test-Path $OutputExcelFile) {
    Write-Host "Output file already exists. Deleting: $OutputExcelFile"
    Remove-Item $OutputExcelFile  # Remove the existing file to avoid conflicts
}

$UniqueFiles | Export-Excel -Path $OutputExcelFile -WorksheetName "UniqueFiles" -AutoSize  # Export unique files
$DuplicateFiles | Export-Excel -Path $OutputExcelFile -WorksheetName "DuplicateFiles" -AutoSize  # Export duplicate files

Write-Host "Report generation complete! File saved to: $OutputExcelFile"
