Function New-RandomFile {
    <#
        .SYNOPSIS
            Generates a number of random files for sample data.
        
        .DESCRIPTION
            Generates a defined number of files until reaching a maximum size.
        
        .PARAMETER TotalSize
            Specify the total size you would all the files combined should use on the harddrive.
            This parameter accepts the following size values (KB,MB,GB,TB)
                5MB
                3GB
                200KB
        
        .PARAMETER NumberOfFiles
            Specify a number of files that need to be created. This can be used to generate a big number of small files in order to simulate
            User backup specefic behaviour.
        
        .PARAMETER FileTypes
            This parameter is not mandatory, but two choices are valid:
                Office : Will generate files with the following extensions: ".pptx",".docx",".doc",".xls",".docx",".doc",".pdf",".ppt",".pptx",".dot"
                Multimedia : Will create random files with the following extensions : ".avi",".midi",".mov",".mp3",".mp4",".mpeg",".mpeg2",".mpeg3",".mpg",".ogg",".ram",".rm",".wma",".wmv"
            If FileTypes parameter is not set, by default, the script will create both office and multimedia type of files.
        
        .PARAMETER PathForFiles
            Specify a path where the files should be generated. If the Path doesn't exist, it will be created.
        
        .PARAMETER Verbose
            Allow to run the script in verbose mode for debbuging purposes.
        
        .EXAMPLE
        New-RandomFile -TotalSize 50MB -NumberOfFiles 13 -Path C:\Users\Sterling
        
        Will generate randonmly 13 files for a total of 50mb in the path C:\Users\Sterling
        
        .EXAMPLE
        New-RandomFile -TotalSize 5GB -NumberOfFiles 3 -Path C:\Users\Sterling
        
        Will generate randonmly 3 files for a total of 5 Gigabytes in the path C:\Users\Sterling
        
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the number of files you would like to generate")]
        [int]$NumberOfFiles,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the path where the files should be generated")]
        [string]$PathForFiles,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false,
            HelpMessage = "Specify the total size you would like each file to be in KB, MB, GB, or TB")]
        [string]$TotalSize
    )

    begin {
        if ($PathForFiles) {
            try {
                Write-Verbose "Testing if path for files exists..." #| Out-Log
                if (-not (Test-Path -Path $PathForFiles)) {
                    Write-Verbose "Path does not exist. Creating directory at $PathForFiles" #| Out-Log
                    New-Item -Path $PathForFiles -Type Directory
                } else {
                    Write-Verbose "Path $PathForFiles is valid. Continuing..." #| Out-Log
                }
            } catch {
                Write-Error "Error creating directory at $PathForFiles" #| Out-Log
                Write-Error "Please check local permissions." #| Out-Log
                Write-Host $_ #| Out-Log
            }
        }

        Write-Verbose "Generating files..." #| Out-Log
        $AllCreatedFiles = @()

        Function New-FileName {
            [CmdletBinding(SupportsShouldProcess = $true)]
            Param(
                [Parameter(mandatory = $false)]
                [ValidateSet("Multimedia", "Office", "all", "")]
                [String]$FileTypes = $all
            )

            begin {
                $allExtensions = @()
                $multimediaExtensions = ".avi", ".midi", ".mov", ".mp3", ".mp4", ".mpeg", ".mpeg2", ".mpeg3", ".mpg", ".ogg", ".ram", ".rm", ".wma", ".wmv"
                $officeExtensions = ".pptx", ".docx", ".doc", ".xls", ".docx", ".doc", ".pdf", ".ppt", ".pptx", ".dot"
                $allExtensions = $multimediaExtensions + $officeExtensions
                $extension = $null
            }
            process {
                Write-Verbose "Creating file name..." #| Out-Log
                #$Extension = $MultimediaFiles | Get-Random -Count 1

                switch ($FileTypes) {
                    "Multimedia" { $extension = $multimediaExtensions | Get-Random }
                    "Office" { $extension = $officeExtensions | Get-Random }
                    default {
                        $extension = $allExtensions | Get-Random
                    }
                }

                Get-Verb | Select-Object verb | Get-Random -Count 2 | ForEach-Object { $Name += $_.verb }
                $FullName = $name + $extension
                Write-Verbose "File name created : $FullName" #| Out-Log
            }
            end {
                return $FullName
            }
        }
    }
    #----------------Process-----------------------------------------------

    process {

        $fileSize = $TotalSize / $NumberOfFiles
        $fileSize = [Math]::Round($fileSize, 0)

        # Start reporting logic
        if ($fileSize -ge 1TB) {
            $reportFileSize = [Math]::Round($fileSize / 1TB, 2)
            $unit = "TB"
        } elseif ($fileSize -ge 1GB) {
            $reportFileSize = [Math]::Round($fileSize / 1GB, 2)
            $unit = "GB"
        } elseif ($fileSize -ge 1MB) {
            $reportFileSize = [Math]::Round($fileSize / 1MB, 2)
            $unit = "MB"
        } elseif ($fileSize -ge 1KB) {
            $reportFileSize = [Math]::Round($fileSize / 1KB, 2)
            $unit = "KB"
        } else {
            $reportFileSize = [Math]::Round($fileSize, 2)
            $unit = "bytes"
        }


        while ($totalFileSize -lt $TotalSize) {
            $totalFileSize = $totalFileSize + $fileSize

            $fileName = New-FileName -FileTypes $FileTypes

            Write-Verbose "Creating : $fileName at $reportFileSize $unit." #| Out-Log

            Write-Verbose "Filesize = $reportFilesize $unit" #| Out-Log

            $FullPath = Join-Path $PathForFiles -ChildPath $fileName
            Write-Verbose "Generating file : $FullPath at $reportFileSize $unit" #| Out-Log
            try {
                fsutil.exe file createnew $FullPath $fileSize | Out-Null
            } catch {
                $_
            }

            $fileCreated = ""
            $properties = @{
                'FullPath' = $FullPath
                'Size'     = $fileSize 
            }

            $fileCreated = New-Object -TypeName psobject -Property $properties
            $AllCreatedFiles += $fileCreated
            Write-Verbose "File created and located at: $FullPath with a size of $reportFileSize $unit" #| Out-Log
        }

    }
    end {
        Write-Verbose "Opening Explorer to location..." #| Out-Log
        Invoke-Item $PathForFiles
        Write-Host "All files have been created and are located at $PathForFiles" -ForegroundColor Green #| Out-Log
    }
}