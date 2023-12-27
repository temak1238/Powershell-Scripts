#Parameter
[CmdletBinding()]
Param
(
    [Parameter(Mandatory = $false)][string]$ConfigFilePath = "C:\ScheduledScripts\Compress-ArtbooksToCBZ\config.json"
)

#MODULES
try {
    Import-Module PS-utilities -MinimumVersion 2023.12.14 -ErrorAction Stop
    Import-Module 7Zip4Powershell
}
catch {
    Write-Error "Something went wrong on Import-Module Part: $error[0]"
    Exit
}

#VARIABLES
Set-PSUtilsWriteLog -SetLogEnable:$true -SetLogLevel DEBUG -SetLogFilePath "$PSScriptRoot\logs" -SetLogFileName "startup" -SetLogRotation $true -SetLogRetentionTime "14"
try {
    if (-not (Test-Path -Path $PSScriptRoot\logs -ErrorAction Stop) ) {
        New-Item -Path $PSScriptRoot\logs -ItemType Directory -ErrorAction Stop | Out-Null
    }
}
catch {
    Write-Error "Unable to initialize the log folder!"
    exit
}    
Write-PSUtilsLogfile -LogLevel INFO -msg "======== GIB FEUER MARIANNE ========"

#Config File import
if ($ConfigFilePath) {
    $jsonData = Get-PSUtilsJSONConfig -filepath $ConfigFilePath
    Write-PSUtilsLogfile -LogLevel INFO -msg "JSON Data from '$ConfigFilePath' loaded."
}
else {
    Write-PSUtilsLogfile -LogLevel ERROR -msg "Could not load JSON Config '$ConfigFilePath': $($error[0])"
    exit
}
foreach ($basefolder in $jsonData.source) {
    if (Test-Path -Path $basefolder\*) {
        Write-PSUtilsLogfile -LogLevel INFO -msg "We will use $basefolder as basefolder now."
        $folders = Get-Item -Path ($basefolder + "*")
        $destfolder = $basefolder + "CBZ"
        try {
            if (Test-Path -Path $destfolder -ErrorAction Stop) {
                Remove-Item -Path $destfolder -Recurse -Force -ErrorAction Stop
            }
        }
        catch {
            Write-PSUtilsLogfile -LogLevel ERROR -msg "CBZ folder was already there but could not be removed!"
            exit
        }
        try {
            if (-not (Test-Path -Path $destfolder -ErrorAction Stop) ) {
                New-Item -Path $destfolder -ItemType Directory -ErrorAction Stop | Out-Null
            }
        }
        catch {
            Write-PSUtilsLogfile -LogLevel ERROR -msg "Unable to initialize the CBZ folder!"
            exit
        }
        foreach ($folder in $folders) {
            $destfile = $destfolder + "\" + $($folder.name) + ".zip"
            $filename = $($folder.name) + ".zip"
            $destcbz = $($folder.name) + ".cbz"
            try {
                Compress-7Zip -ArchiveFileName $filename -Path $($folder.FullName) -OutputPath $destfolder -Format Zip -CompressionLevel Normal
                Write-PSUtilsLogfile -LogLevel INFO -msg "$($folder.name) is now compressed."
                $success = $true
            }
            catch {
                Write-PSUtilsLogfile -LogLevel ERROR -msg "$($folder.name) could not be compressed: $error[0]"
                $success = $false
            }
            if ($success) {
                try {
                    Rename-Item -Path $destfile -NewName $destcbz -ErrorAction Stop
                    Write-PSUtilsLogfile -LogLevel INFO -msg "$($folder.name) is now renamed."
                    $success = $true
                }
                catch {
                    Write-PSUtilsLogfile -LogLevel ERROR -msg "$($folder.name) could not be renamed: $error[0]"
                    $success = $false
                }
            }
            if ($success) {
                switch ($basefolder) {
                    "D:\Mangas\Artbooks\" {
                        try {
                            Move-Item -Path $destfolder\* -Destination "D:\Mangas\Manga_Reader\Artbooks_Old" -Force -ErrorAction Stop
                            Write-PSUtilsLogfile -LogLevel INFO -msg "$($folder.name) is now moved."
                        }
                        catch {
                            Write-PSUtilsLogfile -LogLevel ERROR -msg "$($folder.name) could not be moved: $error[0]"
                        }  
                    }
                    "D:\Mangas\Ero Artbooks\" {
                        try {
                            Move-Item -Path $destfolder\* -Destination "D:\Mangas\Manga_Reader\Ero_Art" -Force -ErrorAction Stop
                            Write-PSUtilsLogfile -LogLevel INFO -msg "$($folder.name) is now moved."
                        }
                        catch {
                            Write-PSUtilsLogfile -LogLevel ERROR -msg "$($folder.name) could not be moved: $error[0]"
                        }  
                    }
                    "D:\Mangas\Ero Cosplay Sets\" {
                        try {
                            Move-Item -Path $destfolder\* -Destination "D:\Mangas\Manga_Reader\Ero_Cosplay" -Force -ErrorAction Stop
                            Write-PSUtilsLogfile -LogLevel INFO -msg "$($folder.name) is now moved."
                        }
                        catch {
                            Write-PSUtilsLogfile -LogLevel ERROR -msg "$($folder.name) could not be moved: $error[0]"
                        }  
                    }
                    Default { Write-PSUtilsLogfile -LogLevel ERROR -msg "NÖP" }
                }
            }
        }
        try {
            Remove-Item -Path $basefolder\* -Recurse -Force -ErrorAction Stop
            Write-PSUtilsLogfile -LogLevel INFO -msg "$destfolder was removed."
        }
        catch {
            Write-PSUtilsLogfile -LogLevel ERROR -msg "$destfolder could not be removed."
        }
    }
}
Write-PSUtilsLogfile -LogLevel INFO -msg "======== MARIANNE IST FERTIG ========"