#Parameter
[CmdletBinding()]
Param
(
    [Parameter(Mandatory = $false)][string]$ConfigFilePath = "C:\ScheduledScripts\Compress-ImmichToOnedrive\config.json"
)

#MODULES
try {
    Import-Module PS-utilities -MinimumVersion 2023.12.14 -ErrorAction Stop
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

try {
    $Private:convertetpw = ConvertTo-SecureString $($jsondata.compresspw)
    if (Test-Path -Path "$($($jsondata.destination) + $($jsondata.filename))" -PathType Leaf -ErrorAction Stop){
        Compress-7Zip -ArchiveFileName $($jsondata.filename) -Path $($jsondata.source) -OutputPath $($jsondata.destination) -Format SevenZip  -CompressionLevel High -Append -SecurePassword $convertetpw
        Write-PSUtilsLogfile -LogLevel INFO -msg "$($($jsondata.destination) + $($jsondata.filename)) was updated."
    }
    else{
        Compress-7Zip -ArchiveFileName $($jsondata.filename) -Path $($jsondata.source) -OutputPath $($jsondata.destination) -Format SevenZip  -CompressionLevel High -SecurePassword $convertetpw
        Write-PSUtilsLogfile -LogLevel INFO -msg "$($($jsondata.destination) + $($jsondata.filename)) was compressed."
    }
}
catch {
    Write-PSUtilsLogfile -LogLevel ERROR -msg "$($($jsondata.destination) + $($jsondata.filename)) could not be compressed: $error[0]"
}
Write-PSUtilsLogfile -LogLevel INFO -msg "======== MARIANNE IST FERTIG ========"