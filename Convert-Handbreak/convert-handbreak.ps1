#Parameter
[CmdletBinding()]
Param
(
    [Parameter(Mandatory = $false)][string]$ConfigFilePath = "C:\ScheduledScripts\Convert-Handbreak\config.json"
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
$filelist = Get-ChildItem $($jsonData.sourcefolder) -filter *.mkv -recurse

$num = $filelist | Measure-Object
$filecount = $num.count
Write-PSUtilsLogfile -LogLevel INFO -msg "$filecount files found to convert."
     
$i = 0;
ForEach ($file in $filelist) {
    $i++;
    $oldfile = $file.DirectoryName + "\" + $file.BaseName + $file.Extension;
    $newfile = $file.DirectoryName + "\" + $file.BaseName + ".mp4";
          
    $progress = ($i / $filecount) * 100
    $progress = [Math]::Round($progress, 2)
     
    Clear-Host
    Write-Host -------------------------------------------------------------------------------
    Write-Host Handbrake Batch Encoding
    Write-Host "Processing - $oldfile"
    Write-Host "File $i of $filecount - $progress%"
    Write-Host -------------------------------------------------------------------------------
    Write-PSUtilsLogfile -LogLevel INFO -msg "Start to convert $($file.BaseName)"
    try {
        Start-Process $($jsonData.handbreak) -ArgumentList "--preset-import-file $($jsonData.presettouse) -Z $($jsonData.presetname) -i "$oldfile" -o "$newfile" --verbose=0" -Wait -NoNewWindow -ErrorAction Stop
        Write-PSUtilsLogfile -LogLevel INFO -msg "Converted $($file.BaseName)"     
    }
    catch {
        Write-PSUtilsLogfile -LogLevel ERROR -msg "Error on $($file.BaseName) convert process: $error[0]"     
    }
}
Write-PSUtilsLogfile -LogLevel INFO -msg "======== MARIANNE IST FERTIG ========"

