#Parameter
[CmdletBinding()]
Param
(
    [Parameter(Mandatory = $false)][string]$ConfigFilePath = "D:\GitHub\PowerShell\Convert-Handbreak\config.json"
)
try {
    Set-ExecutionPolicy Bypass -scope Process -Force -ErrorAction Stop
}
catch {
    Write-Error "Could not set the ExecutionPolicy: $error[0]"
    Exit
}

#MODULES
try {
    Import-Module PS-Utilities -MinimumVersion 2023.12.14 -ErrorAction Stop
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

if ($jsonData.presettouse){
    $preset = $($jsonData.presettouse)
}
else {
    $preset = "$PSScriptRoot\preset.json"
}

if ($jsonData.presetname) {
    $presetname = $($jsonData.presetname)
}
else {
    $presetname = "H.265 10Bit Original Resolution"
}

#Config File import
if ($ConfigFilePath) {
    $jsonData = Get-PSUtilsJSONConfig -filepath $ConfigFilePath
    Write-PSUtilsLogfile -LogLevel INFO -msg "JSON Data from '$ConfigFilePath' loaded."
}
else {
    Write-PSUtilsLogfile -LogLevel ERROR -msg "Could not load JSON Config '$ConfigFilePath': $($error[0])"
    exit
}

foreach ($_folder in $jsondata.sourcefolder){
    $filelist = Get-ChildItem $_folder -filter *.mkv -recurse

    $num = $filelist | Measure-Object
    $filecount = $num.count
    Write-PSUtilsLogfile -LogLevel INFO -msg "$filecount files found to convert."
         
    ForEach ($file in $filelist) {
        $oldfile = $file.DirectoryName + "\" + $file.BaseName + $file.Extension;
        $newfile = $file.DirectoryName + "\" + $file.BaseName + ".mp4";

        if ($newfile -like '*264*'){
            $newfile = $newfile -replace '264', '265'
        }
        
        Write-Host "Processing - $oldfile"
        Write-PSUtilsLogfile -LogLevel INFO -msg "Start to convert $($file.BaseName)"
        try {
            Start-Process $($jsonData.handbreak) -ArgumentList "--preset-import-file `"$preset`" -Z `"$presetname`" -i `"$oldfile`" -o `"$newfile`"" -Wait -NoNewWindow -ErrorAction Stop
            Write-PSUtilsLogfile -LogLevel INFO -msg "Converted $($file.BaseName)"
            $success = $true
        }
        catch {
            Write-PSUtilsLogfile -LogLevel ERROR -msg "Error on $($file.BaseName) convert process: $error[0]"
            $success = $false
        }
        if ($success){
            try {
                Remove-Item -Path $oldfile -Force -ErrorAction Stop
                Write-PSUtilsLogfile -LogLevel INFO -msg "Removed $oldfile"
            }
            catch {
                Write-PSUtilsLogfile -LogLevel ERROR -msg "$oldfile could not be deleted $error[0]"
            }    
        }
    }
}
Write-PSUtilsLogfile -LogLevel INFO -msg "======== MARIANNE IST FERTIG ========"