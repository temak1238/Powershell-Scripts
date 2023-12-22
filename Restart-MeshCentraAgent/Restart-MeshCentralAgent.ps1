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
    $meshprocess = Get-Process "MeshAgent"
    $mem=$meshprocess.PrivateMemorySize/1024/1024
    write-host $mem
    if ($mem -gt 300) {
        try {
            Write-PSUtilsLogfile -LogLevel INFO -msg "Restart MeshAgent, Memory usage > 300MB"
            Restart-Service -Name "Mesh Agent" -Force -ErrorAction Stop
        }
        catch {
            Write-PSUtilsLogfile -LogLevel ERROR -msg "Could not restart MeshAgent: $error[0]"
        }
    }
    else {
        Write-PSUtilsLogfile -LogLevel INFO -msg "MeshAgent memory usage is fine."
    }
Write-PSUtilsLogfile -LogLevel INFO -msg "======== MARIANNE IST FERTIG ========"