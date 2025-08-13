[CmdletBinding()]
param (
    [switch]$ForceRun
)

# Get the script's path dynamically
$scriptpath = $PSScriptRoot

# Read and parse settings.json
$settings = Get-Content (Join-Path $scriptpath "settings.json") -Raw | ConvertFrom-Json

# Create new log file path
$logDirectory = Join-Path $scriptpath "log"
if (!(Test-Path $logDirectory)) {
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
}
$log = Join-Path $logDirectory "$(Get-Date -format FileDateTime).log"
New-Item -Path $log -Force | Out-Null

Write-Output "----------------------------------------" | Tee-Object -File $log -Append
Write-Output "Starting send-mrc.ps1 script execution at $(Get-Date)" | Tee-Object -File $log -Append
Write-Output "----------------------------------------" | Tee-Object -File $log -Append

# This script deletes old marc export files and then renames them and send them to collection hq

# Only run this script on certain days of the week
if (-not $ForceRun.IsPresent -and (Get-Date).DayOfWeek -ne $settings.dayofweektorun) {
    Write-Output "Not the proper day of the week to run... exiting" | Tee-Object -File $log -Append
    exit
}

$allFilesToUpload = [System.Collections.ArrayList]@()

# Loop through each folder specified in settings
foreach ($folder in $settings.folders) {
    $path = Join-Path $settings.basepath $folder
    Write-Output "Processing path: $path" | Tee-Object -File $log -Append

    if (!(Test-Path $path)) {
        Write-Output "Path not found: $path. Skipping." | Tee-Object -File $log -Append
        continue
    }

    # Delete all non MRC txt files regardless of age
    Get-childitem -Path $path -File | Where-Object { ($_.Extension -ne $settings.ext) } | Remove-Item

    # Delete all Files more than NN hours old
    Get-ChildItem -path $path | Where-Object LastWriteTime -lt ((Get-Date).AddHours(-1 * $settings.hours)) | Remove-Item

    # Enumerate the remaining file(s)
    $files = Get-ChildItem -path $path

    Write-Output "Checking for files in $path" | Tee-Object -File $log -Append

    # Rename files based on the rules in settings.json
    foreach ($file in $files) {

        $hourcreated = $file.CreationTime.ToString("HH")
        $fileext = $file.Extension
        $fileBasename = $file.BaseName

        $libraryname = $settings.defaultLibraryName

        foreach ($mapping in $settings.libraryMappings) {
            if (($mapping.psobject.Properties.Name -contains 'hour' -and $hourcreated -eq $mapping.hour) -or `
                ($mapping.psobject.Properties.Name -contains 'basename' -and $fileBasename -eq $mapping.basename)) {
                $libraryname = $mapping.libraryName
                break
            }
        }

        try {
            Rename-Item -Path $file.FullName -NewName "$libraryname$($file.LastWriteTime.ToString('yyyyMMddTHHmmssffff'))$fileext" -ErrorAction Stop
        }
        catch {
            Write-Output "FATAL: Failed to rename file $($file.FullName). Error: $($_.Exception.Message)" | Tee-Object -File $log -Append
            Write-Output "Exiting script." | Tee-Object -File $log -Append
            exit 1
        }
    }

    # Add the processed files from this path to the upload list
    $processedFiles = Get-ChildItem -Path $path | Where-Object Extension -EQ $settings.ext
    if ($null -ne $processedFiles) {
        $allFilesToUpload.AddRange($processedFiles)
    }
}


# --- FTP Upload using WinSCP ---
if ($allFilesToUpload.Count -gt 0) {
    # Load WinSCP .NET assembly
    Add-Type -Path (Join-Path $scriptpath "WinSCPnet.dll")

    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Ftp
        HostName = $settings.ftp.url
        UserName = $settings.ftp.user
        Password = $settings.ftp.pass
    }

    $retries = if ($settings.ftp.retries) { $settings.ftp.retries } else { 2 }
    $attempt = 0
    $uploadSuccess = $false

    while ($attempt -lt $retries -and -not $uploadSuccess) {
        $attempt++
        try {
            Write-Output "Setting up FTP session to $($settings.ftp.url) (Attempt $attempt of $retries)" | Tee-Object -File $log -Append
            $session = New-Object WinSCP.Session
            try {
                # Connect
                $session.Open($sessionOptions)

                foreach ($file in $allFilesToUpload) {
                    Write-Output "Uploading Marc $($file.Name) from $($file.DirectoryName)..." | Tee-Object -File $log -Append

                    $transferOptions = New-Object WinSCP.TransferOptions
                    $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

                    # Upload files
                    $transferResult = $session.PutFiles($file.FullName, $settings.ftp.remotepath, $False, $transferOptions)

                    # Throw on any error
                    $transferResult.Check()

                    # Print results
                    foreach ($transfer in $transferResult.Transfers) {
                        Write-Output "Upload of $($transfer.FileName) succeeded" | Tee-Object -File $log -Append
                    }
                }
            }
            finally {
                $session.Dispose()
            }
            Write-Output "FTP Session Closed" | Tee-Object -File $log -Append
            $uploadSuccess = $true
        }
        catch {
            Write-Output "Error during FTP upload (Attempt $attempt of $retries): $($_.Exception.Message)" | Tee-Object -File $log -Append
            if ($attempt -ge $retries) {
                Write-Output "FTP upload failed after $retries attempts. Exiting." | Tee-Object -File $log -Append
                exit 1
            }
            Write-Output "Retrying in 5 seconds..." | Tee-Object -File $log -Append
            Start-Sleep -Seconds 5
        }
    }

    # --- Summary ---
    Write-Output "----------------------------------------" | Tee-Object -File $log -Append
    Write-Output "Script finished at $(Get-Date)" | Tee-Object -File $log -Append
    Write-Output "Summary: $($allFilesToUpload.Count) files processed and uploaded." | Tee-Object -File $log -Append
    Write-Output "----------------------------------------" | Tee-Object -File $log -Append

} else {
    Write-Output "No files found to upload." | Tee-Object -File $log -Append

    # --- Summary ---
    Write-Output "----------------------------------------" | Tee-Object -File $log -Append
    Write-Output "Script finished at $(Get-Date)" | Tee-Object -File $log -Append
    Write-Output "Summary: 0 files processed for upload." | Tee-Object -File $log -Append
    Write-Output "----------------------------------------" | Tee-Object -File $log -Append
}
