# ================================================
# Set Default App Associations Script for Action1
# ================================================
# Description:
#   - This script downloads an XML file that contains default app associations for Windows 11.
#   - The XML file is applied using DISM to set default app associations.
#
# Requirements:
#   - Admin rights are required.
#   - The XML file should be hosted via a direct download link (e.g., GitHub).
# ================================================

$ProgressPreference = 'SilentlyContinue'

# Define the URL or local path for the default app associations XML file
$defaultAppAssocPath = ${XML File Path} # Replace this with your XML file's direct link.


function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = "$env:SystemDrive\Logs\Action1.log", # Default log file path
        [string]$Level = "INFO"  # Log level: INFO, WARN, ERROR
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    # Ensure the directory for the log file exists
    $logFileDirectory = Split-Path -Path $LogFilePath -Parent
    if (!(Test-Path -Path $logFileDirectory)) {
        try {
            New-Item -Path $logFileDirectory -ItemType Directory -Force | Out-Null
        } catch {
            Write-Error "Failed to create log file directory: $logFileDirectory. $_"
            return
        }
    }

    # Check log file size and recreate if too large
    if (Test-Path -Path $LogFilePath) {
        $logSize = (Get-Item -Path $LogFilePath -ErrorAction Stop).Length
        if ($logSize -ge 5242880) {
            Remove-Item -Path $LogFilePath -Force -ErrorAction Stop | Out-Null
            Out-File -FilePath $LogFilePath -Encoding utf8 -ErrorAction Stop
            Add-Content -Path $LogFilePath -Value "[$timestamp] [INFO] The log file exceeded the 5 MB limit and was deleted and recreated."
        }
    }
    
    # Write log entry to the log file
    Add-Content -Path $LogFilePath -Value $logMessage

    # Write output to Action1 host
    Write-Output "$Message"
}


function Get-RemoteFile {
    param (
        [string]$fileURL,   # URL to the file to be downloaded
        [string]$destinationPath  # Local path where the file should be saved
    )

    try {
        if ($fileURL -match "^https?://") {
            Write-Log "Downloading file from remote URL: $fileURL" -Level "INFO"
            Invoke-WebRequest -Uri $fileURL -OutFile $destinationPath -ErrorAction Stop
        } elseif (Test-Path $fileURL) {
            Write-Log "Copying file from local/network path: $fileURL" -Level "INFO"
            Copy-Item -Path $fileURL -Destination $destinationPath -Force
        } else {
            throw "The file does not exist or is inaccessible: $fileURL"
        }
    } catch {
        Write-Log "Failed to download or copy the file: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# ================================
# Main Script Logic
# ================================
try {
    # Define the temp folder for the downloaded XML file
    $tempFolder = "$env:TEMP\Action1Files"
    if (-not (Test-Path $tempFolder)) {
        New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null
    }
    $xmlFilePath = Join-Path $tempFolder "DefaultAppAssoc.xml"

    # Download the default app associations XML file
    Get-RemoteFile -fileURL $defaultAppAssocPath -destinationPath $xmlFilePath

    # Apply the default app associations using DISM
    Write-Log "Applying default app associations using DISM." -Level "INFO"
    try {
        Start-Process dism.exe -ArgumentList "/Online /Import-DefaultAppAssociations:$xmlFilePath" -Wait -NoNewWindow
        Write-Log "Default app associations applied successfully." -Level "INFO"
    } catch {
        Write-Log "Failed to apply default app associations: $($_.Exception.Message)" -Level "ERROR"
    }

} catch {
    Write-Log "An error occurred during script execution: $($_.Exception.Message)" -Level "ERROR"
} finally {
    # Clean up the temp folder
    try {
        if (Test-Path $tempFolder) {
            Remove-Item -Path $tempFolder -Recurse -Force
            Write-Log "Temporary files cleaned up." -Level "INFO"
        }
    } catch {
        Write-Log "Failed to clean up temp folder: $($_.Exception.Message)" -Level "ERROR"
    }
}  