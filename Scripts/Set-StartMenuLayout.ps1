# ================================================
# PowerShell Script: Copy Start2.bin
# ================================================
# Description:
#   - This script downloads or copies the Start2.bin file from an online URL, network share, or local path 
#     and copies it to the Start Menu Experience folder on the machine.
#
# Requirements:
#   - Internet access if downloading from a URL.
#   - Admin rights to access user folders and make modifications.
# ================================================

$ProgressPreference = 'SilentlyContinue'

# ================================
# Parameters Section (Customizable)
# ================================
$StartMenuBINFile = {Start Menu BIN File}  # Replace with your actual path or URL
$tempBinPath = "$env:TEMP\Start2.bin"  # Temp file path for downloaded .bin file
$destFolderPath = "$env:SystemDrive\Users\Default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState"

# ================================
# Logging Function: Write-Log
# ================================
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = "$env:SystemDrive\LST-Action1.log", # Default log file path
        [string]$Level = "INFO"  # Log level: INFO, WARN, ERROR
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write log entry to the log file
    Add-Content -Path $LogFilePath -Value $logMessage
}

# ================================
# Pre-Check
# ================================
try {
    Write-Log "Starting the process to copy Start2.bin..." -Level "INFO"
    
    # Check if $StartMenuBINFile is a URL or a local/network path
    if ($StartMenuBINFile -match '^http[s]?://') {
        Write-Log "Downloading Start2.bin from URL: $StartMenuBINFile" -Level "INFO"
        try {
            Invoke-WebRequest -Uri $StartMenuBINFile -OutFile $tempBinPath -UseBasicParsing
            Write-Log "Download completed successfully." -Level "INFO"
        } catch {
            Write-Log "Failed to download Start2.bin: $($_.Exception.Message)" -Level "ERROR"
            return
        }
    } elseif (Test-Path $StartMenuBINFile) {
        # Use local or network path directly
        Write-Log "Using local/network Start2.bin file: $StartMenuBINFile" -Level "INFO"
        $tempBinPath = $StartMenuBINFile
    } else {
        Write-Log "Invalid file path or URL: $StartMenuBINFile" -Level "ERROR"
        return
    }

} catch {
    Write-Log "Pre-check failed: $($_.Exception.Message)" -Level "ERROR"
    return
}

# ================================
# Main Script Logic
# ================================
try {
    # Create the destination folder if it doesn't exist
    Write-Log "Creating destination folder if it doesn't exist: $destFolderPath" -Level "INFO"
    New-Item -ItemType "directory" -Path $destFolderPath -Force | Out-Null
    
    # Copy the Start2.bin file to the destination
    Write-Log "Copying Start2.bin to $destFolderPath" -Level "INFO"
    Copy-Item -Path $tempBinPath -Destination "$destFolderPath\Start2.bin" -Force | Out-Null
    Write-Log "Start2.bin copied successfully." -Level "INFO"
    
} catch {
    Write-Log "An error occurred while copying Start2.bin: $($_.Exception.Message)" -Level "ERROR"
    return
}

# ================================
# Cleanup Files
# ================================
try {
    Write-Log "Cleaning up temporary files..." -LogFilePath $LogFilePath -Level "INFO"
    
    # If the .bin was downloaded, remove the temp file
    if ($StartMenuBINFile -match '^http[s]?://') {
        Remove-Item -Path $tempBinPath -Force -ErrorAction SilentlyContinue
        Write-Log "Temporary files removed." -LogFilePath $LogFilePath -Level "INFO"
    }

} catch {
    Write-Log "Failed to clean up temporary files: $($_.Exception.Message)" -LogFilePath $LogFilePath -Level "ERROR"
}