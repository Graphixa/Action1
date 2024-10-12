# ================================================
# PowerShell Script Template for Action1
# ================================================
# Description:
#   - Provide a brief description of the script's purpose.
#   - Example: This script installs and configures the XYZ software.
#
# Requirements:
#   - List any prerequisites or requirements (e.g., internet access, admin rights).
#
# Author: [Your Name]
# Date: [Date]
# ================================================

$ProgressPreference = 'SilentlyContinue'

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
# Parameters Section (Customizable)
# ================================
# Define any custom parameters here.
# For example:
#   $softwareName = 'SoftwareName'
#   $installPath = "$env:SystemDrive\Program Files\$softwareName"

$softwareName = 'SoftwareName'
$installPath = "$env:SystemDrive\Program Files\$softwareName"
$tempPath = "$env:SystemDrive\Temp\"

# ================================
# Pre-Check Section (Optional)
# ================================
# Use this section for pre-checks.
# Example: Check if the software is already installed, exit if no action is required.

$checkLocation = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
if (Get-ChildItem $checkLocation -Recurse -ErrorAction Stop | Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -Match "^$softwareName.*" }) {
    Write-Log "$softwareName is already installed. No action required." -Level "INFO"
    return
}

Write-Log "$softwareName is NOT installed. Proceeding with installation..." -Level "INFO"

# ================================
# Main Script Logic
# ================================
# Add your main logic for downloading, installing, configuring, etc.
# Example: Download the software installer, install it, create shortcuts, etc.

try {
    # Example download operation
    $downloadURL = 'https://example.com/software-latest.exe'
    $installerFile = "$tempPath\$softwareName-latest.exe"
    
    Write-Log "Downloading $softwareName from $downloadURL" -Level "INFO"
    Invoke-WebRequest -Uri $downloadURL -OutFile $installerFile
    Write-Log "Download completed: $installerFile" -Level "INFO"
    
    # Example installation
    Write-Log "Installing $softwareName." -Level "INFO"
    Start-Process -Wait -FilePath $installerFile -ArgumentList "/S /D=$installPath" -PassThru
    Write-Log "$softwareName installation completed." -Level "INFO"

    # Example shortcut creation
    $executablePath = "$installPath\$softwareName.exe"
    $desktopShortcutPath = "$env:Public\Desktop\$softwareName.lnk"
    $startMenuShortcutPath = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\$softwareName.lnk"
    
    $WScriptShell = New-Object -ComObject ("WScript.Shell")
    $desktopShortcut = $WScriptShell.CreateShortcut($desktopShortcutPath)
    $desktopShortcut.TargetPath = $executablePath
    $desktopShortcut.Save()

    $startMenuShortcut = $WScriptShell.CreateShortcut($startMenuShortcutPath)
    $startMenuShortcut.TargetPath = $executablePath
    $startMenuShortcut.Save()

    Write-Log "Shortcuts created for $softwareName." -Level "INFO"
    
} catch {
    Write-Log "An error occurred during $softwareName installation: $($_.Exception.Message)" -Level "ERROR"
    return
}

# ================================
# Cleanup Section
# ================================
# Clean up temporary files or logs here.
# Example: Remove temporary installation files.

try {
    Write-Log "Cleaning up temporary files." -Level "INFO"
    Remove-Item -Path $tempPath -Recurse -Force
    Write-Log "Cleanup completed successfully." -Level "INFO"
} catch {
    Write-Log "Failed to remove temporary files: $($_.Exception.Message)" -Level "ERROR"
}

# ================================
# End of Script
# ================================
