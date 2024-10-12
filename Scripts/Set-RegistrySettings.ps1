# ================================================
# Set Registry Settings Script for Action1
# ================================================
# Description:
#   - This script disables OneDrive, CoPilot, 'Meet Now', Taskbar Widgets, News and Interests, Personalized Advertising, Start Menu Tracking, and Start Menu Suggestions.
#   - The script will also restart Explorer to apply the changes.
#
# Requirements:
#   - Admin rights are required.
#   - Script must be run with administrative privileges to modify registry keys.
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
# Main Script Logic
# ================================

try {
    Write-Log "Setting Registry Items for System" -Level "INFO"

    # Disable OneDrive
    if (-not (Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive")) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Force
    }
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Name "DisableFileSyncNGSC" -Value 1 -PropertyType DWord -Force
    Write-Log "OneDrive disabled." -Level "INFO"
    try {
        Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
        Write-Log "OneDrive process stopped." -Level "INFO"
    } catch {
        Write-Log "OneDrive process was not running or could not be stopped." -Level "WARN"
    }

    # Disable Cortana
    if (-not (Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search")) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Force
    }
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -Value 0 -PropertyType DWord -Force
    Write-Log "Cortana disabled." -Level "INFO"

    # Disable CoPilot
    if (-not (Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\WindowsCopilot")) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\WindowsCopilot" -Force
    }
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\WindowsCopilot" -Name "CopilotEnabled" -Value 0 -PropertyType DWord -Force
    Write-Log "CoPilot disabled." -Level "INFO"

    # Disable Privacy Experience (OOBE)
    if (-not (Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OOBE")) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OOBE" -Force
    }
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OOBE" -Name "DisablePrivacyExperience" -Value 1 -PropertyType DWord -Force
    Write-Log "Privacy Experience (OOBE) disabled." -Level "INFO"

    # Disable 'Meet Now' in Taskbar
    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "HideSCAMeetNow" -Value 1 -PropertyType DWord -Force
    Write-Log "'Meet Now' disabled in Taskbar." -Level "INFO"

    # Disable News and Interests
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" -Name "AllowNewsAndInterests" -Value 0 -PropertyType DWord -Force
    Write-Log "News and Interests disabled." -Level "INFO"

    # Disable Personalized Advertising
    New-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -Name "Enabled" -Value 0 -PropertyType DWord -Force
    Write-Log "Personalized Advertising disabled." -Level "INFO"

    # Disable Start Menu Suggestions and Windows Advertising
    New-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-338389Enabled" -Value 0 -PropertyType DWord -Force
    Write-Log "Start Menu Suggestions and Windows Advertising disabled." -Level "INFO"

    # Restart Explorer to apply changes
    Write-Log "Restarting Explorer to apply changes..." -Level "INFO"
    try {
        Stop-Process -Name explorer -Force
        Start-Process explorer
        Write-Log "Explorer restarted successfully." -Level "INFO"
    } catch {
        Write-Log "Failed to restart Explorer: $($_.Exception.Message)" -Level "ERROR"
    }

} catch {
    Write-Log "An error adding registry entries: $($_.Exception.Message)" -Level "ERROR"
    return
}
