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
# Main Script Logic
# ================================

# Add registry settings to system

try {
    Write-Log "Setting Registry Items for System" -Level "INFO"

    # Disable OneDrive
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Name "DisableFileSyncNGSC" -Value 1 -PropertyType DWord -Force
    Write-Log "OneDrive disabled." -Level "INFO"
    Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
    Write-Log "OneDrive process stopped." -Level "INFO"

    # Disable Cortana
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -Value 0 -PropertyType DWord -Force
    Write-Log "Cortana disabled." -Level "INFO"

    # Disable CoPilot
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\WindowsCopilot" -Name "CopilotEnabled" -Value 0 -PropertyType DWord -Force
    Write-Log "CoPilot disabled." -Level "INFO"

    # Disable Privacy Experience (OOBE)
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OOBE" -Name "	DisablePrivacyExperience" -Value 1 -PropertyType DWord -Force
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
    Stop-Process -Name explorer -Force
    Start-Process explorer

} catch {
    Write-Log "An error adding registry entries: $($_.Exception.Message)" -Level "ERROR"
    return
}

# Remove bloatware from the system
try {
    Write-Log "Removing Pre-Installed Bloatware" -Level "INFO"
    
    #Removes bloatware apps for all users
    Get-AppXPackage -AllUsers | Where-Object -Property 'Name' -In -Value @(
    'Microsoft.Microsoft3DViewer';
    'MicrosoftWindows.Client.WebExperience';
    'Microsoft.BingSearch';
    'Clipchamp.Clipchamp';
    'Microsoft.549981C3F5F10';
    'Microsoft.Windows.DevHome';
    'MicrosoftCorporationII.MicrosoftFamily';
    'Microsoft.WindowsFeedbackHub';
    'Microsoft.GetHelp';
    'microsoft.windowscommunicationsapps';
    'Microsoft.WindowsMaps';
    'Microsoft.ZuneVideo';
    'Microsoft.BingNews';
    'Microsoft.MicrosoftOfficeHub';
    'Microsoft.Office.OneNote';
    'Microsoft.OutlookForWindows';
    'Microsoft.Paint';
    'Microsoft.MSPaint';
    'Microsoft.People';
    'Microsoft.PowerAutomateDesktop';
    'MicrosoftCorporationII.QuickAssist';
    'Microsoft.SkypeApp';
    'Microsoft.ScreenSketch';
    'Microsoft.MicrosoftSolitaireCollection';
    'Microsoft.MicrosoftStickyNotes';
    'MSTeams';
    'Microsoft.Getstarted';
    'Microsoft.Windows.PeopleExperienceHost';
    'Microsoft.XboxGameCallableUI';
    'Microsoft.WidgetsPlatformRuntime';
    'Microsoft.Todos';
    'Microsoft.WindowsSoundRecorder';
    'Microsoft.BingWeather';
    'Microsoft.ZuneMusic';
    'Microsoft.WindowsTerminal';
    'Microsoft.Xbox.TCUI';
    'Microsoft.XboxApp';
    'Microsoft.XboxGameOverlay';
    'Microsoft.XboxGamingOverlay';
    'Microsoft.XboxIdentityProvider';
    'Microsoft.XboxSpeechToTextOverlay';
    'Microsoft.GamingApp';
    'Microsoft.549981C3F5F10';
    'Microsoft.MixedReality.Portal';
    'Microsoft.Windows.Ai.Copilot.Provider';
    'Microsoft.WindowsMeetNow';
    ) | Remove-AppXPackage

} catch {
    Write-Log "An error occurred removing bloatware apps: $($_.Exception.Message)" -Level "ERROR"
    return
}