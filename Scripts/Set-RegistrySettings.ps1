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
    
    # Write log entry to the log file
    Add-Content -Path $LogFilePath -Value $logMessage

    # Write output to Action1 host
    Write-Output "$Message"
}

# ================================
# RegistryTouch Function to Add or Remove Registry Entries
# ================================

function RegistryTouch {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("add", "remove")]
        [string]$action,

        [Parameter(Mandatory = $true)]
        [string]$path,

        [Parameter(Mandatory = $true)]
        [string]$name,

        [Parameter()]
        [ValidateSet("String", "ExpandString", "Binary", "DWord", "MultiString", "QWord")]
        [string]$type = "String",  # Default to String

        [Parameter()]
        [object]$value  # Changed to object to handle various data types
    )

    try {
        # Detect the base registry hive
        switch -regex ($path) {
            '^HKLM:\\|^HKEY_LOCAL_MACHINE\\' {
                $baseKey = "HKLM:"
                $pathWithoutHive = $path -replace '^HKLM:\\|^HKEY_LOCAL_MACHINE\\', ''
            }
            '^HKCU:\\|^HKEY_CURRENT_USER\\' {
                $baseKey = "HKCU:"
                $pathWithoutHive = $path -replace '^HKCU:\\|^HKEY_CURRENT_USER\\', ''
            }
            '^HKCR:\\|^HKEY_CLASSES_ROOT\\' {
                $baseKey = "HKCR:"
                $pathWithoutHive = $path -replace '^HKCR:\\|^HKEY_CLASSES_ROOT\\', ''
            }
            '^HKU:\\|^HKEY_USERS\\' {
                $baseKey = "HKU:"
                $pathWithoutHive = $path -replace '^HKU:\\|^HKEY_USERS\\', ''
            }
            '^HKCC:\\|^HKEY_CURRENT_CONFIG\\' {
                $baseKey = "HKCC:"
                $pathWithoutHive = $path -replace '^HKCC:\\|^HKEY_CURRENT_CONFIG\\', ''
            }
            default {
                Write-Log "Unsupported registry hive in path: $path" -Level "WARN"
                return
            }
        }

        # Build the full registry path
        $fullKeyPath = $baseKey
        $subKeyParts = $pathWithoutHive -split "\\"

        # Incrementally create any missing registry keys
        foreach ($part in $subKeyParts) {
            $fullKeyPath = Join-Path $fullKeyPath $part
            if (-not (Test-Path $fullKeyPath)) {
                New-Item -Path $fullKeyPath -Force -ErrorAction Stop | Out-Null
            }
        }

        # Now that all parent keys exist, handle the registry value
        if ($action -eq "add") {
            $itemExists = Get-ItemProperty -Path $fullKeyPath -Name $name -ErrorAction SilentlyContinue

            if (-not $itemExists) {
                Write-Log "Registry item added: $name with value: $value" -Level "INFO"
                New-ItemProperty -Path $fullKeyPath -Name $name -Value $value -PropertyType $type -Force -ErrorAction Stop | Out-Null
            } else {
                # Retrieve the current value with proper data type handling
                $currentValue = (Get-ItemProperty -Path $fullKeyPath -Name $name).$name

                # Convert current value and new value to the appropriate types for comparison
                switch ($type) {
                    "DWord" {
                        $currentValue = [int]$currentValue
                        $value = [int]$value
                    }
                    "QWord" {
                        $currentValue = [long]$currentValue
                        $value = [long]$value
                    }
                    "Binary" {
                        $currentValue = [Byte[]]$currentValue
                        $value = [Byte[]]$value
                    }
                    default {
                        $currentValue = [string]$currentValue
                        $value = [string]$value
                    }
                }

                if (-not ($currentValue -eq $value)) {
                    Write-Log "Registry value differs. Updating item: $name from $currentValue to $value" -Level "INFO"
                    Set-ItemProperty -Path $fullKeyPath -Name $name -Value $value -Force -ErrorAction Stop
                } else {
                    Write-Log "Registry item: $name with value: $value already exists. Skipping." -Level "WARN"
                }
            }
        } elseif ($action -eq "remove") {
            # Check if the registry value exists
            if (Get-ItemProperty -Path $fullKeyPath -Name $name -ErrorAction SilentlyContinue) {
                Write-Log "Removing registry item: $name from path: $fullKeyPath" -Level "INFO"
                Remove-ItemProperty -Path $fullKeyPath -Name $name -Force -ErrorAction Stop
            } else {
                Write-Log "Registry item: $name does not exist at path: $fullKeyPath. Skipping." -Level "WARN"
            }
        }
    } catch {
        Write-Log "Error modifying the registry: $($_.Exception.Message)" -Level "ERROR"
    }
}

# ================================
# Main Script Logic
# ================================

try {
    Write-Log "Setting Registry Items for System" -Level "INFO"

    # Disable OneDrive
    RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -name "DisableFileSyncNGSC" -value 1 -type "DWord"
    Write-Log "OneDrive disabled." -Level "INFO"
    try {
        Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
        Write-Log "OneDrive process stopped." -Level "INFO"
    } catch {
        Write-Log "OneDrive process was not running or could not be stopped." -Level "WARN"
    }

    # Disable Cortana
    RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -name "AllowCortana" -value 0 -type "DWord"

    # Disable CoPilot
    RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\WindowsCopilot" -name "CopilotEnabled" -value 0 -type "DWord"

    # Disable Privacy Experience (OOBE)
    RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OOBE" -name "DisablePrivacyExperience" -value 1 -type "DWord"

    # Disable 'Meet Now' in Taskbar
    RegistryTouch -action add -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" -name "HideSCAMeetNow" -value 1 -type "DWord"

    # Disable News and Interests
    RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" -name "AllowNewsAndInterests" -value 0 -type "DWord"

    # Disable Personalized Advertising
    RegistryTouch -action add -path "HKLM:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -name "Enabled" -value 0 -type "DWord"

    # Disable Start Menu Suggestions and Windows Advertising
    RegistryTouch -action add -path "HKLM:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -name "SubscribedContent-338389Enabled" -value 0 -type "DWord"

    # Disable First Logon Animation
    RegistryTouch -action add -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -name "EnableFirstLogonAnimation" -value 0 -type "DWord"

    # Disable Lock Screen App Notifications
    RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -name "DisableLockScreenAppNotifications" -value 1 -type "DWord"

    # Restart Explorer to apply changes
    try {
        Stop-Process -Name explorer -Force
        Start-Process explorer
        Write-Log "Explorer restarted successfully." -Level "INFO"
    } catch {
        Write-Log "Failed to restart Explorer: $($_.Exception.Message)" -Level "ERROR"
    }

} catch {
    Write-Log "An error occurred while setting registry entries: $($_.Exception.Message)" -Level "ERROR"
}