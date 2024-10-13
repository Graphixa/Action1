# ================================================
# Set Wallpaper and Lock Screen Script for Action1
# ================================================
# Description:
#   - This script downloads and sets the wallpaper and lock screen image on Windows.
#   - The image files can be sourced from a local path, network share, or URL.
#   - Images are downloaded and kept in a folder on $env:SystemDrive\Action1 when necessary.
#
# Requirements:
#   - Admin rights are required.
#   - Supported image formats: .jpg, .png.
# ================================================

$ProgressPreference = 'SilentlyContinue'

$wallpaperUrlOrPath = "${Wallpaper Path}"  # Provide URL or path for the wallpaper image
$lockScreenUrlOrPath = "${LockScreen Path}"  # Provide URL or path for the lock screen image
$downloadLocation = "$env:SystemDrive\Action1" # Path to download and keep wallpaper/lockscreen files *if required*

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
# Download Image Function
# ================================
function Get-Image {
    param (
        [string]$imagePath,   # URL or path to the image file
        [string]$fileName     # File name for saving the image locally
    )

    $localImagePath = if ($imagePath -match "^https?://") { # Only create download location for URLs
        if (-not (Test-Path $downloadLocation)) {
            New-Item -Path $downloadLocation -ItemType Directory -Force | Out-Null
        }
        Join-Path $downloadLocation $fileName
    } else {
        $imagePath
    }

    try {
        if ($imagePath -match "^https?://") {
            Write-Log "Downloading image from URL: $imagePath" -Level "INFO"
            Invoke-WebRequest -Uri $imagePath -OutFile $localImagePath -ErrorAction Stop
        } elseif (-not (Test-Path $imagePath)) {
            throw "The image file does not exist or is inaccessible: $imagePath"
        }

        return $localImagePath
    } catch {
        Write-Log "Failed to download or copy image: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# ================================
# Main Script Logic
# ================================
try {
    Write-Log "Starting Wallpaper and Lock Screen configuration." -Level "INFO"

    # Set Wallpaper
    if ($wallpaperUrlOrPath) {
        try {
            $localWallpaperPath = Get-Image -imagePath $wallpaperUrlOrPath -fileName "wallpaper.jpg"
            Write-Log "Setting wallpaper to: $localWallpaperPath"

            $registryPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
            RegistryTouch -action add -path $registryPath -name "DesktopImagePath" -type "String" -value $localWallpaperPath
            RegistryTouch -action add -path $registryPath -name "DesktopImageUrl" -type "String" -value $localWallpaperPath
            RegistryTouch -action add -path $registryPath -name "DesktopImageStatus" -type "DWord" -value 1
            Write-Log "Wallpaper set successfully." -Level "INFO"
        } catch {
            Write-Log "Error setting wallpaper: $($_.Exception.Message)" -Level "ERROR"
        }
    } else {
        Write-Log "Wallpaper not set. No image path provided." -Level "WARN"
    }

    # Set Lock Screen
    if ($lockScreenUrlOrPath) {
        try {
            $localLockScreenPath = Get-Image -imagePath $lockScreenUrlOrPath -fileName "lockscreen.jpg"
            Write-Log "Setting lock screen image to: $localLockScreenPath"

            $registryPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
            RegistryTouch -action add -path $registryPath -name "LockScreenImagePath" -type "String" -value $localLockScreenPath
            RegistryTouch -action add -path $registryPath -name "LockScreenImageUrl" -type "String" -value $localLockScreenPath
            RegistryTouch -action add -path $registryPath -name "LockScreenImageStatus" -type "DWord" -value 1
            Write-Log "Lock screen image set successfully." -Level "INFO"
        } catch {
            Write-Log "Error setting lock screen image: $($_.Exception.Message)" -Level "ERROR"
        }
    } else {
        Write-Log "Lock screen image not set. No image path provided." -Level "WARN"
    }

    # Restart Explorer for settings to take effect
    Stop-Process -Name explorer -Force
    Start-Process explorer
    Write-Log "Explorer restarted successfully." -Level "INFO"

} catch {
    Write-Log "An error occurred during the setup: $($_.Exception.Message)" -Level "ERROR"
}
