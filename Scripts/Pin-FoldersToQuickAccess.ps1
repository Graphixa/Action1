# =======================================================
# Pin Folders to All Users Quick Access Script for Action1
# =======================================================
# Description:
#   - This script pins specific folders to Quick Access in File Explorer for all users.
#   - Uses registry modifications that apply to both existing and new user profiles.
#   - Supports local user folders (Downloads, Documents, etc.), Google Drive "My Drive", and shared drives.
#   - Changes are applied immediately and persist for new user logins.
#
# Requirements:
#   - Admin rights are required.
#   - Script must be run with administrative privileges to modify registry.
#
# ================================================

$ProgressPreference = 'SilentlyContinue'

$PinMyDrive = ${Pin My Drive} # Example: 1 to pin My Drive, 0 to not pin My Drive
$SharedFolders = ${Shared Drives To Pin} # Example: "Shared Drive 1, Shared Drive 2"
$LocalFolders = ${User Folders To Pin} # Example: "Downloads, Documents" will pin the %userprofile%\Downloads and %userprofile%\Documents folders

$LogFilePath = "$env:SystemDrive\LST\Action1.log" # Default log file path

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = $LogFilePath,
        [string]$Level = "INFO"
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
    
    Add-Content -Path $LogFilePath -Value $logMessage
    Write-Output "$Message"
}

function Set-QuickAccessForAllUsers {
    param (
        [string[]]$FoldersToPin
    )
    
    try {
        # Registry path for Quick Access settings (applies to all users)
        $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HomeFolder\NameSpace\DelegateFolders"
        
        # Create the registry key if it doesn't exist
        if (!(Test-Path -Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
            Write-Log "Created registry key: $registryPath"
        }
        
        # Clear existing Quick Access entries
        Get-ChildItem -Path $registryPath | Remove-Item -Force
        Write-Log "Cleared existing Quick Access entries"
        
        # Add new Quick Access entries
        foreach ($folder in $FoldersToPin) {
            $folderName = ""
            $folderPath = ""
            
            if ($folder -eq "My Drive") {
                $folderName = "My Drive"
                $folderPath = "G:\My Drive"
            }
            elseif ($LocalFolders -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -eq $folder }) {
                $folderName = "$folder"
                $folderPath = "%USERPROFILE%\$folder"
            }
            else {
                $folderName = "$folder"
                $folderPath = "G:\Shared drives\$folder"
            }
            
            # Create a unique GUID for the folder
            $guid = [System.Guid]::NewGuid().ToString("B")
            
            # Create the registry entry
            $newKey = New-Item -Path "$registryPath\$guid" -Force
            Set-ItemProperty -Path "$registryPath\$guid" -Name "Name" -Value $folderName
            Set-ItemProperty -Path "$registryPath\$guid" -Name "Path" -Value $folderPath
            
            Write-Log "Added Quick Access entry: $folderName -> $folderPath"
        }
        
        # Also set the default Quick Access locations for new user profiles
        $defaultUserPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        if (!(Test-Path -Path $defaultUserPath)) {
            New-Item -Path $defaultUserPath -Force | Out-Null
        }
        
        # Set the Quick Access locations for the default user profile
        $quickAccessValue = $FoldersToPin -join "|"
        Set-ItemProperty -Path $defaultUserPath -Name "Quick Access" -Value $quickAccessValue -ErrorAction SilentlyContinue
        
        Write-Log "Successfully configured Quick Access for all users"
        
        # Restart Explorer to apply changes immediately
        Write-Log "Restarting Explorer to apply Quick Access changes..."
        Stop-Process -Name "explorer" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Start-Process "explorer"
        
    } catch {
        Write-Log "Error setting Quick Access for all users: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# Initialize foldersToPin array
$foldersToPin = @()

# Add local user folders FIRST (like Downloads)
if (-not [string]::IsNullOrWhiteSpace($LocalFolders)) {
    $localUserFolders = $LocalFolders -split ',' | ForEach-Object { $_.Trim() }
    $foldersToPin += $localUserFolders
}

# Add My Drive SECOND (if selected)
if ($PinMyDrive -eq 1) {
    $foldersToPin += "My Drive"
}

# Add shared drive folders LAST
if (-not [string]::IsNullOrWhiteSpace($SharedFolders)) {
    $sharedDriveFolders = $SharedFolders -split ',' | ForEach-Object { $_.Trim() }
    $foldersToPin += $sharedDriveFolders
}

# If no folders were provided, log and exit
if (-not $foldersToPin) {
    Write-Log "No folders specified to pin. Exiting script." -Level "WARN"
    return
}

Write-Log "Configuring Quick Access for the following folders: $($foldersToPin -join ', ')"

# Apply Quick Access settings for all users
try {
    Set-QuickAccessForAllUsers -FoldersToPin $foldersToPin
    Write-Log "Quick Access configuration completed successfully"
} catch {
    Write-Log "Failed to configure Quick Access: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}
