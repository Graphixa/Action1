# ================================================
# Install Missive Script for Action1
# ================================================
# Description:
#   - This script checks if Missive is installed. If not, it downloads the latest version and installs it.
#   - It also creates shortcuts for all users and cleans up temporary files afterward.
#
# Requirements:
#   - Admin rights are required.
#   - Internet access is required to download the installer.
# ================================================

$ProgressPreference = 'SilentlyContinue'

# Global Variables
$softwareName = 'Missive'
$downloadUrl = 'https://mail.missiveapp.com/download/win'
$InstallPath = "$env:SystemDrive\Missive"
$tempPath = "$env:temp"  # Changed to match reference script
$missiveFile = "$tempPath\MissiveSetup.exe"
$LogFilePath = "$env:SystemDrive\LST\Action1.log" # Default log file path

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = $LogFilePath, # Default log file path
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

    # Write output to Action1 host using Write-Host instead of Write-Output
    Write-Host $Message
}

function Remove-TempFiles {
    try {
        Write-Log "Cleaning up temporary files and folders..." -Level "INFO"
        if (Test-Path $missiveFile) {
            Remove-Item $missiveFile -Force -ErrorAction Stop
        }
        if (Test-Path $tempPath) {
            Remove-Item $tempPath -Force -ErrorAction Stop
        }
        Start-Sleep -Seconds 2
        Write-Log "Temporary files cleaned up successfully" -Level "INFO"
        return $true
    } catch {
        Write-Log "Failed to remove temporary files: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Test-Prerequisites {
    try {
        Write-Log "Checking if $softwareName is already installed..." -Level "INFO"
        
        # Method 1: Check all user profiles for HKCU installations
        $allUsersInstalled = $false
        $userProfiles = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" -ErrorAction SilentlyContinue
        
        foreach ($userProfile in $userProfiles) {
            try {
                $profilePath = Get-ItemProperty -Path $userProfile.PSPath -Name "ProfileImagePath" -ErrorAction SilentlyContinue
                if ($profilePath.ProfileImagePath -and (Test-Path $profilePath.ProfileImagePath)) {
                    $userSid = $userProfile.PSChildName
                    $userHKCU = "Registry::HKEY_USERS\$userSid\Software\Microsoft\Windows\CurrentVersion\Uninstall"
                    
                    $userInstalled = Get-ChildItem $userHKCU -Recurse -ErrorAction SilentlyContinue | 
                        Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue | 
                        Where-Object { $_.DisplayName -Match "^$softwareName.*" }
                    
                    if ($userInstalled) {
                        Write-Log "Found $softwareName installed for user profile: $($profilePath.ProfileImagePath)" -Level "INFO"
                        $allUsersInstalled = $true
                        break
                    }
                }
            } catch {
                # Continue checking other profiles if one fails
                continue
            }
        }
        
        # Method 2: Check machine-wide installations
        $machineUninstallPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
        $wow64UninstallPath = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
        
        $machineInstalled = Get-ChildItem $machineUninstallPath -Recurse -ErrorAction SilentlyContinue | 
            Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue | 
            Where-Object { $_.DisplayName -Match "^$softwareName.*" }
            
        $wow64Installed = Get-ChildItem $wow64UninstallPath -Recurse -ErrorAction SilentlyContinue | 
            Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue | 
            Where-Object { $_.DisplayName -Match "^$softwareName.*" }
        
        # Method 3: Check if executable exists in common installation paths
        $commonPaths = @(
            "$env:ProgramFiles\Missive\Missive.exe",
            "$env:ProgramFiles(x86)\Missive\Missive.exe",
            "$env:SystemDrive\Missive\Missive.exe",
            "$env:LOCALAPPDATA\Programs\Missive\Missive.exe"
        )
        
        $executableExists = $false
        foreach ($path in $commonPaths) {
            if (Test-Path $path) {
                Write-Log "Found $softwareName executable at: $path" -Level "INFO"
                $executableExists = $true
                break
            }
        }
        
        # Method 4: Check current user (SYSTEM) uninstall info
        $currentUserInstalled = Get-ChildItem 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall' -Recurse -ErrorAction SilentlyContinue | 
            Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue | 
            Where-Object { $_.DisplayName -Match "^$softwareName.*" }
        
        # Determine if software is installed
        if ($allUsersInstalled -or $machineInstalled -or $wow64Installed -or $executableExists -or $currentUserInstalled) {
            Write-Log "$softwareName is already installed on this system" -Level "INFO"
            return $false
        }
        
        Write-Log "$softwareName is NOT installed on this system" -Level "INFO"
        return $true
        
    } catch {
        Write-Log "Failed to check installation status: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Get-MissiveInstaller {
    try {
        Write-Log "Downloading Missive installer from $downloadUrl" -Level "INFO"
        
        # Create Temp Folder if it doesn't exist
        if (-not (Test-Path $tempPath)) {
            [void](New-Item -ItemType Directory -Force -Path $tempPath)
        }
        
        # Remove existing installer if present
        if (Test-Path $missiveFile) {
            Remove-Item $missiveFile -Force -ErrorAction Stop
        }
        
        # Download installer
        Invoke-WebRequest -Uri $downloadUrl -OutFile $missiveFile -UseBasicParsing -ErrorAction Stop
        
        if (Test-Path $missiveFile) {
            Write-Log "Download completed successfully" -Level "INFO"
            return $true
        }
        throw "Download completed but installer not found"
    } catch {
        Write-Log "Failed to download installer: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Install-MissiveClient {
    try {
        Write-Log "Installing Missive for all users..." -Level "INFO"
        $process = Start-Process -FilePath $missiveFile -ArgumentList "/S /D=$InstallPath" -Wait -PassThru
        
        if ($process.ExitCode -ne 0) {
            throw "Installation failed with exit code: $($process.ExitCode)"
        }
        
        Start-Sleep -Seconds 5 # Wait for installation to complete
        Write-Log "Installation completed successfully" -Level "INFO"
        return $true
    } catch {
        Write-Log "Installation failed: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Set-MissiveShortcuts {
    try {
        Write-Log "Creating shortcuts for all users..." -Level "INFO"
        $missiveExecutable = "$InstallPath\Missive.exe"
        
        if (-not (Test-Path $missiveExecutable)) {
            throw "Missive executable not found at: $missiveExecutable"
        }

        $WScriptObj = New-Object -ComObject ("WScript.Shell")
        
        # Create Start Menu shortcut
        $startMenuPath = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Missive.lnk"
        $shortcutStart = $WScriptObj.CreateShortcut($startMenuPath)
        $shortcutStart.TargetPath = $missiveExecutable
        $shortcutStart.Save()
        
        # Create Desktop shortcut
        $desktopPath = "$env:Public\Desktop\Missive.lnk"
        $shortcutDesktop = $WScriptObj.CreateShortcut($desktopPath)
        $shortcutDesktop.TargetPath = $missiveExecutable
        $shortcutDesktop.Save()

        Write-Log "Shortcuts created successfully" -Level "INFO"
        return $true
    } catch {
        Write-Log "Failed to create shortcuts: $($_.Exception.Message)" -Level "WARN"
        return $false
    }
}

function Start-MissiveApp {
    try {
        $missiveExecutable = "$InstallPath\Missive.exe"
        if (Test-Path $missiveExecutable) {
            Start-Process $missiveExecutable
            Write-Log "Missive launched successfully" -Level "INFO"
            return $true
        }
        Write-Log "Cannot launch Missive - executable not found at: $missiveExecutable" -Level "WARN"
        return $false
    } catch {
        Write-Log "Failed to launch Missive: $($_.Exception.Message)" -Level "WARN"
        return $false
    }
}

# ================================
# Main Script Logic
# ================================

try {
    Write-Log "Starting Missive installation process" -Level "INFO"

    if (-not (Test-Prerequisites)) {
        Remove-TempFiles
        exit 0
    }

    Write-Log "$softwareName is NOT installed. Installing Now..." -Level "INFO"

    $success = Get-MissiveInstaller
    if (-not $success) {
        Remove-TempFiles
        throw "Failed to get Missive installer"
    }

    $success = Install-MissiveClient
    if (-not $success) {
        Remove-TempFiles
        throw "Failed to install Missive"
    }

    Set-MissiveShortcuts
    Start-MissiveApp

    # Final Cleanup
    Remove-TempFiles
    Write-Log "Missive deployment completed successfully" -Level "INFO"
    exit 0

} catch {
    Write-Log $_.Exception.Message -Level "ERROR"
    Remove-TempFiles
    exit 1
}