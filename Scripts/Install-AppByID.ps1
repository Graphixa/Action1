# ================================================
# Install App by ID (Winget) Script for Action1
# ================================================
# Description:
#   - This script installs a single application using Winget based on the provided App ID.
#   - The application can be installed by specifying the Winget App ID and an optional version.
#
# Requirements:
#   - Admin rights are required.
# ================================================


$ProgressPreference = 'SilentlyContinue'

$wingetAppID = ${App ID}
$version = ${Version}


function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = "$env:SystemDrive\LST\Action1.log", # Default log file path
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


function Get-WinGetExecutable {
    $winget = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq 'Microsoft.DesktopAppInstaller' }

    if ($null -ne $winget) {
        $wingetFilePath = Join-Path -Path $($winget.InstallLocation) -ChildPath 'winget.exe'
        $wingetFile = Get-Item -Path $wingetFilePath
        return $wingetFile
    } else {
        Write-Log 'The WinGet executable is not detected, please proceed to Microsoft Store to update the Microsoft.DesktopAppInstaller application.' -Level "WARN"
        return $false
    }
}

# ================================
# Main Script Logic
# ================================
try {
    Write-Log "Starting Winget application installation process for App ID: $wingetAppID" -Level "INFO"

    # Step 1: Check for WinGet executable
    $wingetExe = Get-WinGetExecutable
    if (-not $wingetExe) {
        throw "Winget executable not found. Exiting script."
    }

    # Step 2: Install the application via Winget
    try {
        Write-Log "Installing application via Winget using the App ID: $wingetAppID" -Level "INFO"

        # Set up the version flag
        $versionOption = ""
        if ($version) {
            $versionOption = "--version $version"
            Write-Log "Specified version: $version" -Level "INFO"
        } else {
            Write-Log "No version specified. Installing the latest version." -Level "INFO"
        }

        # Install the app
        & $wingetExe.FullName install --id $wingetAppID -h --accept-package-agreements --accept-source-agreements $versionOption --verbose-logs

        Write-Log "Application installed successfully via Winget: $wingetAppID, Version: $version" -Level "INFO"
    } catch {
        Write-Log "Error occurred during Winget application installation: $($_.Exception.Message)" -Level "ERROR"
        throw
    }

    Write-Log "Winget application installation complete for App ID: $wingetAppID" -Level "INFO"
}
catch {
    Write-Log "An error occurred during the Winget setup: $($_.Exception.Message)" -Level "ERROR"
}
