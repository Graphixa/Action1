# ================================================
# Chocolatey Package Installation Script for Action1
# ================================================
# Description:
#   - This script checks if Chocolatey is installed. If not, it installs Chocolatey.
#   - It then installs packages listed in a specified Chocolatey app manifest.
#
# Requirements:
#   - Admin rights are required.
#   - Internet access is required.
# ================================================

$ProgressPreference = 'SilentlyContinue'

$ChocolateyAppManifest = ${App Manifest Link}
$tempPath = "$env:Temp"  # Temporary path for script use


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


function Install-Chocolatey {
    Write-Log "Installing Chocolatey Package Manager..." -Level "INFO"
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Write-Log "Chocolatey Package Manager installed successfully." -Level "INFO"
        return $true
    } catch {
        Write-Log "Failed to install Chocolatey Package Manager: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# ================================
# Pre-Checks
# ================================
try {
    Write-Log "Performing pre-checks for Chocolatey installation..." -Level "INFO"
    
    # Check if Chocolatey is installed
    if (!(Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Log "Chocolatey is NOT installed. Attempting to install now..." -Level "WARN"
        if (-not (Install-Chocolatey)) {
            Write-Log "Failed to install Chocolatey. Exiting script with error code 1." -Level "ERROR"
            exit 1  # Exit with error code 1 for Action1
        }
    } else {
        Write-Log "Chocolatey is already installed." -Level "INFO"
    }
} catch {
    Write-Log "Pre-check failed: $($_.Exception.Message)" -Level "ERROR"
    exit 1  # Exit with error code 1 for Action1
}

# ================================
# Main Script Logic
# ================================
try {
    Write-Log "Executing main script logic for Chocolatey package installation..." -Level "INFO"
    
    # Download the Chocolatey App Manifest
    $manifestFile = "$tempPath\apps-chocolatey.config"
    try {
        Write-Log "Downloading Chocolatey app manifest from $ChocolateyAppManifest..." -Level "INFO"
        Invoke-WebRequest -Uri $ChocolateyAppManifest -OutFile $manifestFile
        Write-Log "Downloaded Chocolatey app manifest successfully." -Level "INFO"
    } catch {
        Write-Log "Failed to download Chocolatey app manifest: $($_.Exception.Message)" -Level "ERROR"
        exit 1  # Exit with error code 1 for Action1
    }

    # Install packages from the manifest
    try {
        Write-Log "Installing Chocolatey packages from the manifest..." -Level "INFO"
        & "C:\ProgramData\chocolatey\bin\choco.exe" install $manifestFile --yes
        Write-Log "Chocolatey packages installed successfully." -Level "INFO"
    } catch {
        Write-Log "Failed to install Chocolatey packages: $($_.Exception.Message)" -Level "ERROR"
        exit 1  # Exit with error code 1 for Action1
    }
} catch {
    Write-Log "An error occurred during script execution: $($_.Exception.Message)" -Level "ERROR"
    exit 1  # Exit with error code 1 for Action1
}

# ================================
# Cleanup Section
# ================================
try {
    Write-Log "Cleaning up temporary files..." -Level "INFO"
    Remove-Item -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Temporary files cleaned up successfully." -Level "INFO"
} catch {
    Write-Log "Failed to clean up temporary files: $($_.Exception.Message)" -Level "ERROR"
}

# ================================
# End of Script
# ================================
Write-Log "Script execution completed." -Level "INFO"
exit 0  # Exit with success code 0 for Action1
