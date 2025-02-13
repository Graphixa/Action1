# ================================================
# PowerShell Script Template for Action1
# ================================================
# Description:
#   - This script installs Chocolatey Package Manager.
#
# Requirements:
#   - Internet access.
#   - Administrator rights.
# ================================================

$ProgressPreference = 'SilentlyContinue'


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

$tempPath = "$env:SystemDrive\Temp\"


Write-Log "Chocolatey Pre-Installation Check..." -Level "INFO"

# Check if Chocolatey is already installed
if (Get-Command choco -ErrorAction SilentlyContinue) {
    Write-Log "Chocolatey is already installed. No further action required." -Level "INFO"
    exit
}

function Install-Chocolatey {
    Write-Log "Installing Chocolatey Package Manager..." -Level "INFO"
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Write-Log "Chocolatey Package Manager installed successfully." -Level "INFO"
    } catch {
        Write-Log "Failed to install Chocolatey Package Manager: $($_.Exception.Message)" -Level "ERROR"
        return
    }
}

# ================================
# Main Script Logic
# ================================

try {
    Install-Chocolatey
} catch {
    Write-Log "An error occurred during Chocolatey installation: $($_.Exception.Message)" -Level "ERROR"
    return
}


# Cleanup Temporary Files
try {
    Write-Log "Cleaning up temporary files..." -Level "INFO"
    if (Test-Path $tempPath) {
        Remove-Item -Path $tempPath -Recurse -Force
        Write-Log "Temporary files cleaned up successfully." -Level "INFO"
    }
} catch {
    Write-Log "Failed to clean up temporary files: $($_.Exception.Message)" -Level "ERROR"
}
