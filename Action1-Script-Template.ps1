# ================================================
# PowerShell Script Template for Action1
# ================================================
# Description:
#   - Provide a brief description of the script's purpose.
#   - Example: This script installs and configures the XYZ software.
#
# Requirements:
#   - List any prerequisites or requirements (e.g., internet access, admin rights).
# ================================================

$ProgressPreference = 'SilentlyContinue'

# ================================
# Logging Function: Write-Log
# ================================
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

# ================================
# Parameters Section (Customizable)
# ================================
# Define any custom parameters here.

$softwareName = 'SoftwareName'  # Placeholder for software name
$installPath = "$env:SystemDrive\Program Files\$softwareName"
$tempPath = "$env:SystemDrive\Temp\"

# ================================
# Pre-Check Section (Optional)
# ================================
# Use this section for pre-checks.
# Example: Check if the software is already installed, exit if no action is required.

try {
    # Example pre-check (modify as needed)
    Write-Log "Performing pre-checks..." -Level "INFO"
    # Add your pre-check logic here.
} catch {
    Write-Log "Pre-check failed: $($_.Exception.Message)" -Level "ERROR"
    return
}

# ================================
# Main Script Logic
# ================================
# Add your main logic for downloading, installing, configuring, etc.

try {
    Write-Log "Executing main script logic..." -Level "INFO"
    # Add your main script logic here (e.g., downloading, installing).
} catch {
    Write-Log "An error occurred during the main script logic: $($_.Exception.Message)" -Level "ERROR"
    return
}

# ================================
# Cleanup Section
# ================================
# Clean up temporary files or logs here.

try {
    Write-Log "Cleaning up temporary files..." -Level "INFO"
    # Add your cleanup logic here (e.g., removing temp files).
} catch {
    Write-Log "Failed to clean up temporary files: $($_.Exception.Message)" -Level "ERROR"
}

# ================================
# End of Script
# ================================
