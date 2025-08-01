# ================================================
# Get Action1 Log Script for Action1
# ================================================
# Description:
#   - This script reads the Action1 log file from the remote machine
#   - Displays the log contents to the host's screen
#   - Handles missing log files and access errors
#
# Requirements:
#   - Admin rights are required to access system drive.
# ================================================

$ProgressPreference = 'SilentlyContinue'

# Global Variables
$LogFilePath = "$env:SystemDrive\LST\Action1.log"

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = $LogFilePath, # Default log file path    
        [string]$Level = "INFO"  # Log level: INFO, WARN, ERROR
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Use color coding for console output based on level
    switch ($Level) {
        "ERROR" { Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor Red -BackgroundColor Black }
        "WARN"  { Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor Yellow }
        "INFO"  { Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor White }
        default { Write-Host "[$timestamp] [$Level] $Message" }
    }
}

function Get-Action1LogContent {
    try {
        # Check if log file exists
        if (-not (Test-Path -Path $LogFilePath)) {
            Write-Log "Log file not found at: $LogFilePath" -Level "WARN"
            return $false
        }

        # Get file info
        $logFile = Get-Item -Path $LogFilePath
        $logSize = [math]::Round($logFile.Length / 1MB, 2)
        $lastWriteTime = $logFile.LastWriteTime

        # Display log file information
        Write-Log "Log File Details:" -Level "INFO"
        Write-Log "  Path: $LogFilePath" -Level "INFO"
        Write-Log "  Size: $logSize MB" -Level "INFO"
        Write-Log "  Last Modified: $lastWriteTime" -Level "INFO"
        Write-Log "----------------------------------------" -Level "INFO"

        # Read and display log content
        Write-Log "Log Contents:" -Level "INFO"
        Write-Log "----------------------------------------" -Level "INFO"
        Get-Content -Path $LogFilePath | ForEach-Object {
            # Parse log line to extract level for color coding
            if ($_ -match '\[(INFO|WARN|ERROR)\]') {
                $level = $matches[1]
                Write-Log $_ -Level $level
            } else {
                Write-Host $_
            }
        }
        Write-Log "----------------------------------------" -Level "INFO"
        return $true
    }
    catch {
        Write-Log "Failed to read log file: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# Main execution
try {
    Write-Log "Starting Action1 log retrieval" -Level "INFO"
    
    $success = Get-Action1LogContent
    if (-not $success) {
        throw "Failed to retrieve Action1 log content"
    }

    Write-Log "Log retrieval completed successfully" -Level "INFO"
    exit 0

} catch {
    Write-Log $_.Exception.Message -Level "ERROR"
    exit 1
} 