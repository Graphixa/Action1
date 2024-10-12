$ProgressPreference = 'SilentlyContinue'

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

$qa = New-Object -ComObject shell.application
$quickAccessFolder = $qa.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}')
$items = $quickAccessFolder.Items()


# Array of folders to pin
$foldersToPin = @(
    "G:\My Drive",
    "G:\Shared drives\Management",
    "G:\Shared drives\Admin",
    "G:\Shared drives\Clinical",
    "G:\Shared drives\Document Library",
    "G:\Shared drives\Creative",
    "G:\Shared drives\HR Drive",
    "G:\Shared drives\IT"
)

# ================================
# Main Script Logic
# ================================
# Add your main logic for downloading, installing, configuring, etc.

try {
    Write-Log "Executing main script logic..." -Level "INFO"
    # Unpin all existing folders
    foreach ($item in $items) {
        $item.InvokeVerb('unpinfromhome')
    }

    # Pin new folders to Quick Access
    foreach ($folder in $foldersToPin) {
        $QuickAccess = New-Object -ComObject shell.application
        $QuickAccess.Namespace($folder).Self.InvokeVerb("pintohome")
    }

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
