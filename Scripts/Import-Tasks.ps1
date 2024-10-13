# ================================================
# Import Scheduled Tasks Script for Action1
# ================================================
# Description:
#   - This script imports specified task files (XML) into Windows Task Scheduler.
#   - The task files can be sourced from a local share, GitHub repository, or other services with direct download links.
#   - Tasks are imported directly into Task Scheduler.
#   - Temporary XML files are cleaned up after import.
#
# Requirements:
#   - Admin rights are required.
#   - Task files should be in XML format.
# ================================================

$ProgressPreference = 'SilentlyContinue'

$taskFiles = @("${Import Task Path}") # Can be URL, Local Path, or //Network_Share. Separate multiple task file paths by commas (,)
$tempTaskFolder = "$env:TEMP\Action1Tasks" # Declare temp folder globally

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
# Function: Import Task File
# ================================
function Import-Task {
    param (
        [string]$taskFile    # Path or URL to the task file (XML)
    )

    # Ensure temp folder for task files exists
    if (-not (Test-Path $tempTaskFolder)) {
        New-Item -Path $tempTaskFolder -ItemType Directory -Force | Out-Null
    }

    # Extract file name and remove the .xml extension
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($taskFile)
    $tempTaskFile = Join-Path $tempTaskFolder "$fileName.xml"

    try {
        # Check if it's a remote URL or a local/network path
        if ($taskFile -match "^https?://") {
            Write-Log "Downloading task file from remote URL: $taskFile" -Level "INFO"
            Invoke-WebRequest -Uri $taskFile -OutFile $tempTaskFile -ErrorAction Stop
        } elseif (Test-Path $taskFile) {
            Write-Log "Copying task file from local/network path: $taskFile" -Level "INFO"
            Copy-Item -Path $taskFile -Destination $tempTaskFile -Force
        } else {
            Write-Log "The task file does not exist or is inaccessible: $taskFile" -Level "ERROR"
            return
        }

        # Import the task into Task Scheduler
        Write-Log "Importing task into Task Scheduler with name: $fileName" -Level "INFO"
        Register-ScheduledTask -TaskName "Action1-$fileName" -Xml (Get-Content $tempTaskFile | Out-String) -Force | Out-Null

        Write-Log "Successfully imported task: $fileName" -Level "INFO"
    } catch {
        Write-Log "Failed to import task: $($_.Exception.Message)" -Level "ERROR"
    }
}

# ================================
# Main Script Logic
# ================================

try {
    foreach ($taskFile in $taskFiles) {
        Import-Task -taskFile $taskFile
    }

    Write-Log "Scheduled task(s) import complete." -Level "INFO"
} catch {
    Write-Log "An error occurred while importing scheduled tasks: $($_.Exception.Message)" -Level "ERROR"
} finally {
    # Clean up the temp folder
    try {
        if (Test-Path $tempTaskFolder) {
            Remove-Item -Path $tempTaskFolder -Recurse -Force
            Write-Log "Temporary task files cleaned up." -Level "INFO"
        }
    } catch {
        Write-Log "Failed to clean up temp folder: $($_.Exception.Message)" -Level "ERROR"
    }
}
