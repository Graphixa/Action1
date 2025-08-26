<#
Title: [Script Name] - [Brief Action Description] (e.g "Set User Permissions - Manage Local Group Memberships")

.SYNOPSIS
    Brief one-line description of what the script does.

.DESCRIPTION
    Detailed description of the script's purpose and functionality.
    Include key features and any important behavioral notes.
    Example: This script installs and configures XYZ software across the network and handles installation, configuration, and cleanup automatically.

.PARAMETER Action
    Action to perform. Example: "Add" or "Remove"
    Required: Yes
    Validation: Must be one of ["Add", "Remove"]

.PARAMETER UserList
    Comma-separated list of users to process (e.g. "john_doe, jane_smith")
    Required: Yes
    Validation: Cannot be empty, must contain valid usernames

.NOTES
    Required Action1 Permissions:
        - Run as System/Admin
        - File System Access (if needed)
        - Registry Access (if needed)

    Action1 Configuration:
        Required Parameters:
            - Name: "Action"
              Type: String
              Options: ["Add", "Remove"]
            
            - Name: "User List"
              Type: String
              Format: Comma-separated usernames
#>

$ProgressPreference = 'SilentlyContinue'

# ================================
# Action1 Parameters
# ================================

# Parameters passed from Action1 platform
$Action = ${Action}              # Example: "Add" or "Remove"
$UserList = ${User List}         # Example: "john_doe, jane_smith"


# Script Constants
$LogFilePath = "$env:SystemDrive\LST\Action1.log"

# ================================
# Logging Function: Write-Log
# ================================
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

# ================================
# Parameter Validation
# ================================
# Validate Action1 parameters
if ([string]::IsNullOrWhiteSpace($Action)) {
    Write-Log "Action parameter is required" -Level "ERROR"
    exit 1
}

if ([string]::IsNullOrWhiteSpace($UserList)) {
    Write-Log "User List parameter is required" -Level "ERROR"
    exit 1
}

# Convert comma-separated string to array and trim whitespace
$Users = $UserList.Split(',').Trim()

# Validate we have users after splitting
if ($Users.Count -eq 0 -or ($Users.Count -eq 1 -and [string]::IsNullOrWhiteSpace($Users[0]))) {
    Write-Log "No valid users provided in User List" -Level "ERROR"
    exit 1
}

# ================================
# Pre-Check Section
# ================================
try {
    Write-Log "Starting pre-checks..." -Level "INFO"
    
    # Example: Check if users exist
    foreach ($User in $Users) {
        # Add your validation logic here
        Write-Log "Validating user: $User" -Level "INFO"
    }
} 
catch {
    Write-Log "Pre-check failed: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}

# ================================
# Main Script Logic
# ================================
try {
    Write-Log "Starting main script execution..." -Level "INFO"
    Write-Log "Action: $Action" -Level "INFO"
    Write-Log "Processing users: $($Users -join ', ')" -Level "INFO"

    # Add your main script logic here
    foreach ($User in $Users) {
        Write-Log "Processing user: $User" -Level "INFO"
        # Add your per-user logic here
    }
} 
catch {
    Write-Log "Script failed: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}

Write-Log "Script completed successfully" -Level "INFO"
Exit 0