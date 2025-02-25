# ================================================
# Local User Passwords Never Expires Action1 Script
# ================================================
# Description:
#   - Sets the "Password Never Expires" flag for all local user accounts
#   - Excludes built-in system accounts from modification
#
# Requirements:
#   - Administrative privileges
#   - Windows OS
# ================================================

$ProgressPreference = 'SilentlyContinue'

$ExcludedUsers = @("Administrator", "DefaultAccount", "Guest")

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
    Write-Host "$Message"
}

# Main Script Logic
try {
    
    Write-Log "Setting local user passwords to never expire" -Level "INFO"

    # Get all local user accounts
    $Users = Get-LocalUser -ErrorAction Stop
    Write-Log "Retrieved $(($Users | Measure-Object).Count) local user accounts" -Level "INFO"

    foreach ($User in $Users) {
        if ($User.Name -notin $ExcludedUsers) {
            try {
                Write-Log "Processing user: $($User.Name)" -Level "INFO"
                
                # Check current setting
                if (-not $User.PasswordNeverExpires) {
                    Set-LocalUser -Name $User.Name -PasswordNeverExpires $true -ErrorAction Stop
                    Write-Log "Successfully set password to never expire for user: $($User.Name)" -Level "INFO"
                } else {
                    Write-Log "Password already set to never expire for user: $($User.Name)" -Level "INFO"
                }
            }
            catch {
                Write-Log "Error processing user $($User.Name): $_" -Level "ERROR"
            }
        }
        else {
            Write-Log "Skipping excluded user: $($User.Name)" -Level "INFO"
        }
    }

} catch {
    Write-Log "An error occurred during the main script logic: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}
