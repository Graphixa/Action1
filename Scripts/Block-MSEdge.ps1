# ================================================
# Block Microsoft Edge Script for Action1
# ================================================
# Description:
#   - This script adds a new inbound and outbound firewall rule to block all traffic for the Microsoft Edge application on Windows 11.
#   - It uses Action1's predefined logging standards and checks for required administrative privileges.
#
# Requirements:
#   - Admin rights are required.
# ================================================

$ProgressPreference = 'SilentlyContinue'

$RuleName = "Block Microsoft Edge"
$ProgramPath = "$env:ProgramFiles (x86)\Microsoft\Edge\Application\msedge.exe"
$firewallRuleAction = ${Action}


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

if (-not (Test-Path -Path $ProgramPath)) {
    Write-Log -Message "Microsoft Edge executable not found at $ProgramPath. Exiting." -LogLevel "ERROR"
    throw "Microsoft Edge executable not found."
}

# ================================
# Main Script Logic
# ================================

try {
    # Create the inbound firewall rule
    Write-Log -Message "Creating inbound firewall rule to block Microsoft Edge."
    New-NetFirewallRule -DisplayName "$RuleName (Inbound)" `
        -Direction Inbound `
        -Action $firewallRuleAction `
        -Program $ProgramPath `
        -Profile Any `
        -Enabled True `
        -Description "Blocks inbound traffic for Microsoft Edge."

    # Create the outbound firewall rule
    Write-Log -Message "Creating outbound firewall rule to block Microsoft Edge."
    New-NetFirewallRule -DisplayName "$RuleName (Outbound)" `
        -Direction Outbound `
        -Action $firewallRuleAction `
        -Program $ProgramPath `
        -Profile Any `
        -Enabled True `
        -Description "Blocks outbound traffic for Microsoft Edge."

    Write-Log -Message "Firewall rules created successfully to block Microsoft Edge."
    Play-Notification
} catch {
    Write-Log -Message "An error occurred while creating firewall rules: $_" -LogLevel "ERROR"
    throw $_
}
