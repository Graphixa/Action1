# ================================================
# PowerShell Script for Action1: Install Lithnet Idle Logoff
# ================================================
# Description:
#   - Downloads and installs Lithnet Idle Logoff MSI
#   - Configures registry settings for idle logoff functionality
#   - Sets up 9-hour idle timeout with warning system
#
# Requirements:
#   - Administrative privileges
#   - Internet access for downloading MSI
#   - Windows OS
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
# Pre-Check Section
# ================================
try {
    Write-Log "Starting Lithnet Idle Logoff installation..." -Level "INFO"
    
    # Verify administrative privileges
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Log "Script requires administrative privileges." -Level "ERROR"
        exit 1
    }

    # Check if already installed
    $installed = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*Lithnet Idle Logoff*" }
    if ($installed) {
        Write-Log "Lithnet Idle Logoff is already installed. Skipping installation." -Level "INFO"
        exit 0
    }

} catch {
    Write-Log "Pre-check failed: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}

# ================================
# Main Script Logic
# ================================
try {
    Write-Log "Installing Lithnet Idle Logoff" -Level "INFO"

    # Download and install MSI
    $msiUrl = "https://github.com/lithnet/idle-logoff/releases/latest/download/lithnet.idlelogoff.setup.msi"
    $msiPath = Join-Path -Path $env:TEMP -ChildPath "lithnet.idlelogoff.setup.msi"

    try {
        Write-Log "Downloading Lithnet Idle Logoff MSI..." -Level "INFO"
        Invoke-WebRequest -Uri $msiUrl -OutFile $msiPath -UseBasicParsing
        Write-Log "Successfully downloaded Lithnet Idle Logoff MSI" -Level "INFO"
    }
    catch {
        Write-Log "Failed to download Lithnet Idle Logoff MSI: $($_.Exception.Message)" -Level "ERROR"
        exit 1
    }

    # Install MSI silently
    Write-Log "Installing Lithnet Idle Logoff MSI..." -Level "INFO"
    $installResult = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /qn" -Wait -PassThru
    if ($installResult.ExitCode -eq 0) {
        Write-Log "Successfully installed Lithnet Idle Logoff" -Level "INFO"
    }
    else {
        Write-Log "Failed to install Lithnet Idle Logoff. Exit code: $($installResult.ExitCode)" -Level "ERROR"
        exit 1
    }

    # Configure registry settings
    Write-Log "Configuring registry settings..." -Level "INFO"
    
    # Create registry path if it doesn't exist
    $registryPath = "HKLM:\SOFTWARE\Lithnet\IdleLogOff"
    if (!(Test-Path -Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }

    # Set registry values for Idle Logoff
    Set-ItemProperty -Path $registryPath -Name "Enabled" -Value 1 -Type DWord
    Set-ItemProperty -Path $registryPath -Name "IdleLimit" -Value 0x21C -Type DWord # 540 minutes (9 hours)
    Set-ItemProperty -Path $registryPath -Name "IgnoreDisplayRequested" -Value 1 -Type DWord
    Set-ItemProperty -Path $registryPath -Name "WarningEnabled" -Value 1 -Type DWord
    Set-ItemProperty -Path $registryPath -Name "WarningMessage" -Value "Your session has been idle for more than 7 hours, and you will be logged out in {0}" -Type String
    Set-ItemProperty -Path $registryPath -Name "WarningPeriod" -Value 0x78 -Type DWord # 120 seconds
    Set-ItemProperty -Path $registryPath -Name "WaitForInitialInput" -Value 0 -Type DWord
    Set-ItemProperty -Path $registryPath -Name "Action" -Value 0 -Type DWord

    Write-Log "Successfully configured Lithnet Idle Logoff registry settings" -Level "INFO"

    # Set registry setting to enable fast user switching in windows 11
    $registryPolicyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    if (!(Test-Path -Path $registryPolicyPath)) {
        New-Item -Path $registryPolicyPath -Force | Out-Null
    }
    
    Set-ItemProperty -Path $registryPolicyPath -Name "HideFastUserSwitching" -Value 0 -Type DWord
    Write-Log "Successfully enabled fast user switching in windows 11" -Level "INFO"


} catch {
    Write-Log "An error occurred during the main script logic: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}

# ================================
# Cleanup Section
# ================================
try {
    Write-Log "Cleaning up temporary files..." -Level "INFO"
    if (Test-Path $msiPath) {
        Remove-Item -Path $msiPath -Force
        Write-Log "Removed temporary MSI file" -Level "INFO"
    }
} catch {
    Write-Log "Failed to clean up temporary files: $($_.Exception.Message)" -Level "ERROR"
}

# ================================
# End of Script
# ================================
Write-Log "Script completed successfully" -Level "INFO"
exit 0 