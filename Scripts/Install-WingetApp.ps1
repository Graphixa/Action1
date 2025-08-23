# ================================================
# Install Winget Application by ID Script for Action1
# ================================================
# Description:
#   - This script installs a single application using Winget based on the provided App ID.
#   - The application can be installed by specifying the Winget App ID and an optional version.
#   - Enhanced with proper error handling, installation validation, and exit code checking.
#
# Requirements:
#   - Admin rights are required.
# ================================================


$ProgressPreference = 'SilentlyContinue'

$wingetAppID = ${App ID}
$version = ${Version} # leave blank to install the latest version
$LogFilePath = "$env:SystemDrive\LST\Action1.log" # Default log file path

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = $LogFilePath, # Default log file path
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


function Get-WinGetExecutable {
    $winget = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq 'Microsoft.DesktopAppInstaller' }

    if ($null -ne $winget) {
        $wingetFilePath = Join-Path -Path $($winget.InstallLocation) -ChildPath 'winget.exe'
        if (Test-Path $wingetFilePath) {
            return Get-Item -Path $wingetFilePath
        }
    }
    
    Write-Log 'The WinGet executable is not detected, please proceed to Microsoft Store to update the Microsoft.DesktopAppInstaller application.' -Level "ERROR"
    return $false
}

function Test-ApplicationInstalled {
    param (
        [string]$AppID
    )
    
    try {
        $wingetExe = Get-WinGetExecutable
        if (-not $wingetExe) { return $false }
        
        $result = & $wingetExe.FullName list --id $AppID --exact 2>&1
        $exitCode = $LASTEXITCODE
        
        if ($exitCode -eq 0) {
            # Check if the app appears in the list output
            $appFound = $result | Where-Object { $_ -match $AppID }
            return $null -ne $appFound
        }
        return $false
    } catch {
        Write-Log "Error checking if application is installed: $($_.Exception.Message)" -Level "WARN"
        return $false
    }
}

# ================================
# Main Script Logic
# ================================

# Input validation
if ([string]::IsNullOrWhiteSpace($wingetAppID)) {
    Write-Log "Missing required variable: App ID" -Level "ERROR"
    Write-Output "ERROR: App ID is required"
    exit 1
}

try {
    Write-Log "Starting Winget application installation process for App ID: $wingetAppID" -Level "INFO"

    # Step 1: Check for WinGet executable
    $wingetExe = Get-WinGetExecutable
    if (-not $wingetExe) {
        throw "Winget executable not found. Please install Microsoft.DesktopAppInstaller from the Microsoft Store."
    }

    # Step 2: Check if app is already installed
    if (Test-ApplicationInstalled -AppID $wingetAppID) {
        Write-Log "Application $wingetAppID is already installed. Skipping installation." -Level "INFO"
        Write-Output "Application $wingetAppID is already installed."
        exit 0
    }

    # Step 3: Install the application via Winget
    Write-Log "Installing application via Winget using the App ID: $wingetAppID" -Level "INFO"

    # Build winget command with proper parameter ordering
    $wingetArgs = @(
        "install",
        "--id", $wingetAppID,
        "--exact",
        "--accept-package-agreements",
        "--accept-source-agreements",
        "--silent",
        "--scope", "machine"
    )
    
    if ($version -and -not [string]::IsNullOrWhiteSpace($version)) {
        $wingetArgs += "--version", $version
        Write-Log "Specified version: $version" -Level "INFO"
    } else {
        Write-Log "No version specified. Installing the latest version." -Level "INFO"
    }

    Write-Log "Executing winget command: winget $($wingetArgs -join ' ')" -Level "INFO"
    
    # Execute winget command
    $result = & $wingetExe.FullName @wingetArgs 2>&1
    $exitCode = $LASTEXITCODE
    
    # Log the output
    if ($result) {
        Write-Log "Winget output: $($result -join "`r`n")" -Level "INFO"
    }
    
    # Check exit code - winget uses specific exit codes
    switch ($exitCode) {
        0 { 
            Write-Log "Winget command completed successfully (exit code: $exitCode)" -Level "INFO" 
        }
        17002 { 
            throw "Winget: No application found with ID: $wingetAppID" 
        }
        17003 { 
            throw "Winget: Multiple applications found with ID: $wingetAppID" 
        }
        17004 { 
            throw "Winget: Application installation failed" 
        }
        default { 
            throw "Winget command failed with exit code: $exitCode" 
        }
    }

    # Step 4: Verify installation
    Write-Log "Verifying application installation..." -Level "INFO"
    Start-Sleep -Seconds 5  # Give winget time to complete
    
    if (Test-ApplicationInstalled -AppID $wingetAppID) {
        Write-Log "Application $wingetAppID has been successfully installed and verified." -Level "INFO"
        Write-Output "SUCCESS: Application $wingetAppID has been installed successfully."
    } else {
        throw "Application installation verification failed. The app was not found after installation."
    }

    Write-Log "Winget application installation complete for App ID: $wingetAppID" -Level "INFO"
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Log "An error occurred during the Winget installation: $errorMessage" -Level "ERROR"
    Write-Output "ERROR: $errorMessage"
    exit 1
}
