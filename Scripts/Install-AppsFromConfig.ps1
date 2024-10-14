# ================================================
# Winget Import Script for Action1
# ================================================
# Description:
#   - This script downloads and imports a Winget configuration file to install apps.
#   - The configuration file can be sourced from a local path, network share, or URL.
#   - The configuration file is validated to ensure it is a valid Winget import file.
#
# Requirements:
#   - Admin rights are required.
#   - Configuration file must follow Winget schema and be in JSON format.
# ================================================

$ProgressPreference = 'SilentlyContinue'

$wingetConfigPath = "${Winget Configuration Path}"  # Provide URL or path for the Winget configuration file
$downloadLatestVersions = ${Download Latest Versions}  # Boolean: 1 to download latest versions or 0 to download version info in configuration file

$downloadLocation = "$env:temp\winget-import"  # Path to store the Winget config file temporarily, e.g., "$env:SystemDrive\Action1"


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
# Function: Download or Access Winget Configuration File
# ================================
function Get-WingetConfigFile {
    param (
        [string]$configPath,   # URL or path to the configuration file
        [string]$fileName      # File name for saving the configuration file locally
    )

    if (-not (Test-Path $downloadLocation)) {
        New-Item -Path $downloadLocation -ItemType Directory -Force | Out-Null
    }

    $localConfigPath = Join-Path $downloadLocation $fileName

    try {
        if ($configPath -match "^https?://") {
            Write-Log "Downloading configuration file from URL: $configPath" -Level "INFO"
            Invoke-WebRequest -Uri $configPath -OutFile $localConfigPath -ErrorAction Stop
        } elseif (Test-Path $configPath) {
            Write-Log "Copying configuration file from local/network path: $configPath" -Level "INFO"
            Copy-Item -Path $configPath -Destination $localConfigPath -Force
        } else {
            throw "The configuration file does not exist or is inaccessible: $configPath"
        }

        return $localConfigPath
    } catch {
        Write-Log "Failed to download or copy configuration file: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# ================================
# Function: Validate Winget Configuration File
# ================================
function Validate-WingetConfig {
    param (
        [string]$configFile
    )

    try {
        $configContent = Get-Content -Path $configFile -Raw
        $configJson = $configContent | ConvertFrom-Json

        # Check if the file contains the expected "Packages" and "SourceDetails"
        if (-not $configJson.Sources.Packages) {
            throw "The configuration file is missing the 'Packages' section."
        }
        if (-not $configJson.Sources.SourceDetails) {
            throw "The configuration file is incorrect or malformed."
        }

        return $true
    } catch {
        Write-Log "Invalid Winget configuration file: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# ================================
# Main Script Logic
# ================================
try {
    Write-Log "Starting Winget application installation process." -Level "INFO"

    # Step 1: Download or access the configuration file
    $wingetConfigFile = Get-WingetConfigFile -configPath $wingetConfigPath -fileName "winget-config.json"

    # Step 2: Validate the Winget configuration file
    if (-not (Validate-WingetConfig -configFile $wingetConfigFile)) {
        throw "The configuration file is invalid. Exiting script."
    }

    # Step 3: Reset Winget sources and accept agreements
    try {
        Write-Log "Resetting Winget sources and accepting agreements." -Level "INFO"
        
        # Set flag for ignoring versions in config file
        $ignoreVersions = ""
        if ($downloadLatestVersions -eq 1) {
            $ignoreVersions = "--ignore-versions"
        }

        winget source reset --force
        winget import -i $wingetConfigFile --accept-package-agreements --accept-source-agreements $ignoreVersions

        Write-Log "Applications installed successfully via Winget." -Level "INFO"
    } catch {
        Write-Log "Error occurred during Winget application installation: $($_.Exception.Message)" -Level "ERROR"
        throw
    }

    Write-Log "Winget application installation complete." -Level "INFO"
}
catch {
    Write-Log "An error occurred during the Winget setup: $($_.Exception.Message)" -Level "ERROR"
}
