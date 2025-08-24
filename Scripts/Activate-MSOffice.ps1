# ================================================
# PowerShell Script for Office Activation via Action1
# ================================================
# Description:
#   - Activates the currently installed Microsoft Office suite
#   - Uses Office Software Protection Platform (OSPP.VBS)
#   - Supports Office 2013, 2016, 2019, and 2021
#
# Requirements:
#   - Microsoft Office must be installed
#   - Admin rights required
#   - Internet connectivity for activation
#   - Valid Office product key
# ================================================

$ProgressPreference = 'SilentlyContinue'

# ====================
# Parameters Section
# ====================
$ProductKey = ${Office Product Key}  # Action1 parameter for Office product key
$LogFilePath = "$env:SystemDrive\LST\Action1.log"


# ================================
# Logging Function: Write-Log
# ================================
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = $LogFilePath,
        [string]$Level = "INFO"
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
    
    Add-Content -Path $LogFilePath -Value $logMessage
    Write-Output "$Message"
}

# ================================
# Helper Functions
# ================================
function Find-OfficeInstallPath {
    $possiblePaths = @(
        "${env:ProgramFiles}\Microsoft Office\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
        "${env:ProgramFiles}\Microsoft Office\Office15",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15"
    )

    foreach ($path in $possiblePaths) {
        if (Test-Path -Path "$path\OSPP.VBS") {
            return $path
        }
    }
    return $null
}

function Get-OfficeVersion {
    param (
        [string]$OfficePath
    )

    if ($OfficePath -match "Office16") {
        # Could be 2016, 2019, or 2021
        return "2016/2019/2021"
    }
    elseif ($OfficePath -match "Office15") {
        return "2013"
    }
    return "Unknown"
}

# ================================
# Pre-Check Section
# ================================
try {
    Write-Log "Starting Office activation process..." -Level "INFO"

    # Validate product key parameter
    if (-not $ProductKey) {
        Write-Log "Product key parameter is required." -Level "ERROR"
        exit 1
    }

    # Check if Office is installed
    $officePath = Find-OfficeInstallPath
    if (-not $officePath) {
        Write-Log "Microsoft Office installation not found." -Level "ERROR"
        exit 1
    }

    $officeVersion = Get-OfficeVersion -OfficePath $officePath
    Write-Log "Found Microsoft Office $officeVersion installation at: $officePath" -Level "INFO"

    # Check if OSPP.VBS exists
    $osppPath = Join-Path $officePath "OSPP.VBS"
    if (-not (Test-Path $osppPath)) {
        Write-Log "OSPP.VBS not found at expected location: $osppPath" -Level "ERROR"
        exit 1
    }

} catch {
    Write-Log "Pre-check failed: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}

# ================================
# Main Script Logic
# ================================
try {
    # Check current license status
    Write-Log "Checking current license status..." -Level "INFO"
    $status = cscript //nologo "$osppPath" /dstatus
    
    # If already licensed, we're done
    if ($status -match "LICENSE STATUS:\s+Licensed") {
        Write-Log "Microsoft Office is already activated. No action needed." -Level "INFO"
        exit 0
    }
    
    # Install product key
    Write-Log "Installing product key..." -Level "INFO"
    $keyResult = cscript //nologo "$osppPath" /inpkey:$ProductKey
    if (-not ($keyResult -match "Product key installation successful")) {
        Write-Log "Failed to install product key: $keyResult" -Level "ERROR"
        exit 1
    }
    Write-Log "Product key installed successfully." -Level "INFO"
    
    # Attempt activation
    Write-Log "Attempting to activate Office..." -Level "INFO"
    $activationResult = cscript //nologo "$osppPath" /act 2>&1
    Write-Log "Activation output: $activationResult" -Level "INFO"
    
    # Verify activation
    $finalStatus = cscript //nologo "$osppPath" /dstatus
    if ($finalStatus -match "LICENSE STATUS:\s+Licensed") {
        Write-Log "Microsoft Office activated successfully!" -Level "INFO"
        exit 0
    } else {
        Write-Log "Activation failed. See activation output above for details." -Level "ERROR"
        exit 1
    }

} catch {
    Write-Log "An error occurred: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}