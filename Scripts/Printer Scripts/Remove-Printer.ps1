# ================================================
# Printer Removal Script for Action1
# ================================================
# Description:
#   - This script removes a printer and its associated port from the system.
#   - It can optionally remove the printer driver if no other printers are using it.
#   - Enhanced with robust error handling and logging.
#
# Requirements:
#   - Admin rights are required.
#

# ================================================

$ProgressPreference = 'SilentlyContinue'

# Action1 Variables
$PrinterName = ${Printer Name} # Name of the printer to remove
$RemoveDriver = ${Remove Driver} # Optional: Set to "true" to also remove the driver if unused
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

# ================================
# Main Script Logic
# ================================

# Quick sanity check on inputs
if ([string]::IsNullOrWhiteSpace($PrinterName)) {
    Write-Log "Missing required variable: PrinterName" -Level "ERROR"
    return
}

# Check if printer exists
try {
    $printer = Get-Printer -Name $PrinterName -ErrorAction Stop
    Write-Log "Found printer: $PrinterName" -Level "INFO"
} catch {
    Write-Log "Printer '$PrinterName' not found. Nothing to remove." -Level "INFO"
    Write-Output "Printer '$PrinterName' not found. Nothing to remove."
    return
}

# Store printer details before removal
$printerPort = $printer.PortName
$printerDriver = $printer.DriverName
$isDefault = $printer.PrinterStatus -eq "Idle" -and (Get-WmiObject -Class Win32_Printer -Filter "Name='$PrinterName'").Default

Write-Log "Printer details - Port: $printerPort, Driver: $printerDriver, Default: $isDefault" -Level "INFO"

# Remove the printer
try {
    Remove-Printer -Name $PrinterName -ErrorAction Stop
    Write-Log "Successfully removed printer: $PrinterName" -Level "INFO"
} catch {
    Write-Log "Failed to remove printer '$PrinterName': $($_.Exception.Message)" -Level "ERROR"
    return
}

# Remove the printer port if it's no longer used
try {
    $remainingPrinters = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.PortName -eq $printerPort }
    if (-not $remainingPrinters) {
        try {
            Remove-PrinterPort -Name $printerPort -ErrorAction Stop
            Write-Log "Removed unused printer port: $printerPort" -Level "INFO"
        } catch {
            Write-Log "Could not remove printer port '$printerPort': $($_.Exception.Message)" -Level "WARN"
        }
    } else {
        Write-Log "Printer port '$printerPort' is still in use by other printers. Keeping port." -Level "INFO"
    }
} catch {
    Write-Log "Could not check printer port usage: $($_.Exception.Message)" -Level "WARN"
}

# Optionally remove the driver if no other printers are using it
if ($RemoveDriver -eq "true") {
    try {
        $remainingPrinters = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.DriverName -eq $printerDriver }
        if (-not $remainingPrinters) {
            try {
                Remove-PrinterDriver -Name $printerDriver -ErrorAction Stop
                Write-Log "Removed unused printer driver: $printerDriver" -Level "INFO"
            } catch {
                Write-Log "Could not remove printer driver '$printerDriver': $($_.Exception.Message)" -Level "WARN"
            }
        } else {
            Write-Log "Printer driver '$printerDriver' is still in use by other printers. Keeping driver." -Level "INFO"
        }
    } catch {
        Write-Log "Could not check printer driver usage: $($_.Exception.Message)" -Level "WARN"
    }
} else {
    Write-Log "Driver removal not requested. Keeping printer driver: $printerDriver" -Level "INFO"
}

# If this was the default printer, set a new default
if ($isDefault) {
    try {
        $availablePrinters = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.PrinterStatus -eq "Idle" }
        if ($availablePrinters) {
            $newDefault = $availablePrinters | Select-Object -First 1
            Set-Printer -Name $newDefault.Name -IsDefault $true -ErrorAction SilentlyContinue | Out-Null
            Write-Log "Set new default printer: $($newDefault.Name)" -Level "INFO"
        }
    } catch {
        Write-Log "Could not set new default printer: $($_.Exception.Message)" -Level "WARN"
    }
}

Write-Log "Printer '$PrinterName' has been successfully removed from the system." -Level "INFO"
Write-Output "Printer '$PrinterName' has been successfully removed from the system."
