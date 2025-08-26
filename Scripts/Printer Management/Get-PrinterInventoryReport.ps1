# ================================================
# Printer Inventory Report Script for Action1
# ================================================
# Description:
#   - This script generates a comprehensive report of all installed printers on the system.
#   - Shows printer details including IP addresses, ports, drivers, status, and configuration.
#   - Outputs a formatted report suitable for Action1 console display.
#
# Requirements:
#   - Admin rights are required.
#

# ================================================

$ProgressPreference = 'SilentlyContinue'

# Action1 Variables
$IncludeOffline = ${Include Offline} # Optional: Set to "true" to include offline/error printers
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

function Get-PrinterDetails {
    param (
        [Microsoft.Management.Infrastructure.CimInstance]$Printer
    )
    
    $details = @{}
    
    # Basic printer info
    $details.Name = $Printer.Name
    $details.PortName = $Printer.PortName
    $details.DriverName = $Printer.DriverName
    $details.Shared = $Printer.Shared
    $details.Published = $Printer.Published
    $details.IsDefault = $Printer.PrinterStatus -eq "Idle" -and (Get-WmiObject -Class Win32_Printer -Filter "Name='$($Printer.Name)'").Default
    
    # Extract IP address from port name
    if ($Printer.PortName -match "TCPPort:(.+)") {
        $details.IPAddress = $matches[1]
    } elseif ($Printer.PortName -match "IP_(.+)") {
        $details.IPAddress = $matches[1]
    } else {
        $details.IPAddress = "N/A"
    }
    
    # Get port details
    try {
        $port = Get-PrinterPort -Name $Printer.PortName -ErrorAction SilentlyContinue
        if ($port) {
            $details.PortHostAddress = $port.PrinterHostAddress
            $details.PortNumber = $port.PortNumber
            $details.PortProtocol = $port.PortProtocol
        } else {
            $details.PortHostAddress = "N/A"
            $details.PortNumber = "N/A"
            $details.PortProtocol = "N/A"
        }
    } catch {
        $details.PortHostAddress = "N/A"
        $details.PortNumber = "N/A"
        $details.PortProtocol = "N/A"
    }
    
    # Get printer configuration
    try {
        $config = Get-PrintConfiguration -PrinterName $Printer.Name -ErrorAction SilentlyContinue
        if ($config) {
            $details.PaperSize = $config.PaperSize
            $details.Color = if ($config.Color -eq 1) { "Color" } else { "Black & White" }
            $details.Duplex = $config.Duplex
            $details.Orientation = $config.Orientation
        } else {
            $details.PaperSize = "N/A"
            $details.Color = "N/A"
            $details.Duplex = "N/A"
            $details.Orientation = "N/A"
        }
    } catch {
        $details.PaperSize = "N/A"
        $details.Color = "N/A"
        $details.Duplex = "N/A"
        $details.Orientation = "N/A"
    }
    
    # Get printer status
    $details.Status = $Printer.PrinterStatus
    $details.Jobs = $Printer.NumberOfJobs
    
    # Get location and comment
    try {
        $wmiPrinter = Get-WmiObject -Class Win32_Printer -Filter "Name='$($Printer.Name)'" -ErrorAction SilentlyContinue
        if ($wmiPrinter) {
            $details.Location = if ($wmiPrinter.Location) { $wmiPrinter.Location } else { "N/A" }
            $details.Comment = if ($wmiPrinter.Comment) { $wmiPrinter.Comment } else { "N/A" }
        } else {
            $details.Location = "N/A"
            $details.Comment = "N/A"
        }
    } catch {
        $details.Location = "N/A"
        $details.Comment = "N/A"
    }
    
    return $details
}

function Format-PrinterReport {
    param (
        [array]$Printers
    )
    
    if ($Printers.Count -eq 0) {
        Write-Output "No printers found on this system."
        return
    }
    
    # Header
    Write-Output "================================================"
    Write-Output "           PRINTER INVENTORY REPORT"
    Write-Output "================================================"
    Write-Output "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Output "Total Printers: $($Printers.Count)"
    Write-Output ""
    
    # Summary
    $defaultPrinters = $Printers | Where-Object { $_.IsDefault }
    $sharedPrinters = $Printers | Where-Object { $_.Shared }
    $networkPrinters = $Printers | Where-Object { $_.IPAddress -ne "N/A" }
    
    Write-Output "SUMMARY:"
    Write-Output "  • Default Printers: $($defaultPrinters.Count)"
    Write-Output "  • Shared Printers: $($sharedPrinters.Count)"
    Write-Output "  • Network Printers: $($networkPrinters.Count)"
    Write-Output ""
    
    # Detailed report for each printer
    foreach ($printer in $Printers) {
        Write-Output "================================================"
        Write-Output "PRINTER: $($printer.Name)"
        Write-Output "================================================"
        
        # Basic Info
        Write-Output "  Basic Information:"
        Write-Output "    • Status: $($printer.Status)"
        Write-Output "    • Driver: $($printer.DriverName)"
        Write-Output "    • Default: $(if ($printer.IsDefault) { 'Yes' } else { 'No' })"
        Write-Output "    • Shared: $(if ($printer.Shared) { 'Yes' } else { 'No' })"
        Write-Output "    • Published: $(if ($printer.Published) { 'Yes' } else { 'No' })"
        Write-Output "    • Active Jobs: $($printer.Jobs)"
        
        # Network Info
        Write-Output "  Network Configuration:"
        Write-Output "    • Port: $($printer.PortName)"
        Write-Output "    • IP Address: $($printer.IPAddress)"
        Write-Output "    • Host Address: $($printer.PortHostAddress)"
        Write-Output "    • Port Number: $($printer.PortNumber)"
        Write-Output "    • Protocol: $($printer.PortProtocol)"
        
        # Configuration
        Write-Output "  Print Configuration:"
        Write-Output "    • Paper Size: $($printer.PaperSize)"
        Write-Output "    • Color: $($printer.Color)"
        Write-Output "    • Duplex: $($printer.Duplex)"
        Write-Output "    • Orientation: $($printer.Orientation)"
        
        # Additional Info
        Write-Output "  Additional Information:"
        Write-Output "    • Location: $($printer.Location)"
        Write-Output "    • Comment: $($printer.Comment)"
        
        Write-Output ""
    }
    
    # Footer
    Write-Output "================================================"
    Write-Output "Report completed successfully."
    Write-Output "================================================"
}

# ================================
# Main Script Logic
# ================================

Write-Log "Starting printer inventory report generation..." -Level "INFO"

# Convert IncludeOffline from boolean to string
if ($IncludeOffline -eq 1) {
    $IncludeOffline = "true"
}
if ($IncludeOffline -eq 0) {
    $IncludeOffline = "false"
}

try {
    # Get all printers
    $allPrinters = Get-Printer -ErrorAction Stop

    # Filter printers based on IncludeOffline setting
    if ($IncludeOffline -eq "true") {
        $filteredPrinters = $allPrinters
        Write-Log "Including all printers (online and offline)" -Level "INFO"
    } else {
        $filteredPrinters = $allPrinters | Where-Object { $_.PrinterStatus -eq "Idle" -or $_.PrinterStatus -eq "Ready" }
        Write-Log "Filtering to show only online/ready printers" -Level "INFO"
    }
    
    if ($filteredPrinters.Count -eq 0) {
        Write-Log "No printers found matching criteria" -Level "WARN"
        Write-Output "No printers found on this system."
        return
    }
    
    Write-Log "Found $($filteredPrinters.Count) printers to report on" -Level "INFO"
    
    # Get detailed information for each printer
    $printerDetails = @()
    foreach ($printer in $filteredPrinters) {
        try {
            $details = Get-PrinterDetails -Printer $printer
            $printerDetails += $details
            Write-Log "Processed printer: $($printer.Name)" -Level "INFO"
        } catch {
            Write-Log "Failed to get details for printer '$($printer.Name)': $($_.Exception.Message)" -Level "WARN"
        }
    }
    
    # Generate and display the report
    Write-Log "Generating formatted report..." -Level "INFO"
    Format-PrinterReport -Printers $printerDetails
    
    Write-Log "Printer inventory report completed successfully" -Level "INFO"
    
} catch {
    Write-Log "Failed to generate printer report: $($_.Exception.Message)" -Level "ERROR"
    Write-Output "ERROR: Failed to generate printer report: $($_.Exception.Message)"
}
