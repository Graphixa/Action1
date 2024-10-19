# ================================================
# Computer Name Assignment Script for Action1
# ================================================
# Description:
#   - This script sets the computer name based on the device's serial number.
#   - The script detects the system type (workstation, notebook, virtual machine) and assigns a name accordingly.
#   - Serial numbers are sanitized by removing spaces and truncating them to 15 characters.
#   - The prefix is optional; if not provided, it defaults to no prefix.
#
# Requirements:
#   - Admin rights are required to rename the computer.
#   - The serial number should be accessible through the Win32_BIOS class.
#
# Example usage:
#   # Run the script with a prefix:
#   & "C:\Scripts\Set-ComputerName.ps1" -prefix "MSFT"
#
#   # Run the script without a prefix:
#   & "C:\Scripts\Set-ComputerName.ps1"
# ================================================

param (
    [string]$prefix = ""  # Company prefix is optional, default is no prefix
)

# Logging Function
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = "$env:SystemDrive\LST-Action1.log",
        [string]$Level = "INFO"  # INFO, WARN, ERROR
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write log entry to the log file
    Add-Content -Path $LogFilePath -Value $logMessage
}

# Set Computer Name from Device Serial Number
$SerialNumber = (Get-CimInstance -class win32_bios).SerialNumber

# Sanitize the serial number - remove spaces and limit length to 15 characters
$SerialNumber = $SerialNumber -replace ' ', ''
$SerialNumber = $SerialNumber.Substring(0, [Math]::Min(15, $SerialNumber.Length))

# Retrieve system type and model information
$systemTest = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty PCSystemType)
$systemModel = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Model)

# Check if the system is a virtual machine by looking at the model
$isVM = $systemModel -match "Virtual|VMware|KVM|Hyper-V|VirtualBox"

# Ensure there is a dash between the company prefix and computer type if a prefix is provided
if ($prefix) {
    $prefix += "-"
}

# Test system type - WS = Workstation, NB = Notebook, VM = Virtual Machine
if ($isVM) {
    $computerName = "${prefix}VM-$SerialNumber"
    Write-Log "Detected system as Virtual Machine. New computer name: $computerName"
}
elseif ($systemTest -eq 1) {
    $computerName = "${prefix}WS-$SerialNumber"
    Write-Log "Detected system as Workstation. New computer name: $computerName"
}
elseif ($systemTest -eq 2) {
    $computerName = "${prefix}NB-$SerialNumber"
    Write-Log "Detected system as Notebook. New computer name: $computerName"
}
elseif ($systemTest -eq 3) {
    $computerName = "${prefix}WS-$SerialNumber"
    Write-Log "Detected system as Workstation. New computer name: $computerName"
}
else {
    $computerName = "${prefix}$SerialNumber"
    Write-Log "System type unrecognized. Defaulting to new computer name: $computerName"
}

# Attempt to rename the computer
try {
    Rename-Computer -NewName $computerName -Force
    Write-Log "Successfully renamed computer to $computerName"
}
catch {
    Write-Log "Error renaming computer: $_" -Level "ERROR"
    Return
}
