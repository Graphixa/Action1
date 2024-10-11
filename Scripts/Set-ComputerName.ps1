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
$SerialNumber = (Get-WmiObject -class win32_bios).SerialNumber

# Sanitize the serial number - remove spaces and limit length to 15 characters
$SerialNumber = $SerialNumber -replace ' ', ''
$SerialNumber = $SerialNumber.Substring(0, [Math]::Min(15, $SerialNumber.Length))

$systemTest = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty PCSystemType)

# Test system type - WS = Workstation, NB = Notebook
if ($systemTest -eq 1) {
    $computerName = "${Company Prefix}-WS-$SerialNumber"
    Write-Log "Detected system as Workstation. New computer name: $computerName"
}
elseif ($systemTest -eq 2) {
    $computerName = "${Company Prefix}-NB-$SerialNumber"
    Write-Log "Detected system as Notebook. New computer name: $computerName"
}
elseif ($systemTest -eq 3) {
    $computerName = "${Company Prefix}-WS-$SerialNumber"
    Write-Log "Detected system as Workstation. New computer name: $computerName"
}
else {
    $computerName = "${Company Prefix}-$SerialNumber"
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