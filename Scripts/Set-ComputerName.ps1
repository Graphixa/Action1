# ================================================
# Computer Name Assignment Script for Action1
# ================================================
# Description:
#   - Sets computer name based on device serial number and system type
#   - Workstation = WS-<Serial Number>
#   - Notebook = NB-<Serial Number>  
#   - Virtual Machine = VM-<Serial Number>
#   - Names are limited to 15 characters for NetBIOS compatibility
#
# Requirements:
#   - Admin rights required
#   - Windows systems only
# ================================================

$CompanyPrefix = ${Computer Name Prefix}
$LogFilePath = "$env:SystemDrive\LST\Action1.log"

# Simple logging function
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

function Set-ComputerName {
    Write-Log "Setting Computer Name" -Level "INFO"

    # Get current computer name from registry (authoritative source)
    $currentName = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName" -Name "ComputerName").ComputerName
    

    # Get serial number from BIOS
    $serialNumber = (Get-CimInstance -class win32_bios).SerialNumber

    # Get system information
    $system = Get-CimInstance -ClassName Win32_ComputerSystem
    $systemType = $system.PCSystemType
    $systemModel = $system.Model

    # Determine system type and build computer name
    $isVM = $systemModel -match "Virtual|VMware|KVM|Hyper-V|VirtualBox|QEMU"
    
    if ($isVM) {
        # For VMs, use the last part of the serial (removes hypervisor prefixes like "VMware-")
        if ($serialNumber -match '.*-(.+)$') {
            $serialNumber = $matches[1]
            Write-Log "VM detected - using unique part of serial: $serialNumber" -Level "INFO"
        }
        
        $typePrefix = "VM"
        $systemTypeDesc = "Virtual Machine"
    }
    elseif ($systemType -eq 1 -or $systemType -eq 3) {
        $typePrefix = "WS"
        $systemTypeDesc = "Workstation (Type $systemType)"
    }
    elseif ($systemType -eq 2) {
        $typePrefix = "NB"
        $systemTypeDesc = "Notebook"
    }
    else {
        $typePrefix = ""
        $systemTypeDesc = "Unknown ($systemType)"
    }

    # Clean up serial number
    $serialNumber = $serialNumber -replace '\s', ''
    
    # Build base name components
    $companyPart = if ($CompanyPrefix) { "$CompanyPrefix-" } else { "" }
    $typePart = if ($typePrefix) { "$typePrefix-" } else { "" }
    $baseName = "$companyPart$typePart"
    
    # Calculate available space for serial number (15 character NetBIOS limit)
    $availableSpace = 15 - $baseName.Length
    if ($serialNumber.Length -gt $availableSpace) {
        $serialNumber = $serialNumber.Substring(0, $availableSpace)
    }
    
    # Build final computer name
    $computerName = "${baseName}${serialNumber}"
    
    # Check if rename is needed
    if ($currentName -eq $computerName) {
        Write-Log "Computer name is already $computerName. No action needed." -Level "WARN"
        exit 0
    }

    # Rename the computer
    try {
        Write-Log "Renaming computer from '$currentName' to '$computerName'" -Level "INFO"
        Rename-Computer -NewName $computerName -Force | Out-Null
        Write-Log "Successfully renamed computer to $computerName" -Level "INFO"
        Write-Log "Restart required for changes to take full effect." -Level "INFO"
    }
    catch {
        Write-Log "Error renaming computer: $($_.Exception.Message)" -Level "ERROR"
        exit 1
    }
}

# Main execution
try {
    Set-ComputerName
    exit 0
} catch {
    Write-Log "Error: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}