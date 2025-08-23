# VMware Tools Installation Script
# This script downloads and installs the latest VMware Tools for Windows
# ================================================

$ProgressPreference = 'SilentlyContinue'

$VMWareToolsURL = "https://packages.vmware.com/tools/esx/latest/windows/x64/"
$VMWareToolsPath = $env:TEMP
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

# Function to check if VMware Tools is already installed
function Test-VMwareToolsInstalled {
    try {
        # Check for VMware Tools service
        $vmwareService = Get-Service -Name "VMTools" -ErrorAction SilentlyContinue
        if ($vmwareService -and $vmwareService.Status -eq "Running") {
            Write-Log "VMware Tools service is running" -Level "INFO"
            return $true
        }
        
        # Check for VMware Tools in Programs and Features
        $vmwareProgram = Get-WmiObject -Class Win32_Product -Filter "Name LIKE '%VMware Tools%'" -ErrorAction SilentlyContinue
        if ($vmwareProgram) {
            Write-Log "VMware Tools is installed (found in Programs and Features)" -Level "INFO"
            return $true
        }
        
        # Check for VMware Tools registry entries
        $vmwareRegistry = Get-ItemProperty -Path "HKLM:\SOFTWARE\VMware, Inc.\VMware Tools" -ErrorAction SilentlyContinue
        if ($vmwareRegistry) {
            Write-Log "VMware Tools registry entries found" -Level "INFO"
            return $true
        }
        
        # Check for VMware Tools installation directory
        $vmwareDir = Get-ItemProperty -Path "HKLM:\SOFTWARE\VMware, Inc.\VMware Tools" -Name "InstallPath" -ErrorAction SilentlyContinue
        if ($vmwareDir -and (Test-Path $vmwareDir.InstallPath)) {
            Write-Log "VMware Tools installation directory found" -Level "INFO"
            return $true
        }
        
        Write-Log "VMware Tools is not installed" -Level "INFO"
        return $false
    }
    catch {
        Write-Log "Error checking VMware Tools installation status: $_" -Level "ERROR"
        return $false
    }
}

# Function to get the latest VMware Tools installer filename
function Get-LatestVMwareToolsInstaller {
    try {
        # Get the directory listing page
        $response = Invoke-WebRequest -Uri $VMWareToolsURL -UseBasicParsing
        
        # Parse the HTML to find the latest installer
        # Look for files ending with .exe that contain "VMware-tools"
        $exeFiles = $response.Links | Where-Object { 
            $_.href -like "*.exe" -and $_.href -like "*VMware-tools*" 
        } | Sort-Object href -Descending
        
        if ($exeFiles.Count -eq 0) {
            throw "No VMware Tools installer found on the page"
        }
        
        # Return the first (latest) installer filename
        return $exeFiles[0].href
    }
    catch {
        Write-Error "Failed to get latest VMware Tools installer: $_"
        return $null
    }
}

# Function to download and install VMware Tools
function Install-VMwareTools {
    param(
        [string]$InstallerFileName
    )
    
    $fullURL = $VMWareToolsURL + $InstallerFileName
    $localPath = Join-Path $VMWareToolsPath $InstallerFileName
    
    try {
        Write-Log "Downloading VMware Tools from: $fullURL" -Level "INFO"
        Write-Log "Saving to: $localPath" -Level "INFO"
        
        # Download the installer
        Invoke-WebRequest -Uri $fullURL -OutFile $localPath -UseBasicParsing
        
        if (Test-Path $localPath) {
            Write-Log "Download completed successfully" -Level "INFO"
            
            # Install VMware Tools silently
            Write-Log "Installing VMware Tools..." -Level "INFO"
            $arguments = "/S /v /qn REBOOT=R"
            Start-Process -FilePath $localPath -ArgumentList $arguments -Wait
            
            Write-Log "VMware Tools installation completed" -Level "INFO"
            
            # Clean up the installer file
            Remove-Item $localPath -Force
            Write-Log "Cleaned up installer file" -Level "INFO"
        } else {
            throw "Download failed - file not found at expected location"
        }
    }
    catch {
        Write-Log "Failed to download or install VMware Tools: $_" -Level "ERROR"
        return $false
    }
    
    return $true
}

# Main execution
Write-Log "Starting VMware Tools installation process..." -Level "INFO"

# Check if VMware Tools is already installed
if (Test-VMwareToolsInstalled) {
    Write-Log "VMware Tools is already installed. Skipping installation." -Level "INFO"
    Write-Host "VMware Tools is already installed. No action needed."
    exit 0
}

# Get the latest installer filename
$installerFileName = Get-LatestVMwareToolsInstaller

if ($installerFileName) {
    Write-Log "Found latest VMware Tools installer: $installerFileName" -Level "INFO"
    
    # Download and install
    $success = Install-VMwareTools -InstallerFileName $installerFileName
    
    if ($success) {
        Write-Log "VMware Tools installation process completed successfully" -Level "INFO"
        Write-Host "VMware Tools installation process completed successfully"
        exit 0
    } else {
        Write-Log "VMware Tools installation failed" -Level "ERROR"
        Write-Error "VMware Tools installation failed"
        exit 1
    }
} else {
    Write-Log "Could not determine latest VMware Tools installer" -Level "ERROR"
    Write-Error "Could not determine latest VMware Tools installer"
    exit 1
}