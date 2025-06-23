# VMware Tools Installation Script
# This script downloads and installs the latest VMware Tools for Windows
# ================================================

$ProgressPreference = 'SilentlyContinue'

$VMWareToolsURL = "https://packages.vmware.com/tools/esx/latest/windows/x64/"
$VMWareToolsPath = $env:TEMP


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
        Write-Host "Downloading VMware Tools from: $fullURL"
        Write-Host "Saving to: $localPath"
        
        # Download the installer
        Invoke-WebRequest -Uri $fullURL -OutFile $localPath -UseBasicParsing
        
        if (Test-Path $localPath) {
            Write-Host "Download completed successfully"
            
            # Install VMware Tools silently
            Write-Host "Installing VMware Tools..."
            $arguments = "/S /v /qn REBOOT=R"
            Start-Process -FilePath $localPath -ArgumentList $arguments -Wait
            
            Write-Host "VMware Tools installation completed"
            
            # Clean up the installer file
            Remove-Item $localPath -Force
            Write-Host "Cleaned up installer file"
        } else {
            throw "Download failed - file not found at expected location"
        }
    }
    catch {
        Write-Error "Failed to download or install VMware Tools: $_"
        return $false
    }
    
    return $true
}

# Main execution
Write-Host "Starting VMware Tools installation process..."

# Get the latest installer filename
$installerFileName = Get-LatestVMwareToolsInstaller

if ($installerFileName) {
    Write-Host "Found latest VMware Tools installer: $installerFileName"
    
    # Download and install
    $success = Install-VMwareTools -InstallerFileName $installerFileName
    
    if ($success) {
        Write-Host "VMware Tools installation process completed successfully"
        exit 0
    } else {
        Write-Error "VMware Tools installation failed"
        exit 1
    }
} else {
    Write-Error "Could not determine latest VMware Tools installer"
    exit 1
}