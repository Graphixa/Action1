# =================================================================
# Install App via Winget Script for Action1
# =================================================================
# Description:
#   - Installs applications using Winget in SYSTEM context
#   - Handles prerequisites and system context setup
#   - Provides detailed logging of the installation process
#
# Parameters:
#   - App ID: The Winget package identifier
#   - Version: Optional specific version to install
# =================================================================

$ProgressPreference = 'SilentlyContinue'

$AppID = ${App ID}
$Version = ${Version}

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = "$env:SystemDrive\Logs\Action1.log",
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
            Add-Content -Path $LogFilePath -Value "[$timestamp] [INFO] Log file exceeded 5 MB limit and was reset."
        }
    }
    
    # Write log entry to the log file
    Add-Content -Path $LogFilePath -Value $logMessage

    # Write output to Action1 host
    Write-Output "$Message"
}

function Test-Prerequisites {
    Write-Log "Checking prerequisites..."
    
    # Check OS version
    $osVersion = [Environment]::OSVersion.Version
    if ($osVersion.Major -lt 10 -or ($osVersion.Major -eq 10 -and $osVersion.Build -lt 17763)) {
        Write-Log "Windows 10 1809 (build 17763) or later is required." -Level "ERROR"
        return $false
    }
    
    # Check for Winget
    $winget = Get-AppxPackage -AllUsers | Where-Object { $_.Name -eq 'Microsoft.DesktopAppInstaller' }
    if (!$winget) {
        Write-Log "Winget is not installed." -Level "ERROR"
        return $false
    }
    
    return $true
}

function Initialize-SystemContext {
    # Set system environment variables
    $env:SystemRoot = "$env:SystemDrive\Windows"
    $env:TEMP = "$env:SystemRoot\TEMP"
    $env:TMP = "$env:SystemRoot\TEMP"
    
    # Create system profile directories if they don't exist
    $paths = @(
        "$env:SystemRoot\System32\config\systemprofile",
        "$env:SystemRoot\System32\config\systemprofile\AppData\Local",
        "$env:SystemRoot\System32\config\systemprofile\AppData\Roaming"
    )
    
    foreach ($path in $paths) {
        if (!(Test-Path $path)) {
            New-Item -Path $path -ItemType Directory -Force | Out-Null
        }
    }
}

function Install-WingetPackage {
    param (
        [string]$Id,
        [string]$Version
    )
    
    Write-Log "Installing $Id$(if($Version){" version $Version"})"
        
    try {
        $winget = Get-AppxPackage -AllUsers | Where-Object { $_.Name -eq 'Microsoft.DesktopAppInstaller' }
        $wingetExe = Join-Path $winget.InstallLocation 'winget.exe'
        
        # Build installation command
        $arguments = "install --silent --exact --scope machine --source winget --accept-package-agreements --accept-source-agreements --id ""$Id"""
        if ($Version) {
            $arguments += " --version ""$Version"""
        }
        
        # Start winget process
        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = $wingetExe
        $pinfo.Arguments = $arguments
        $pinfo.RedirectStandardOutput = $true
        $pinfo.RedirectStandardError = $true
        $pinfo.UseShellExecute = $false
        $pinfo.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $pinfo
        
        $process.Start() | Out-Null
        
        # Capture output
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        
        $process.WaitForExit()
        
        # Check for success patterns in output
        if ($stdout -match "Successfully installed") {
            Write-Log "Package installed successfully"
            return 0
        }
        elseif ($stdout -match "already installed") {
            Write-Log "Package is already installed"
            return -1978335189
        }
        
        # Log errors if any
        if ($stderr) {
            Write-Log $stderr -Level "ERROR"
        }
        
        return $process.ExitCode
    }
    catch {
        Write-Log "Installation failed: $($_.Exception.Message)" -Level "ERROR"
        return -1
    }
}

# Main execution
try {
    Write-Log "Starting installation of $AppID"
    
    if (!(Test-Prerequisites)) {
        Write-Log "Prerequisites check failed" -Level "ERROR"
        exit 1
    }
    
    Initialize-SystemContext
    
    $result = Install-WingetPackage -Id $AppID -Version $Version
    $exitCode = if ($result -is [array]) { $result[-1] } else { $result }
    
    switch ($exitCode) {
        0 { 
            Write-Log "Installation of $AppID completed successfully"
            exit 0
        }
        -1978335189 {
            Write-Log "$AppID is already installed"
            exit 0
        }
        -1978335212 {
            Write-Log "$AppID is not found as a available package" -Level "ERROR"
            exit 1
        }
        default {
            Write-Log "$AppID installation failed (Exit code: $exitCode)" -Level "ERROR"
            exit 1
        }
    }
}
catch {
    Write-Log "Error: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}
