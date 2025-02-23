# =================================================================
# Install Single App via Chocolatey Script for Action1
# =================================================================
# Description:
#   - Installs applications using Chocolatey package manager
#   - Handles prerequisites and logging
#   - Provides detailed logging of the installation process
#   - Installs Chocolatey if not present
#
# Parameters:
#   - App Name: The Chocolatey package name to install
# =================================================================

$ProgressPreference = 'SilentlyContinue'

$AppName = ${App Name}

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = "$env:SystemDrive\LST\Action1.log", # Default log file path
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
            Add-Content -Path $LogFilePath -Value "[$timestamp] [INFO] Log file exceeded 5 MB limit and was reset."
        }
    }
    
    # Write log entry to the log file
    Add-Content -Path $LogFilePath -Value $logMessage

    # Write output to Action1 host
    Write-Output "$Message"
}

function Install-Chocolatey {
    Write-Log "Installing Chocolatey Package Manager..." -Level "INFO"
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        
        # Verify installation
        if (Get-Command choco -ErrorAction SilentlyContinue) {
            Write-Log "Chocolatey Package Manager installed successfully" -Level "INFO"
            return $true
        } else {
            Write-Log "Chocolatey installation completed but verification failed" -Level "ERROR"
            return $false
        }
    } catch {
        Write-Log "Failed to install Chocolatey Package Manager: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Test-Prerequisites {
    Write-Log "Checking prerequisites..."
    
    # Validate App Name parameter
    if ([string]::IsNullOrWhiteSpace($AppName)) {
        Write-Log "No application name provided" -Level "ERROR"
        return $false
    }
    
    # Check if Chocolatey is installed
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Log "Chocolatey is not installed. Attempting to install..." -Level "WARN"
        if (-not (Install-Chocolatey)) {
            Write-Log "Failed to install Chocolatey" -Level "ERROR"
            return $false
        }
    }
    
    return $true
}

function Install-ChocolateyPackage {
    param (
        [string]$PackageName
    )
    
    Write-Log "Installing $PackageName via Chocolatey"
    
    try {
        # Refresh environment variables
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        
        # Execute Chocolatey installation
        $chocoPath = "$env:ProgramData\chocolatey\bin\choco.exe"
        $process = Start-Process -FilePath $chocoPath -ArgumentList "install", $PackageName, "-y", "-r" -NoNewWindow -PassThru -Wait
        
        switch ($process.ExitCode) {
            0 {
                Write-Log "$PackageName installed successfully"
                return $true
            }
            1641 {
                Write-Log "$PackageName installed successfully - reboot required"
                return $true
            }
            3010 {
                Write-Log "$PackageName installed successfully - reboot required"
                return $true
            }
            default {
                Write-Log "Installation failed for $PackageName (Exit code: $($process.ExitCode))" -Level "ERROR"
                return $false
            }
        }
    }
    catch {
        Write-Log "Installation failed: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# Main execution
try {
    Write-Log "Starting installation process for $AppName"
    
    if (!(Test-Prerequisites)) {
        Write-Log "Prerequisites check failed" -Level "ERROR"
        exit 1
    }
    
    if (Install-ChocolateyPackage -PackageName $AppName) {
        Write-Log "Installation completed successfully"
        exit 0
    }
    else {
        Write-Log "Installation failed" -Level "ERROR"
        exit 1
    }
}
catch {
    Write-Log "Error: $($_.Exception.Message)" -Level "ERROR"
    exit 1
} 