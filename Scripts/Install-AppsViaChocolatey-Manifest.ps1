# =================================================================
# Install Apps via Chocolatey Manifest Script for Action1
# =================================================================
# Description:
#   - Installs applications using the Chocolatey package manager
#   - Handles prerequisites and logging
#   - Installs Chocolatey Package Manager if not present
#   - Installs packages from a specified manifest file
#
# Parameters:
#   - App Manifest Link: URL to the Chocolatey package manifest file
# =================================================================

$ProgressPreference = 'SilentlyContinue'

$ChocolateyAppManifest = ${App Manifest} # URL to the Chocolatey package manifest file or local path
$tempPath = "$env:TEMP"
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

function Install-Chocolatey {
    Write-Log "Installing Chocolatey Package Manager..." -Level "INFO"
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        
        # Refresh environment variables
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        
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
    
    # Validate manifest URL
    if ([string]::IsNullOrWhiteSpace($ChocolateyAppManifest)) {
        Write-Log "No manifest URL provided" -Level "ERROR"
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

function Get-ChocolateyManifest {
    param (
        [string]$ManifestUrl,
        [string]$OutputPath
    )
    
    try {
        Write-Log "Downloading Chocolatey manifest from $ManifestUrl"
        Invoke-WebRequest -Uri $ManifestUrl -OutFile $OutputPath
        
        if (Test-Path $OutputPath) {
            Write-Log "Manifest downloaded successfully"
            return $true
        } else {
            Write-Log "Manifest download failed - file not found" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Failed to download manifest: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Install-ChocolateyPackages {
    param (
        [string]$ManifestPath
    )
    
    try {
        Write-Log "Installing packages from manifest"
        
        # Refresh environment variables
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        
        # Execute Chocolatey installation
        $chocoPath = "$env:ProgramData\chocolatey\bin\choco.exe"
        $process = Start-Process -FilePath $chocoPath -ArgumentList "install", $ManifestPath, "--yes" -NoNewWindow -PassThru -Wait
        
        switch ($process.ExitCode) {
            0 {
                Write-Log "All packages installed successfully"
                return $true
            }
            1641 {
                Write-Log "Packages installed successfully - reboot required"
                return $true
            }
            3010 {
                Write-Log "Packages installed successfully - reboot required"
                return $true
            }
            default {
                Write-Log "Installation failed (Exit code: $($process.ExitCode))" -Level "ERROR"
                return $false
            }
        }
    }
    catch {
        Write-Log "Installation failed: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Remove-TempFiles {
    param (
        [string]$ManifestPath
    )
    
    try {
        if (Test-Path $ManifestPath) {
            Remove-Item -Path $ManifestPath -Force
            Write-Log "Temporary files cleaned up successfully"
        }
    }
    catch {
        Write-Log "Failed to clean up temporary files: $($_.Exception.Message)" -Level "WARN"
    }
}

# Main execution
try {
    Write-Log "Starting Chocolatey manifest installation process"
    
    if (!(Test-Prerequisites)) {
        Write-Log "Prerequisites check failed" -Level "ERROR"
        exit 1
    }
    
    $manifestPath = Join-Path $tempPath "chocolatey-manifest.config"
    
    if (!(Get-ChocolateyManifest -ManifestUrl $ChocolateyAppManifest -OutputPath $manifestPath)) {
        Write-Log "Failed to get manifest" -Level "ERROR"
        exit 1
    }
    
    if (Install-ChocolateyPackages -ManifestPath $manifestPath) {
        Write-Log "Installation completed successfully"
        Remove-TempFiles -ManifestPath $manifestPath
        exit 0
    }
    else {
        Write-Log "Installation failed" -Level "ERROR"
        Remove-TempFiles -ManifestPath $manifestPath
        exit 1
    }
}
catch {
    Write-Log "Error: $($_.Exception.Message)" -Level "ERROR"
    Remove-TempFiles -ManifestPath $manifestPath
    exit 1
}
