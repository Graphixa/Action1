# ================================================
# Missive Installation Script for Action1
# ================================================
# Description:
#   - This script checks if Missive is installed. If not, it downloads the latest version and installs it.
#   - It also creates shortcuts for all users and cleans up temporary files afterward.
#
# Requirements:
#   - Admin rights are required.
#   - Internet access is required to download the installer.
# ================================================

$ProgressPreference = 'SilentlyContinue'

$softwareName = 'Missive'
$checkLocation = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
$installerUrl = 'https://mail.missiveapp.com/download/win'
$tempPath = "$env:SystemDrive\Temp"
$missiveFile = "$tempPath\MissiveSetup.exe"

# Logging Function
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
            Add-Content -Path $LogFilePath -Value "[$timestamp] [INFO] The log file exceeded the 5 MB limit and was deleted and recreated."
        }
    }
    
    # Write log entry to the log file
    Add-Content -Path $LogFilePath -Value $logMessage

    # Write output to Action1 host
    Write-Output "$Message"
}

# Function to remove Missive installer and temp files
function Remove-TempFiles {
    try {
        Write-Host "Cleaning up temporary files and folders..."
        if (Test-Path $missiveFile) { Remove-Item $missiveFile -Force }
        if (Test-Path $tempPath) { Remove-Item $tempPath -Force }
        Start-Sleep -Seconds 2
    } catch {
        Write-Host -BackgroundColor Red -ForegroundColor White " Error: Failed to remove temporary files "
        Write-Host $_.Exception.Message
        Start-Sleep -Seconds 3
    }
}

# Check if Missive is already installed
if (Get-ChildItem $checkLocation -Recurse -ErrorAction Stop | 
    Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue | 
    Where-Object {$_.DisplayName -Match "^$softwareName.*"}) {
    Write-Host -ForegroundColor Yellow "$softwareName is already installed. No action required."
    Return 0
}

Write-Host -ForegroundColor Yellow "$softwareName is NOT installed. Installing Now..."

# Create Temp Folder
[void](New-Item -ItemType Directory -Force -Path $tempPath)

# Download Missive installer
try {
    Write-Host "Downloading Missive installer..."
    Invoke-WebRequest -Uri $installerUrl -OutFile $missiveFile -ErrorAction Stop
} catch {
    Write-Host -BackgroundColor Red -ForegroundColor White " Error: Download failed "
    Write-Host $_.Exception.Message
    Remove-TempFiles
    Return 1
}

# Install Missive silently for all users
try {
    Write-Host "Installing Missive for all users..."
    $process = Start-Process -Wait -FilePath $missiveFile -ArgumentList "/S /D=$InstallPath" -PassThru
    if ($process.ExitCode -ne 0) {
        throw "Installation failed with exit code: $($process.ExitCode)"
    }
} catch {
    Write-Host -BackgroundColor Red -ForegroundColor White " Error: Installation failed "
    Write-Host $_.Exception.Message
    Remove-TempFiles
    Return 1
}

# Create shortcuts for all users
try {
    Write-Host "Creating shortcuts for all users..."
    $missiveExecutable = "$env:SystemDrive\$softwareName\Missive.exe"
    $WScriptObj = New-Object -ComObject ("WScript.Shell")

    # Create Start Menu shortcut
    $startMenuPath = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Missive.lnk"
    $shortcutStart = $WscriptObj.CreateShortcut($startMenuPath)
    $shortcutStart.TargetPath = $missiveExecutable
    $shortcutStart.Save()

    # Create Desktop shortcut
    $desktopPath = "$env:Public\Desktop\Missive.lnk"
    $shortcutDesktop = $WscriptObj.CreateShortcut($desktopPath)
    $shortcutDesktop.TargetPath = $missiveExecutable
    $shortcutDesktop.Save()
} catch {
    Write-Host -BackgroundColor Red -ForegroundColor White " Error: Failed to create shortcuts "
    Write-Host $_.Exception.Message
}

# Cleanup
Remove-TempFiles
Write-Host -ForegroundColor Green "Missive installation completed successfully!"