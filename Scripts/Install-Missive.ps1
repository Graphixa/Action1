$ProgressPreference = 'SilentlyContinue'

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

$softwareName = 'Missive'
$checkLocation = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
$tempPath = "$env:SystemDrive\Temp\"
$jsonFile = $tempPath + "latest.json"

# Check if Missive is already installed
if (Get-ChildItem $checkLocation -Recurse -ErrorAction Stop | Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -Match "^$softwareName.*" }) {
    Write-Log "$softwareName is already installed. No action required." -Level "INFO"
    return
}

Write-Log "$softwareName is NOT installed. Installing Now..." -Level "INFO"

# URL for the latest Missive version JSON file (Hosted by Missiveapp.com)
$jsonUrl = 'https://missiveapp.com/download/latest.json'

# Create Temp Folder in the Root of the System Drive
try {
    New-Item -ItemType Directory -Force -Path $tempPath
    Write-Log "Created temporary directory: $tempPath" -Level "INFO"
} catch {
    Write-Log "Failed to create temp directory: $tempPath" -Level "ERROR"
    return
}

# Download JSON file to Temp Folder Location
try {
    Invoke-RestMethod -Method Get -Uri $jsonUrl -OutFile $jsonFile
    Write-Log "Downloaded JSON file for Missive version info." -Level "INFO"
} catch {
    Write-Log "Failed to download JSON file for Missive." -Level "ERROR"
    Write-Log $_.Exception.Message -Level "ERROR"
    return
}

# Search JSON file and return the URL of the latest Missive release for Windows
$json = (Get-Content $jsonFile -Raw) | ConvertFrom-Json
$version = $json.version
$windowsDL = $json.downloads.windows.direct
$missiveFile = "$tempPath\$softwareName-$version.exe"

# Download latest Missive version for Windows
try {
    Write-Log "Downloading Missive Installer from $windowsDL" -Level "INFO"
    Invoke-WebRequest -Method Get -Uri $windowsDL -OutFile $missiveFile
    Write-Log "Download completed: $missiveFile" -Level "INFO"
} catch {
    Write-Log "Failed to download Missive installer." -Level "ERROR"
    Write-Log $_.Exception.Message -Level "ERROR"
    return
}

# Install Missive silently for all users
try {
    Write-Log "Installing Missive for all users." -Level "INFO"
    Start-Process -Wait -FilePath $missiveFile -ArgumentList "/S /D=$env:SystemDrive\$softwareName" -PassThru
    Write-Log "Missive installation completed." -Level "INFO"
} catch {
    Write-Log "Missive installation failed." -Level "ERROR"
    Write-Log $_.Exception.Message -Level "ERROR"
    return
}

# Create Shortcuts for all users
$missiveExecutable = "$env:SystemDrive\$softwareName\Missive.exe"
$startMenuPath = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Missive.lnk"
$desktopPath = "$env:Public\Desktop\Missive.lnk"
$WScriptObj = New-Object -ComObject ("WScript.Shell")

try {
    Write-Log "Creating desktop and start menu shortcuts for all users." -Level "INFO"
    
    # Create Shortcut in All-Users Start Menu
    $shortcutStart = $WscriptObj.CreateShortcut($startMenuPath)
    $shortcutStart.TargetPath = $missiveExecutable
    $shortcutStart.Save()

    # Create Shortcut on All-Users Desktop
    $shortcutDesktop = $WscriptObj.CreateShortcut($desktopPath)
    $shortcutDesktop.TargetPath = $missiveExecutable
    $shortcutDesktop.Save()

    Write-Log "Shortcuts created successfully." -Level "INFO"
} catch {
    Write-Log "Failed to create desktop and start menu shortcuts." -Level "ERROR"
    Write-Log $_.Exception.Message -Level "ERROR"
}

# Cleanup temporary files at the end
try {
    Write-Log "Cleaning up temporary files and folders." -Level "INFO"
    Remove-Item -Path $tempPath -Recurse -Force
    Write-Log "Temporary files and folders cleaned up successfully." -Level "INFO"
} catch {
    Write-Log "Failed to remove temporary files and folders." -Level "ERROR"
    Write-Log $_.Exception.Message -Level "ERROR"
}