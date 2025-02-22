# ================================================
# Generate and Schedule Quick Access Pin Script for Action1
# ================================================
# Description:
#   - This script generates another script that pins specific Google Drive folders to Quick Access in File Explorer.
#   - You can optionally pin "My Drive" and provide a list of shared drive folders to pin.
#   - The generated script is saved in the Windows Startup folder so that it runs automatically when the system starts.
#   - Both the "My Drive" and the shared drive folders can be pinned independently or together based on the input.
#
# Requirements:
#   - Admin rights are required.
#   - Script must be run with administrative privileges to write to the Startup folder.
#
# ================================================

$ProgressPreference = 'SilentlyContinue'

$PinMyDrive = ${Pin My Drive}
$SharedFolders = ${Shared Drives To Pin}


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

# Define the startup folder path
$scriptGenerateLocation = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\StartUp\"

# Initialize foldersToPin array
$foldersToPin = @()

# Add My Drive if $PinMyDrive is set to 1
if ($PinMyDrive -eq 1) {
    $foldersToPin += "My Drive"
}

# Split the Folders parameter into an array if provided
if (-not [string]::IsNullOrWhiteSpace($SharedFolders)) {
    $sharedDriveFolders = $SharedFolders -split ',' | ForEach-Object { $_.Trim() }
    $foldersToPin += $sharedDriveFolders
}

# If no folders were provided, log and exit
if (-not $foldersToPin) {
    Write-Log "No folders specified to pin. Exiting script." -Level "WARN"
    return
}

# Generated Script Content (the script that will be created)
$scriptContent = @"
`$ProgressPreference = 'SilentlyContinue'

`$qa = New-Object -ComObject shell.application
`$quickAccessFolder = `$qa.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}')
`$items = `$quickAccessFolder.Items()

# Array of folders to pin
`$foldersToPin = @(
"@

# Add each folder to the generated script content
foreach ($folder in $foldersToPin) {
    if ($folder -eq "My Drive") {
        $scriptContent += "`"G:\My Drive`",`n"
    } else {
        $scriptContent += "`"G:\Shared drives\$folder`",`n"
    }
}

# Complete the script content
$scriptContent += @"
)

# Unpin all existing folders
foreach (`$item in `$items) {
    `$item.InvokeVerb('unpinfromhome')
}

# Pin new folders to Quick Access
foreach (`$folder in `$foldersToPin) {
    `$QuickAccess = New-Object -ComObject shell.application
    `$QuickAccess.Namespace(`$folder).Self.InvokeVerb('pintohome')
}
"@


# Define the path of the new script in the startup folder
$scriptPath = Join-Path $scriptGenerateLocation "PinFoldersToQuickAccess.ps1"

# Write the script content to the file in the startup folder
try {
    Set-Content -Path $scriptPath -Value $scriptContent -Force
    Write-Log "Script created successfully at: $scriptPath" -Level "INFO"
} catch {
    Write-Log "Failed to create script: $($_.Exception.Message)" -Level "ERROR"
}
