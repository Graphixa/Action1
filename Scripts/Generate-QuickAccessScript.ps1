$ProgressPreference = 'SilentlyContinue'

# Log Function
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = "$env:SystemDrive\LST-Action1.log", # Default log file path
        [string]$Level = "INFO"  # Log level: INFO, WARN, ERROR
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write log entry to the log file
    Add-Content -Path $LogFilePath -Value $logMessage
}

# Define the startup folder path
$scriptGenerateLocation = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\StartUp\"


# Generated Script Content (the script that will be created)
$scriptContent = @"
\$ProgressPreference = 'SilentlyContinue'


\$qa = New-Object -ComObject shell.application
\$quickAccessFolder = \$qa.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}')
\$items = \$quickAccessFolder.Items()

# Array of folders to pin
\$foldersToPin = @(
    "G:\My Drive",
    "G:\Shared drives\Management",
    "G:\Shared drives\Admin",
    "G:\Shared drives\Clinical",
    "G:\Shared drives\Document Library",
    "G:\Shared drives\Creative",
    "G:\Shared drives\HR Drive",
    "G:\Shared drives\IT"
)


# Unpin all existing folders
foreach (\$item in \$items) {
    \$item.InvokeVerb('unpinfromhome')
}

# Pin new folders to Quick Access
foreach (\$folder in \$foldersToPin) {
    \$QuickAccess = New-Object -ComObject shell.application
    \$QuickAccess.Namespace(\$folder).Self.InvokeVerb("pintohome")
}

# ================================
# End of Script
# ================================
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
