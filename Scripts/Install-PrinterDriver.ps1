# ================================================
# Printer Installation Script for Action1
# ================================================
# Description:
#   - This script installs a printer using a TCP/IP port and installs the driver from a specified download.
#   - It sets up the printer with default configurations, including paper size and color settings.
#
# Requirements:
#   - Admin rights are required.
#   - Internet access is required to download the driver.
#

# ================================================

$ProgressPreference = 'SilentlyContinue'

# Define Variables
$PrinterIP = ${IP Address}
$PrinterName = ${Printer Name}
$PortName = "TCPPort:${PrinterIP}"
$TempDownloadFolder = "$env:TEMP"
$DownloadURL = "https://github.com/Graphixa/PCL6-Driver-for-Universal-Print/archive/refs/heads/main.zip"
$DownloadFileName = "$TempDownloadFolder\PCL6-Driver.zip"
$DownloadPath = "$TempDownloadFolder\PCL6-Driver"
$TempExtractPath = "$TempDownloadFolder\DriverFiles"

# Logging Function
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = "$env:SystemDrive\Logs\Action1.log", # Default log file path
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
    
    # Write log entry to the log file
    Add-Content -Path $LogFilePath -Value $logMessage

    # Write output to Action1 host
    Write-Output "$Message"
}

# Create $DownloadPath directory if it doesn't exist
if (-not (Test-Path -Path $DownloadPath)) {
    New-Item -Path $DownloadPath -ItemType Directory
    Write-Log "Created directory: $DownloadPath"
} else {
    Write-Log "Directory already exists: $DownloadPath"
}

# Create temporary extraction directory if it doesn't exist
if (-not (Test-Path -Path $TempExtractPath)) {
    New-Item -Path $TempExtractPath -ItemType Directory
    Write-Log "Created temporary extraction directory: $TempExtractPath"
} else {
    Write-Log "Temporary extraction directory already exists: $TempExtractPath"
}

# Download the ZIP file
try {
    Write-Log "Downloading file from: $DownloadURL"
    Invoke-WebRequest -Uri $DownloadURL -OutFile $DownloadFileName
    Write-Log "Download completed: $DownloadFileName"
} catch {
    Write-Log "Failed to download the ZIP file from $DownloadURL. Error: $_" -Level "ERROR"
    return
}

# Extract the ZIP file to the temporary extraction folder
try {
    Expand-Archive -Path $DownloadFileName -DestinationPath $TempExtractPath -Force
    Write-Log "Extraction completed to $TempExtractPath"
} catch {
    Write-Log "Failed to extract the ZIP file. Error: $_" -Level "ERROR"
    return
}

# Move the contents from the temporary extraction folder to $DownloadPath
try {
    $ExtractedFolder = Get-ChildItem -Path $TempExtractPath | Select-Object -First 1
    if ($ExtractedFolder) {
        Get-ChildItem -Path $ExtractedFolder.FullName -Force | Move-Item -Destination $DownloadPath -Force
        Write-Log "Moved contents to $DownloadPath"
    } else {
        Write-Log "No extracted folder found in $TempExtractPath" -Level "ERROR"
    }
} catch {
    Write-Log "Failed to move extracted files to $DownloadPath. Error: $_" -Level "ERROR"
    return
}

# Clean up the temporary extraction folder
try {
    Remove-Item -Path $TempExtractPath -Recurse -Force
    Write-Log "Cleaned up temporary extraction directory: $TempExtractPath"
} catch {
    Write-Log "Failed to clean up temporary extraction directory: $TempExtractPath. Error: $_" -Level "ERROR"
}

# Clean up the downloaded ZIP file
try {
    Remove-Item -Path $DownloadFileName -Force
    Write-Log "Cleaned up ZIP file: $DownloadFileName"
} catch {
    Write-Log "Failed to clean up ZIP file: $DownloadFileName. Error: $_" -Level "ERROR"
}

# Change directory to $DownloadPath for any further operations
try {
    Set-Location -Path $DownloadPath
    Write-Log "Changed directory to $DownloadPath"
} catch {
    Write-Log "Failed to change directory to $DownloadPath. Error: $_" -Level "ERROR"
}

# Get the driver file (first *.inf file found)
$inf = Get-ChildItem -Path "$DownloadPath" -Recurse -Filter "*.inf" |
    Where-Object Name -NotLike "Autorun.inf" |
    Select-Object -First 1 |
    Select-Object -ExpandProperty FullName

if (-not $inf) {
    Write-Log "No driver file found." -Level "ERROR"
    Exit
}

function Install-PrinterDrivers {
    # Install the driver
    try {
        PNPUtil.exe /add-driver $inf /install
        Write-Log "Printer driver installed from $inf"
    } catch {
        Write-Log "Printer driver is already loaded into the system drivers" -Level "WARN"
    }

    # Retrieve driver info using DISM
    $DismInfo = Dism.exe /online /Get-DriverInfo /driver:$inf

    # Retrieve the printer driver name from DISM output
    $DriverName = ($DismInfo | Select-String -Pattern "Description" | Select-Object -Last 1) -split " : " |
        Select-Object -Last 1

    # Add driver to the list of available printers
    try {
        Add-PrinterDriver -Name $DriverName -Verbose -ErrorAction SilentlyContinue
        Write-Log "Printer driver $DriverName added successfully."
    } catch {
        Write-Log "Printer driver $DriverName is already available on this system." -Level "WARN"
    }

    # Add printer port
    try {
        Add-PrinterPort -Name $PortName -PrinterHostAddress $PrinterIP -ErrorAction SilentlyContinue
        Write-Log "Printer port $PortName created."
    } catch {
        Write-Log "Printer port $PortName is already installed." -Level "WARN"
    }

    # Add the printer
    try {
        Add-Printer -DriverName $DriverName -Name $PrinterName -PortName $PortName -Verbose -ErrorAction SilentlyContinue
        Write-Log "Printer $PrinterName added successfully."
    } catch {
        Write-Log "Printer $PrinterName is already installed." -Level "WARN"
    }
}

function Set-PrinterDefaults {
    # Set printer as default
    try {
        $CimInstance = Get-CimInstance -Class Win32_Printer -Filter "Name='$PrinterName'"
        Invoke-CimMethod -InputObject $CimInstance -MethodName SetDefaultPrinter
        Write-Log "Printer $PrinterName set as default."
    } catch {
        Write-Log "Could not set $PrinterName as default: $($_.Exception.Message)" -Level "WARN"
    }
    
    # Set paper size to A4
    try {
        Set-PrintConfiguration -PrinterName $PrinterName -PaperSize A4
        Write-Log "Paper size set to A4 for $PrinterName."
    } catch {
        Write-Log "Could not set paper size for ${PrinterName}: $($_.Exception.Message)" -Level "WARN"
    }
    
    # Set default color setting (black and white)
    try {
        Set-PrintConfiguration -PrinterName $PrinterName -Color 0
        Write-Log "Color setting set to black and white for $PrinterName."
    } catch {
        Write-Log "Could not set color setting for ${PrinterName}: $($_.Exception.Message)" -Level "WARN"
    }
}

# FUNCTIONS RUN HERE
Install-PrinterDrivers
Set-PrinterDefaults

# Confirm successful installation
if (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue) {
    Write-Log "Printer $PrinterName has installed successfully."
} else {
    Write-Log "Failed to install printer $PrinterName." -Level "ERROR"
}

Remove-Item -Path "$DownloadPath" -Recurse -Force -ErrorAction SilentlyContinue