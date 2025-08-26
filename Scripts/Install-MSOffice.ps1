# ================================================
# PowerShell Script: Install Microsoft Office
# ================================================
# Description:
#   Installs Microsoft Office using the Office Deployment Tool (ODT)
#   with preconfigured settings. Launches installation as background process
#   to avoid Action1 timeout issues. Installation continues after script exits.
#
# Requirements:
#   - Admin rights
#   - Internet access
#   - Action1 platform variables: Product Key, Product ID, Excluded Apps, Company Name
#   - Update Channel: Hardcoded to "Current" for maximum compatibility
#
# Configuration:
#   - Set "Excluded Apps" in Action1 console as comma-separated list (e.g., "Lync, OneDrive, OneNote")
#   - Leave "Excluded Apps" empty in Action1 to install all available Office applications
#   - Set "Company Name" for branding in Office applications
#   - Common exclusions: Access, Lync, OneDrive, OneNote, Outlook, Publisher
#
# For a list of compatible Product IDs, see: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/office-suite-issues/product-ids-supported-office-deployment-click-to-run
# Note: This script launches Office installation as a background process and exits.
#       Office installation continues after the script completes to avoid Action1 timeouts.
#       Check the completion marker file for installation details.
# ================================================

# Set error action preference for better error handling
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ====================
# Action1 Variables
# ====================

$OfficeProductKey = ${Product Key} # "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99" Generic test key for testing
$ProductID = ${Product ID} # ProPlusVolume, ProPlus2019Retail, O365ProPlusRetail
$companyName = ${Company Name} # Name of the company to be displayed in the Office applications

# Action1 format: "Lync, OneDrive, OneNote, Outlook, Groove, Publisher, Access" or leave empty to not exclude any apps
$ExcludedAppsString = ${Excluded Apps} 

# Note: Update Channel is hardcoded to "Current" for maximum compatibility
# Advanced users can modify the script to use custom channels if needed
$OfficeChannel = "Current"

# Convert comma-separated string to array, handling various input formats
if ([string]::IsNullOrWhiteSpace($ExcludedAppsString)) {
    $ExcludedApps = @()
    Write-Host "No excluded apps specified - will install all available Office applications"
} else {
    # Split by comma, trim whitespace, and filter out empty entries
    $ExcludedApps = $ExcludedAppsString -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    Write-Host "Parsed excluded apps from Action1: $($ExcludedApps -join ', ')"
}

# ====================
# Configuration Variables
# ====================
$odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16731-20398.exe"

$LogFilePath = "$env:SystemDrive\LST\Action1.log"
$odtPath = $env:TEMP
$odtFile = "$odtPath\ODTSetup.exe"
$ConfigurationXMLFile = "$odtPath\configuration.xml"
$odtExtractPath = "$odtPath\ODT"
$SetupExePath = "$odtExtractPath\Setup.exe"

# ================================
# Logging Function: Write-Log
# ================================
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = $LogFilePath,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    $logFileDirectory = Split-Path -Path $LogFilePath -Parent
    if (!(Test-Path -Path $logFileDirectory)) {
        try {
            New-Item -Path $logFileDirectory -ItemType Directory -Force | Out-Null
        } catch {
            Write-Error "Failed to create log directory: $logFileDirectory. $_"
            return
        }
    }

    if (Test-Path -Path $LogFilePath) {
        $logSize = (Get-Item -Path $LogFilePath -ErrorAction Stop).Length
        if ($logSize -ge 5242880) {
            Remove-Item -Path $LogFilePath -Force -ErrorAction Stop | Out-Null
            Out-File -FilePath $LogFilePath -Encoding utf8 -ErrorAction Stop
            Add-Content -Path $LogFilePath -Value "[$timestamp] [INFO] Log file exceeded 5MB and was recreated."
        }
    }

    Add-Content -Path $LogFilePath -Value $logMessage
    Write-Output "$Message"
}

# ================================
# Validation Functions
# ================================
function Test-RequiredParameters {
    $missingParams = @()
    
    if ([string]::IsNullOrWhiteSpace($OfficeProductKey)) {
        $missingParams += "Product Key"
    }
    if ([string]::IsNullOrWhiteSpace($ProductID)) {
        $missingParams += "Product ID"
    }
    if ([string]::IsNullOrWhiteSpace($OfficeChannel)) {
        $missingParams += "Office Channel"
    }
    if ([string]::IsNullOrWhiteSpace($companyName)) {
        $missingParams += "Company Name"
    }
    if ($missingParams.Count -gt 0) {
        $missingList = $missingParams -join ", "
        throw "Missing required Action1 platform variables: $missingList"
    }
}

# ================================
# Pre-Check Section
# ================================
try {
    Write-Log "Starting Microsoft Office installation script..." -Level "INFO"
    
    # Validate required parameters
    Test-RequiredParameters
    Write-Log "Required parameters validated successfully." -Level "INFO"
    
    # Check for existing Office installation
    Write-Log "Checking for existing Microsoft Office installation..." -Level "INFO"
    $officeInstalled = Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -like "Microsoft Office*" -or $_.Name -like "Office*"
    }

    if ($officeInstalled) {
        Write-Log "Microsoft Office is already installed (Version: $($officeInstalled.Version)). No installation required." -Level "INFO"
        exit 0
    }
    
    Write-Log "No existing Office installation found. Proceeding with installation." -Level "INFO"
    
} catch {
    Write-Log "Pre-check failed: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}

# ================================
# Main Script Logic
# ================================

try {
    # Download Office Deployment Tool
    Write-Log "Downloading Office Deployment Tool from $odtUrl" -Level "INFO"
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($odtUrl, $odtFile)
    
    if (-not (Test-Path $odtFile)) {
        throw "Failed to download Office Deployment Tool - file not found"
    }
    Write-Log "Office Deployment Tool downloaded successfully." -Level "INFO"
    
    # Extract Office Deployment Tool
    Write-Log "Extracting Office Deployment Tool..." -Level "INFO"
    $extractArgs = "/quiet", "/extract:$odtExtractPath"
    $extractProcess = Start-Process -FilePath $odtFile -ArgumentList $extractArgs -Wait -PassThru -NoNewWindow
    
    if ($extractProcess.ExitCode -ne 0) {
        throw "Office Deployment Tool extraction failed with exit code: $($extractProcess.ExitCode)"
    }
    
    if (-not (Test-Path $SetupExePath)) {
        throw "Setup.exe not found after extraction. Extraction may have failed."
    }
    Write-Log "Office Deployment Tool extracted successfully." -Level "INFO"
    
    # Create configuration XML
    Write-Log "Creating Office configuration..." -Level "INFO"
    
    # Build excluded apps XML dynamically
    $excludedAppsXML = ""
    if ($ExcludedApps.Count -gt 0) {
        foreach ($app in $ExcludedApps) {
            $excludedAppsXML += "      <ExcludeApp ID=`"$app`" />`n"
        }
        Write-Log "Excluding $($ExcludedApps.Count) Office applications: $($ExcludedApps -join ', ')" -Level "INFO"
    } else {
        Write-Log "No Office applications excluded - installing all available apps" -Level "INFO"
    }
    
    $configContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$OfficeChannel">
    <Product ID="$ProductID">
      <Language ID="en-us" />
$excludedAppsXML    </Product>
  </Add>
  <Property Name="PIDKEY" Value="$OfficeProductKey" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <Setup Name="Company" Value="$companyName" />
  </AppSettings>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Setting Id="REBOOT" Value="ReallySuppress"/>
  <Logging Level="Standard" Path="C:\Logs" />
</Configuration>
"@
    
    $configContent | Out-File -FilePath $ConfigurationXMLFile -Encoding UTF8
    Write-Log "Configuration file created successfully." -Level "INFO"
    
    # Install Microsoft Office as background process
    Write-Log "Launching Office installation as background process..." -Level "INFO"
    $installArgs = "/configure", "`"$ConfigurationXMLFile`""
    
    # Launch Office installation and detach - don't wait for completion
    # Using Normal window style so user can see installation progress
    $installProcess = Start-Process -FilePath $SetupExePath -ArgumentList $installArgs -WindowStyle Normal -PassThru
    
    
    if ($installProcess) {
        Write-Log "Office installation launched successfully with Process ID: $($installProcess.Id)" -Level "INFO"
        Write-Log "Installation window is now visible - you can monitor progress" -Level "INFO"
        Write-Log "Installation will continue after this script exits." -Level "INFO"
        
        # Create a completion marker file for potential tracking (in temp location)
        $completionMarker = "$odtPath\Office-Installation-Started.txt"
        $markerContent = @"
Office Installation Status
==========================
Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Process ID: $($installProcess.Id)
Configuration: $ConfigurationXMLFile
Excluded Apps: $($ExcludedApps -join ', ')
Product: $ProductID
Channel: $OfficeChannel
Company: $companyName

Note: This script has exited. Office installation continues in background.
To check progress, look for Office processes or check installed programs.
"@
        $markerContent | Out-File -FilePath $completionMarker -Encoding UTF8
        Write-Log "Completion marker created: $completionMarker" -Level "INFO"
        
    } else {
        Write-Log "Failed to launch Office installation process." -Level "ERROR"
        throw "Could not start Office installation process"
    }
    
} catch {
    Write-Log "Error during Microsoft Office installation: $($_.Exception.Message)" -Level "ERROR"
    return
}


Write-Log "Script completed successfully. Office installation launched in background." -Level "INFO"
Write-Log "Check the completion marker file for installation details." -Level "INFO"
exit 0
