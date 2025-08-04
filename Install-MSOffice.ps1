# ================================================
# PowerShell Script: Install Microsoft Office
# ================================================
# Description:
#   Installs Microsoft Office using the Office Deployment Tool (ODT)
#   with preconfigured settings and post-installation verification.
#
# Requirements:
#   - Admin rights
#   - Internet access
# ================================================

$ProgressPreference = 'SilentlyContinue'

# ====================
# Parameters Section
# ====================

$OfficeProductKey = ${Product Key} # "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99" Generic test key for testing
$ProductID = ${Product ID} # ProPlus2019Retail, ProPlusRetail
$OfficeChannel = ${Office Channel} # PerpetualVL2019, PerpetualVL2016, SemiAnnual
$softwareName = 'Microsoft Office'
$LogFilePath = "$env:SystemDrive\LST\Action1.log"
$odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16731-20398.exe"
$odtPath = $env:TEMP
$odtFile = "$odtPath\ODTSetup.exe"
$ConfigurationXMLFile = "$odtPath\configuration.xml"

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
# Pre-Check Section
# ================================
try {
    Write-Log "Checking for existing Microsoft Office installation..." -Level "INFO"
    $officeInstalled = Get-WmiObject -Class Win32_Product | Where-Object {
        $_.Name -like "Microsoft Office*" -or $_.Name -like "Office*"
    }

    if ($officeInstalled) {
        Write-Log "Microsoft Office is already installed (Version: $($officeInstalled.Version)). No installation required." -Level "INFO"
        return
    }
} catch {
    Write-Log "Pre-check failed: $($_.Exception.Message)" -Level "ERROR"
    return
}

# ================================
# Main Script Logic
# ================================
try {
    Write-Log "Downloading Office Deployment Tool from $odtUrl" -Level "INFO"
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtFile -ErrorAction Stop
} catch {
    Write-Log "Failed to download Office Deployment Tool: $($_.Exception.Message)" -Level "ERROR"
    return
}

@"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$OfficeChannel">
    <Product ID="$ProductID">
      <Language ID="en-us" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="Publisher" />
    </Product>
  </Add>
  <Property Name="PIDKEY" Value="$OfficeProductKey" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <Setup Name="Company" Value="LST" />
  </AppSettings>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Setting Id="REBOOT" Value="ReallySuppress"/>
  <Logging Level="Standard" Path="C:\Logs" />
</Configuration>
"@ | Out-File $ConfigurationXMLFile

try {
    Write-Log "Extracting Office Deployment Tool." -Level "INFO"
    Start-Process $odtFile -ArgumentList "/quiet /extract:$odtPath" -Wait
} catch {
    Write-Log "Error running the Office Deployment Tool: $($_.Exception.Message)" -Level "ERROR"
    return
}

try {
    Write-Log "Installing Microsoft Office." -Level "INFO"
    $installProcess = Start-Process "$odtPath\Setup.exe" -ArgumentList "/configure `"$ConfigurationXMLFile`"" -Wait -PassThru

    if ($installProcess.ExitCode -eq 0) {
        Write-Log "Microsoft Office installation completed successfully." -Level "INFO"
        Start-Sleep -Seconds 10
        $verifyOffice = Get-WmiObject -Class Win32_Product | Where-Object {
            $_.Name -like "Microsoft Office*" -or $_.Name -like "Office*"
        }
        if ($verifyOffice) {
            Write-Log "Office installation verified successfully." -Level "INFO"
        } else {
            Write-Log "Office installation may have failed – not found in installed programs." -Level "WARN"
        }
    } else {
        Write-Log "Microsoft Office installation failed with exit code: $($installProcess.ExitCode)." -Level "ERROR"
        return
    }
} catch {
    Write-Log "Error during Microsoft Office installation: $($_.Exception.Message)" -Level "ERROR"
    return
}

# ================================
# Cleanup Section
# ================================
try {
    Write-Log "Cleaning up temporary files..." -Level "INFO"
    Remove-Item $odtFile -Force -ErrorAction SilentlyContinue
    Remove-Item $ConfigurationXMLFile -Force -ErrorAction SilentlyContinue
} catch {
    Write-Log "Failed to remove temporary files: $($_.Exception.Message)" -Level "ERROR"
}

Write-Log "Script completed successfully." -Level "INFO"
