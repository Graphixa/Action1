# ================================================
# Deploy Google Workspace MDM for Action1
# ================================================
# Description:
#   - This script deploys Google Chrome Enterprise, GCPW (Google Credential Provider for Windows), and Google Drive File Stream, with required configurations.
#
# Requirements:
#   - Admin rights are required.
#   - Internet access is required to download the installers.
# ================================================


$ProgressPreference = 'SilentlyContinue'
Set-StrictMode -Version Latest

$googleEnrollmentToken = ${Enrollment Token}
$domainsAllowedToLogin = ${Domains Allowed To Login}

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

# ================================
# Main Script Logic
# ================================
function Test-ProgramInstalled {
    param(
        [string]$ProgramName
    )

    $InstalledSoftware = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" |
                         ForEach-Object { [PSCustomObject]@{ 
                            DisplayName = $_.GetValue('DisplayName')
                            DisplayVersion = $_.GetValue('DisplayVersion')
                        }}

    return $InstalledSoftware | Where-Object { $_.DisplayName -like "*$ProgramName*" }
}

function Install-ChromeEnterprise {
    $chromeFileName = if ([Environment]::Is64BitOperatingSystem) {
        'googlechromestandaloneenterprise64.msi'
    } else {
        'googlechromestandaloneenterprise.msi'
    }
    $chromeUrl = "https://dl.google.com/chrome/install/$chromeFileName"
    
    if (Test-ProgramInstalled 'Google Chrome') {
        Write-Log "Google Chrome Enterprise is already installed. Skipping installation." -Level "INFO"
    } else {
        Write-Log "Downloading Google Chrome Enterprise..." -Level "INFO"
        Invoke-WebRequest -Uri $chromeUrl -OutFile "$env:TEMP\$chromeFileName" | Out-Null

        try {
            $arguments = "/i `"$env:TEMP\$chromeFileName`" /qn"
            $installProcess = Start-Process msiexec.exe -ArgumentList $arguments -PassThru -Wait

            if ($installProcess.ExitCode -eq 0) {
                Write-Log "Google Chrome Enterprise installed successfully." -Level "INFO"
            } else {
                Write-Log "Failed to install Google Chrome Enterprise. Exit code: $($installProcess.ExitCode)" -Level "ERROR"
            }
        } finally {
            Remove-Item -Path "$env:TEMP\$chromeFileName" -Force -ErrorAction SilentlyContinue
        }
    }
}

function Install-GCPW {
    $gcpwFileName = if ([Environment]::Is64BitOperatingSystem) {
        'gcpwstandaloneenterprise64.msi'
    } else {
        'gcpwstandaloneenterprise.msi'
    }
    $gcpwUrl = "https://dl.google.com/credentialprovider/$gcpwFileName"

    if (Test-ProgramInstalled 'Credential Provider') {
        Write-Log "GCPW is already installed. Skipping..." -Level "INFO"
    } else {
        Write-Log "Downloading GCPW from $gcpwUrl" -Level "INFO"
        Invoke-WebRequest -Uri $gcpwUrl -OutFile "$env:TEMP\$gcpwFileName" | Out-Null

        try {
            $arguments = "/i `"$env:TEMP\$gcpwFileName`" /quiet"
            $installProcess = Start-Process msiexec.exe -ArgumentList $arguments -PassThru -Wait

            if ($installProcess.ExitCode -eq 0) {
                Write-Log "GCPW installed successfully." -Level "INFO"

                # Set registry keys for enrollment token
                $gcpwRegistryPath = 'HKLM:\SOFTWARE\Policies\Google\CloudManagement'
                New-Item -Path $gcpwRegistryPath -Force -ErrorAction Stop
                Set-ItemProperty -Path $gcpwRegistryPath -Name "EnrollmentToken" -Value $googleEnrollmentToken -ErrorAction Stop

                # Set domain registry key only if $domainsAllowedToLogin is not null
                if ($domainsAllowedToLogin) {
                    Set-ItemProperty -Path "HKLM:\Software\Google\GCPW" -Name "domains_allowed_to_login" -Value $domainsAllowedToLogin
                    $domains = Get-ItemPropertyValue -Path "HKLM:\Software\Google\GCPW" -Name "domains_allowed_to_login"
                    if ($domains -eq $domainsAllowedToLogin) {
                        Write-Log 'Domains have been set.' -Level "INFO"
                    }
                } else {
                    Write-Log 'No domains provided. Skipping domain configuration.' -Level "INFO"
                }
            } else {
                Write-Log "Failed to install GCPW. Exit code: $($installProcess.ExitCode)" -Level "ERROR"
            }
        } finally {
            Remove-Item -Path "$env:TEMP\$gcpwFileName" -Force -ErrorAction SilentlyContinue
        }
    }
}

function Install-GoogleDrive {
    $driveFileName = 'GoogleDriveFSSetup.exe'
    $driveUrl = "https://dl.google.com/drive-file-stream/$driveFileName"

    if (Test-ProgramInstalled 'Google Drive') {
        Write-Log 'Google Drive is already installed. Skipping...' -Level "INFO"
    } else {
        Write-Log "Downloading Google Drive from $driveUrl" -Level "INFO"
        Invoke-WebRequest -Uri $driveUrl -OutFile "$env:TEMP\$driveFileName" | Out-Null

        try {
            Start-Process -FilePath "$env:TEMP\$driveFileName" -ArgumentList '--silent' -Wait
            Write-Log 'Google Drive installed successfully!' -Level "INFO"

            # Set registry keys for Google Drive configurations
            $driveRegistryPath = 'HKLM:\SOFTWARE\Google\DriveFS'
            New-Item -Path $driveRegistryPath -Force -ErrorAction Stop
            Set-ItemProperty -Path $driveRegistryPath -Name 'AutoStartOnLogin' -Value 1 -Type DWord -Force -ErrorAction Stop
            Set-ItemProperty -Path $driveRegistryPath -Name 'DefaultWebBrowser' -Value "$env:systemdrive\Program Files\Google\Chrome\Application\chrome.exe" -Type String -Force -ErrorAction Stop
            Set-ItemProperty -Path $driveRegistryPath -Name 'OpenOfficeFilesInDocs' -Value 0 -Type DWord -Force -ErrorAction Stop

            Write-Log 'Google Drive policies have been set.' -Level "INFO"
        } catch {
            Write-Log "Failed to install Google Drive: $($_.Exception.Message)" -Level "ERROR"
        } finally {
            Remove-Item -Path "$env:TEMP\$driveFileName" -Force -ErrorAction SilentlyContinue
        }
    }
}

try {
    Write-Log "Deploying Google Workspace MDM..." -Level "INFO"
    
    # Run installation functions
    Install-ChromeEnterprise
    Install-GCPW
    Install-GoogleDrive

    Write-Log "GoogleMDM deployment completed." -Level "INFO"
} catch {
    Write-Log "Error during deployment: $($_.Exception.Message)" -Level "ERROR"
}
