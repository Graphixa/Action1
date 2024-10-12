$ProgressPreference = 'SilentlyContinue'

# Set strict mode to catch errors
Set-StrictMode -Version Latest

$domainsAllowedToLogin = "littlesparrows.com.au"
$googleEnrollmentToken = '040301d8-9be4-4509-82ba-205270232bbd'

function CheckIsAdmin {
    $admin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match 'S-1-5-32-544')
    return $admin
}

function Test-ProgramInstalled {
    param(
        [string]$ProgramName
    )

    $InstalledSoftware = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" |
                         ForEach-Object { [PSCustomObject]@{ 
                            DisplayName = $_.GetValue('DisplayName')
                            DisplayVersion = $_.GetValue('DisplayVersion')
                        }}

    # Check if the partial program name exists in the filtered list
    $isProgramInstalled = $InstalledSoftware | Where-Object { $_.DisplayName -like "*$ProgramName*" } 

    if ($isProgramInstalled) {
        # Output the installed programs for reference (optional)
        # $InstalledSoftware | ForEach-Object { Write-Host "$($_.DisplayName) - $($_.DisplayVersion)" }

        # Return true if the program is found
        return $true
    }

    # Return false if the program is not found
    return $false
}


function Install-ChromeEnterprise {
    $chromeFileName = if ([Environment]::Is64BitOperatingSystem) {
        'googlechromestandaloneenterprise64.msi'
    }
    else {
        'googlechromestandaloneenterprise.msi'
    }

    $chromeUrl = "https://dl.google.com/chrome/install/$chromeFileName"
    
    if (Test-ProgramInstalled 'Google Chrome') {
        Write-Host "Google Chrome Enterprise is already installed. Skipping installation." -ForegroundColor "Cyan"
    } 
    else {
        Write-Host "Downloading: Google Chrome Enterprise"
        Invoke-WebRequest -Uri $chromeUrl -OutFile "$env:TEMP\$chromeFileName" | Out-Null

        try {
            $arguments = "/i `"$env:TEMP\$chromeFileName`" /qn"
            $installProcess = Start-Process msiexec.exe -ArgumentList $arguments -PassThru -Wait

            if ($installProcess.ExitCode -eq 0) {
                Write-Host "Google Chrome Enterprise installed successfully." -ForegroundColor "Green"
            }
            else {
                Write-Host "Failed to install Google Chrome Enterprise. Exit code: $($installProcess.ExitCode)" -ForegroundColor "Red"
            }
        }
        finally {
            Remove-Item -Path "$env:TEMP\$chromeFileName" -Force -ErrorAction SilentlyContinue
        }
    }
}


function Install-GCPW {
    $gcpwFileName = if ([Environment]::Is64BitOperatingSystem) {
        'gcpwstandaloneenterprise64.msi'
    }
    else {
        'gcpwstandaloneenterprise.msi'
    }

    $gcpwUrl = "https://dl.google.com/credentialprovider/$gcpwFileName"
    if (Test-ProgramInstalled 'Credential Provider') {
        Write-Output "GCPW already installed. Skipping..."
    }
    else {
        Write-Host "Downloading GCPW from $gcpwUrl"
        Invoke-WebRequest -Uri $gcpwUrl -OutFile "$env:TEMP\$gcpwFileName"

        try {
            $arguments = "/i `"$env:TEMP\$gcpwFileName`" /quiet"
            $installProcess = Start-Process msiexec.exe -ArgumentList $arguments -PassThru -Wait

            if ($installProcess.ExitCode -eq 0) {
                Write-Output "GCPW Installation completed successfully!"
                
                try {
                    $gcpwRegistryPath = 'HKLM:\SOFTWARE\Policies\Google\CloudManagement'
                    New-Item -Path $gcpwRegistryPath -Force -ErrorAction Stop
                    Set-ItemProperty -Path $gcpwRegistryPath -Name "EnrollmentToken" -Value $googleEnrollmentToken -ErrorAction Stop
                }
                catch {
                    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
                    Start-Sleep 6
                }

                Set-ItemProperty -Path "HKLM:\Software\Google\GCPW" -Name "domains_allowed_to_login" -Value $domainsAllowedToLogin
                $domains = Get-ItemPropertyValue -Path "HKLM:\Software\Google\GCPW" -Name "domains_allowed_to_login"
                if ($domains -eq $domainsAllowedToLogin) {
                    Write-Output 'Domains have been set'
                }
            }
            else {
                Write-Output "Failed to install GCPW. Exit code: $($installProcess.ExitCode)"
            }
        }
        finally {
            Remove-Item -Path "$env:TEMP\$gcpwFileName" -Force -ErrorAction SilentlyContinue
        }
    }
}

function Install-GoogleDrive {
    $driveFileName = 'GoogleDriveFSSetup.exe'
    $driveUrl = "https://dl.google.com/drive-file-stream/$driveFileName"
    if (Test-ProgramInstalled 'Google Drive') {
        Write-Output 'Google Drive already installed. Skipping...'
    }
    else {
        Write-Host "Downloading Google Drive from $driveUrl"
        Invoke-WebRequest -Uri $driveUrl -OutFile "$env:TEMP\$driveFileName"

        try {
            Start-Process -FilePath "$env:TEMP\$driveFileName" -Verb runAs -ArgumentList '--silent'
            Write-Output 'Google Drive Installation completed successfully!'
            try {
                Write-Output "Setting Google Drive Configurations"
                $driveRegistryPath = 'HKLM:\SOFTWARE\Google\DriveFS'
                New-Item -Path $driveRegistryPath -Force -ErrorAction Stop
                Set-ItemProperty -Path $driveRegistryPath -Name 'AutoStartOnLogin' -Value 1 -Type DWord -Force -ErrorAction Stop
                Set-ItemProperty -Path $driveRegistryPath -Name 'DefaultWebBrowser' -Value "$env:systemdrive\Program Files\Google\Chrome\Application\chrome.exe" -Type String -Force -ErrorAction Stop
                Set-ItemProperty -Path $driveRegistryPath -Name 'OpenOfficeFilesInDocs' -Value 0 -Type DWord -Force -ErrorAction Stop

                Write-Output 'Google Drive policies have been set'

            }
            catch {
                Write-Output "Google Drive policies have failed to be added to the registry"
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            }
            
        }
        catch {
            Write-Output "Installation failed!"
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            Start-Sleep 6
        }
        finally {
            Remove-Item -Path "$env:TEMP\$driveFileName" -Force -ErrorAction SilentlyContinue
        }
    }
}

try {
    if (-not (CheckIsAdmin)) {
        Write-Output 'Please run as administrator!'
        exit 5
    }

    if ($domainsAllowedToLogin -eq '') {
        Write-Output 'The list of domains cannot be empty! Please edit this script.'
        exit 5
    }

    # Runs the Install Functions
    Install-ChromeEnterprise
    Install-GCPW
    Install-GoogleDrive

    Write-Host "Google Workspace Deployment Complete"
    Start-Sleep 5
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
Start-Sleep 5