# ================================================
# Remove Bloatware Script for Action1
# ================================================
# Description:
#   - This script removes unwanted AppX packages and provisioned AppX packages to debloat Windows.
#   - The AppX packages include pre-installed applications and store apps that are unnecessary or unwanted.
#   - Handles special cases for Edge and OneDrive using winget when available
#   - Supports both Windows 10 and Windows 11
#
# Requirements:
#   - Admin rights are required.
#   - The script must be run with administrative privileges to remove AppX and provisioned packages.
# ================================================

$ProgressPreference = 'SilentlyContinue'

# Get Windows version for different handling of app removal
$WinVersion = [System.Environment]::OSVersion.Version.Build

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

function Test-WingetAvailable {
    try {
        $winget = Get-AppxPackage -Name "Microsoft.DesktopAppInstaller"
        return ($null -ne $winget)
    }
    catch {
        return $false
    }
}

function Remove-SpecialApps {
    param (
        [string]$AppName
    )
    
    if ((Test-WingetAvailable) -and ($AppName -in @("Microsoft.OneDrive", "Microsoft.Edge"))) {
        Write-Log "Attempting to remove $AppName using winget..." -Level "INFO"
        try {
            $process = Start-Process -FilePath "winget" -ArgumentList "uninstall --accept-source-agreements --disable-interactivity --id $AppName" -Wait -PassThru -NoNewWindow
            if ($process.ExitCode -eq 0) {
                Write-Log "$AppName removed successfully via winget" -Level "INFO"
                return $true
            }
            else {
                Write-Log "Failed to remove $AppName via winget (Exit code: $($process.ExitCode))" -Level "ERROR"
                return $false
            }
        }
        catch {
            Write-Log "Error removing $AppName via winget: $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    }
    return $false
}

function Remove-AppPackages {
    param (
        [string]$AppName
    )
    
    $AppPattern = "*${AppName}*"
    
    # Handle app removal based on Windows version
    if ($WinVersion -ge 22000) {
        # Windows 11
        try {
            Get-AppxPackage -Name $AppPattern -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction Continue
            Write-Log "$AppName removed for all users" -Level "INFO"
        }
        catch {
            Write-Log "Failed to remove $AppName for all users: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    else {
        # Windows 10
        try {
            # Remove for current user
            Get-AppxPackage -Name $AppPattern | Remove-AppxPackage -ErrorAction Continue
            
            # Remove for all users
            Get-AppxPackage -Name $AppPattern -PackageTypeFilter Main, Bundle, Resource -AllUsers | 
                Remove-AppxPackage -AllUsers -ErrorAction Continue
            
            Write-Log "$AppName removed for all users" -Level "INFO"
        }
        catch {
            Write-Log "Failed to remove ${AppName}: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    # Remove from OS image
    try {
        Get-AppxProvisionedPackage -Online | 
            Where-Object { $_.PackageName -like $AppPattern } | 
            ForEach-Object { 
                Remove-ProvisionedAppxPackage -Online -AllUsers -PackageName $_.PackageName -ErrorAction Stop
                Write-Log "Removed $AppName from OS image" -Level "INFO"
            }
    }
    catch {
        Write-Log "Failed to remove $AppName from OS image: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Main execution
try {
    Write-Log "Starting bloatware removal process..." -Level "INFO"
    Write-Log "Detected Windows build: $WinVersion" -Level "INFO"
    
    $appsList = @(
        # Microsoft apps
        "Microsoft.3DBuilder",
        "Microsoft.549981C3F5F10",  # Cortana app
        "Microsoft.BingFinance",
        "Microsoft.BingFoodAndDrink",
        "Microsoft.BingHealthAndFitness",
        "Microsoft.BingNews",
        "Microsoft.BingSports",
        "Microsoft.BingTranslator",
        "Microsoft.BingTravel",
        "Microsoft.Copilot",
        "Microsoft.GetHelp",  # Required for some Windows 11 Troubleshooters
        "Microsoft.GamingApp",  # Modern Xbox Gaming App
        "Microsoft.Messaging",
        "Microsoft.Microsoft3DViewer",
        "Microsoft.MicrosoftJournal",
        "Microsoft.MicrosoftOfficeHub",
        "Microsoft.MicrosoftPowerBIForWindows",
        "Microsoft.MicrosoftSolitaireCollection",
        "Microsoft.MixedReality.Portal",
        "Microsoft.NetworkSpeedTest",
        "Microsoft.News",
        "Microsoft.Office.OneNote",
        "Microsoft.Office.Sway",
        "Microsoft.OneConnect",
        "Microsoft.OneDrive",  # OneDrive consumer
        "Microsoft.OutlookForWindows",  # New mail app
        "Microsoft.People",  # Required for Mail & Calendar
        "Microsoft.Print3D",
        "Microsoft.SkypeApp",
        "Microsoft.Todos",
        "Microsoft.Windows.DevHome",
        "Microsoft.WindowsCommunicationsApps",  # Mail & Calendar
        "Microsoft.WindowsFeedbackHub",
        "Microsoft.WindowsMaps",
        "Microsoft.WindowsSoundRecorder",
        "Microsoft.XboxApp",  # Old Xbox Console Companion App
        "Microsoft.XboxGamingOverlay",
        "Microsoft.Xbox.TCUI",
        "Microsoft.XboxGameOverlay",
        "Microsoft.XboxIdentityProvider",  # Xbox sign-in framework
        "Microsoft.YourPhone",  # Phone Link
        "Microsoft.ZuneMusic",  # Modern Media Player
        "Microsoft.ZuneVideo",
        "MicrosoftCorporationII.MicrosoftFamily",  # Family Safety App
        "MicrosoftCorporationII.QuickAssist",  # Quick Assist
        "MicrosoftTeams",  # Old MS Teams personal
        "MicrosoftWindows.CrossDevice",  # Phone integration
        "MSTeams",  # New MS Teams app
        

        # Third-party apps
        "ACGMediaPlayer",
        "ActiproSoftwareLLC",
        "AdobeSystemsIncorporated.AdobePhotoshopExpress",
        "Amazon.com.Amazon",
        "AmazonVideo.PrimeVideo",
        "Asphalt8Airborne",
        "AutodeskSketchBook",
        "CaesarsSlotsFreeCasino",
        "COOKINGFEVER",
        "CyberLinkMediaSuiteEssentials",
        "DisneyMagicKingdoms",
        "Disney",
        "DrawboardPDF",
        "Duolingo-LearnLanguagesforFree",
        "EclipseManager",
        "Facebook",
        "FarmVille2CountryEscape",
        "fitbit",
        "Flipboard",
        "HiddenCity",
        "HULULLC.HULUPLUS",
        "iHeartRadio",
        "Instagram",
        "LinkedInforWindows",  # LinkedIn app
        "king.com.BubbleWitch3Saga",
        "king.com.CandyCrushSaga",
        "king.com.CandyCrushSodaSaga",
        "MarchofEmpires",
        "Netflix",
        "NYTCrossword",
        "OneCalendar",
        "PandoraMediaInc",
        "PhototasticCollage",
        "PicsArt-PhotoStudio",
        "Plex",
        "PolarrPhotoEditorAcademicEdition",
        "Royal Revolt",
        "Shazam",
        "Sidia.LiveWallpaper",
        "SlingTV",
        "Spotify",
        "TikTok",
        "TuneInRadio",
        "Twitter",
        "Viber",
        "WinZipUniversal",
        "Wunderlist",
        "XING"
    )

    foreach ($app in $appsList) {
        Write-Log "Processing $app..." -Level "INFO"
        
        # Try special removal for Edge and OneDrive
        if (-not (Remove-SpecialApps -AppName $app)) {
            # Regular AppX removal for other apps
            Remove-AppPackages -AppName $app
        }
    }
    
    Write-Log "Bloatware removal process completed" -Level "INFO"
}
catch {
    Write-Log "Critical error during bloatware removal: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}
