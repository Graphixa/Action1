# ================================================
# Remove Bloatware Script for Action1
# ================================================
# Description:
#   - This script removes unwanted AppX packages and provisioned AppX packages to debloat Windows.
#   - The AppX packages include pre-installed applications and store apps that are unnecessary or unwanted.
#
# Requirements:
#   - Admin rights are required.
#   - The script must be run with administrative privileges to remove AppX and provisioned packages.
# ================================================

$ProgressPreference = 'SilentlyContinue'

# ================================
# Logging Function: Write-Log
# ================================

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

# ================================
# Main Script Logic - AppX Package and Provisioned Package Uninstall
# ================================
try {
    Write-Log "Removing pre-installed bloatware..." -Level "INFO"
    
    $appxPackages = @(
        # Microsoft apps
        "Microsoft.3DBuilder",
        "Microsoft.549981C3F5F10",  # Cortana app
        "Microsoft.Copilot",
        "Microsoft.Messaging",
        "Microsoft.BingFinance",
        "Microsoft.BingFoodAndDrink",
        "Microsoft.BingHealthAndFitness",
        "Microsoft.BingNews",
        "Microsoft.BingSports",
        "Microsoft.BingTravel",
        "Microsoft.MicrosoftOfficeHub",
        "Microsoft.MicrosoftSolitaireCollection",
        "Microsoft.News",
        "Microsoft.MixedReality.Portal",
        "Microsoft.Office.OneNote",
        "Microsoft.OutlookForWindows",
        "Microsoft.Office.Sway",
        "Microsoft.OneConnect",
        "Microsoft.People",
        "Microsoft.SkypeApp",
        "Microsoft.Todos",
        "Microsoft.WindowsMaps",
        "Microsoft.ZuneVideo",
        "Microsoft.ZuneMusic",
        "MicrosoftCorporationII.MicrosoftFamily",  # Family Safety App
        "MSTeams",
        "Outlook",  # New Outlook app
        "LinkedInforWindows",  # LinkedIn app
        "Microsoft.XboxApp",
        "Microsoft.XboxGamingOverlay",
        "Microsoft.Xbox.TCUI",
        "Microsoft.XboxGameOverlay",
        "Microsoft.WindowsCommunicationsApps",  # Mail app
        "Microsoft.YourPhone",  # Phone Link (Your Phone)
        "MicrosoftCorporationII.QuickAssist",  # Quick Assist

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
        "RoyalRevolt",
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

    foreach ($package in $appxPackages) {
        # First remove for all users
        $appInstance = Get-AppxPackage -AllUsers -Name $package
        if ($appInstance) {
            try {
                # Uninstall the appx package for all users
                Get-AppxPackage -AllUsers -Name $package | Remove-AppxPackage -AllUsers -ErrorAction Continue
                Write-Log "$package successfully removed." -Level "INFO"
            } catch {
                Write-Log "Failed to remove: ${package}: $($_.Exception.Message)" -Level "ERROR"
            }
        }

        # Then remove provisioned package to prevent reinstallation
        $provPackage = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq $package }
        if ($provPackage) {
            try {
                Remove-AppxProvisionedPackage -Online -PackageName $provPackage.PackageName -ErrorAction Stop
                Write-Log "Provisioned package $package removed successfully." -Level "INFO"
            } catch {
                Write-Log "Failed to remove provisioned package ${package}: $($_.Exception.Message)" -Level "ERROR"
            }
        }
    }
    
    Write-Log "Removal of pre-installed bloatware complete." -Level "INFO"

} catch {
    Write-Log "An error occurred while removing AppX packages: $($_.Exception.Message)" -Level "ERROR"
}
