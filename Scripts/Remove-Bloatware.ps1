# ================================================
# Remove Bloatware Script for Action1
# ================================================
# Description:
#   - This script removes unwanted AppX packages to debloat Windows.
#   - The AppX packages include pre-installed applications and store apps that are not necessary.
#
# Requirements:
#   - Admin rights are required.
#   - Script must be run with administrative privileges to remove AppX packages.
# ================================================

$ProgressPreference = 'SilentlyContinue'

# ================================
# Logging Function: Write-Log
# ================================
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

# ================================
# Main Script Logic - AppX Package Uninstall
# ================================
try {
    Write-Log "Removing pre-installed bloatware..." -Level "INFO"
    
    $appxPackages = @(
        # Microsoft apps
        "Microsoft.3DBuilder",
        "Microsoft.549981C3F5F10",  # Cortana app
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
        "Microsoft.Office.Sway",
        "Microsoft.OneConnect",
        "Microsoft.SkypeApp",
        "Microsoft.Todos",
        "Microsoft.WindowsMaps",
        "Microsoft.ZuneVideo",
        "MicrosoftCorporationII.MicrosoftFamily",  # Family Safety App
        "MSTeams",

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
        "LinkedInforWindows",
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
        $appInstance = Get-AppxPackage -AllUsers -Name $package
        if ($appInstance) {
            try {
                # Uninstall the appx package for all users
                Write-Log "Attempting removal of: $package" -Level "INFO"
                $appInstance | Remove-AppxPackage -ErrorAction Stop
                Write-Log "$package successfully removed." -Level "INFO"
            } catch {
                Write-Log "Failed to remove: ${package}: $($_.Exception.Message)" -Level "ERROR"
            }
        }
    }
    
    Write-Log "Removal of pre-installed bloatware complete." -Level "INFO"

} catch {
    Write-Log "An error occurred while removing AppX packages: $($_.Exception.Message)" -Level "ERROR"
}
