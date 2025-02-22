# ================================================
# Install Google Fonts  Script for Action1
# ================================================
# Description:
#   - This script installs a list of Google Fonts from the Google Fonts GitHub repository.
#
# Requirements:
#   - Admin rights are required.
#   - Internet access is required to download the fonts.
# ================================================

$ProgressPreference = 'SilentlyContinue'


$fonts = ${Font List} # Define the list of fonts to install seperated by a comma ( , )
# Example: "notosans, opensans, firasans, merriweather"

$tempDownloadFolder = "$env:TEMP\google_fonts"


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

# Function to check if a font is installed
function Test-FontInstalled {
    param (
        [string]$FontName
    )

    $fontRegistryPath = "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
    $installedFonts = Get-ItemProperty -Path $fontRegistryPath

    # Normalize the font name to lowercase for case-insensitive partial match
    $normalizedFontName = $FontName.ToLower()

    # Loop through the installed fonts and check if any contains the font name
    foreach ($installedFont in $installedFonts.PSObject.Properties.Name) {
        if ($installedFont.ToLower() -like "*$normalizedFontName*") {
            return $true
        }
    }

    return $false
}

# Function to download fonts from GitHub
function Get-Fonts {
    param (
        [string]$fontName,
        [string]$outputPath
    )

    $githubUrl = "https://github.com/google/fonts"
    $fontRepoUrl = "$githubUrl/tree/main/ofl/$fontName"

    # Create output directory if it doesn't exist
    if (-not (Test-Path -Path $outputPath)) {
        New-Item -ItemType Directory -Path $outputPath | Out-Null
    }

    # Fetch font file URLs from GitHub
    $fontFilesPage = Invoke-WebRequest -Uri $fontRepoUrl -UseBasicParsing
    $fontFileLinks = $fontFilesPage.Links | Where-Object { $_.href -match "\.ttf$" -or $_.href -match "\.otf$" }

    foreach ($link in $fontFileLinks) {
        $fileUrl = "https://github.com" + $link.href.Replace("/blob/", "/raw/")
        $fileName = [System.IO.Path]::GetFileName($link.href)

        # Download font file
        Invoke-WebRequest -Uri $fileUrl -OutFile (Join-Path -Path $outputPath -ChildPath $fileName)
    }

    Write-Log "Download complete. Fonts saved to $outputPath"
}

# Function to register font in the system registry
function RegistryTouch {
    param (
        [string]$action,
        [string]$path,
        [string]$name,
        [string]$value,
        [string]$type
    )

    if ($action -eq "add") {
        New-ItemProperty -Path $path -Name $name -Value $value -PropertyType $type -Force
    }
}

# ================================
# Main Script Logic
# ================================

# Split fonts list into an array
$fontsList = $fonts -split ',' | ForEach-Object { $_.Trim().ToLower() }

try {
    Write-Log "Installing Google Fonts..." -Level "INFO"

    foreach ($fontName in $fontsList) {
        # Correct the font names for the GitHub repository
        $correctFontName = $fontName -replace "\+", ""

        # Check if the font is already installed
        $isFontInstalled = Test-FontInstalled -FontName $correctFontName

        if ($isFontInstalled) {
            Write-Log "Font $correctFontName is already installed. Skipping Download." -Level "INFO"
            continue
        }

        Write-Log "Downloading & Installing $correctFontName from Google Fonts GitHub repository." -Level "INFO"

        # Download the font files
        Get-Fonts -fontName $correctFontName -outputPath $tempDownloadFolder

        # Install the font files
        $allFonts = Get-ChildItem -Path $tempDownloadFolder -Include *.ttf, *.otf -Recurse
        foreach ($font in $allFonts) {
            $fontDestination = Join-Path -Path $env:windir\Fonts -ChildPath $font.Name
            Copy-Item -Path $font.FullName -Destination $fontDestination -Force

            # Use RegistryTouch to register the font in the registry
            RegistryTouch -action add `
                -path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts" `
                -name $font.BaseName `
                -value $font.Name `
                -type "String"
        }

        Write-Log "Font installed: $correctFontName" -Level "INFO"

        # Clean up the downloaded font files
        Remove-Item -Path $tempDownloadFolder -Recurse -Force

    }

    Write-Log "All fonts installed successfully." -Level "INFO"
} catch {
    Write-Log "Error installing fonts: $($_.Exception.Message)" -Level "ERROR"
}