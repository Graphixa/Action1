# ================================================
# Install Google Fonts Script for Action1
# ================================================
# Description:
#   - This script installs a list of Google Fonts using the Google Fonts API.
#
# Requirements:
#   - Admin rights are required.
#   - Internet access is required to download the fonts.
#   - Google Fonts API key is required.
# ================================================

$ProgressPreference = 'SilentlyContinue'

$fonts = ${Font List} # Define the list of fonts to install seperated by a comma ( , )
# Example: "Noto Sans, Open Sans, Fira Sans, Merriweather"

# Google Fonts API Key
$GoogleFontsAPIKey = ${API Key}

$tempDownloadFolder = "$env:TEMP\google_fonts"
$LogFilePath = "$env:SystemDrive\LST\Action1.log" # Default log file path

# Add Windows Font Installation API types
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    
    public class FontInstaller {
        [DllImport("gdi32.dll")]
        public static extern int AddFontResource(string lpFileName);
        
        [DllImport("gdi32.dll")]
        public static extern int RemoveFontResource(string lpFileName);
        
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    }
"@

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = $LogFilePath, # Default log file path
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
    Write-Host "$Message"
}

# Function to convert text to title case if fonts in array are not in proper case
function Convert-ToTitleCase {
    param([string]$Text)
    
    # Use PowerShell's built-in ToTitleCase method
    $culture = Get-Culture
    return $culture.TextInfo.ToTitleCase($Text.ToLower())
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

# Function to get font information from Google Fonts API
function Get-FontInfoFromAPI {
    param (
        [string]$FontFamilyName,
        [string]$APIKey
    )

    try {
        $apiUrl = "https://www.googleapis.com/webfonts/v1/webfonts?key=$APIKey&family=$FontFamilyName"
        Write-Log "Fetching font info for '$FontFamilyName' from Google Fonts API" -Level "INFO"
        
        $response = Invoke-WebRequest -Uri $apiUrl -UseBasicParsing -ErrorAction Stop
        $fontData = $response.Content | ConvertFrom-Json
        
        if ($fontData.items -and $fontData.items.Count -gt 0) {
            $font = $fontData.items[0]
            Write-Log "Successfully retrieved font info for '$FontFamilyName'" -Level "INFO"
            return $font
        } else {
            Write-Log "No font found with family name '$FontFamilyName'" -Level "WARN"
            return $null
        }
    }
    catch {
        Write-Log "Failed to fetch font info for '$FontFamilyName' from API: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

# Function to install a font using Windows Font Installation API
function Install-Font {
    param (
        [string]$FontPath,
        [string]$FontFamily,
        [string]$Variant
    )

    try {
        # Create a proper font name for the registry
        $properFontName = if ($Variant -eq "regular") {
            $FontFamily
        } else {
            "$FontFamily $Variant"
        }

        # Copy font to Windows Fonts folder
        $fontName = [System.IO.Path]::GetFileName($FontPath)
        $fontDestination = Join-Path -Path $env:windir\Fonts -ChildPath $fontName
        
        if (-not (Test-Path $fontDestination)) {
            Copy-Item -Path $FontPath -Destination $fontDestination -Force -ErrorAction Stop
            Write-Log "Font file copied to Windows Fonts folder: $fontName" -Level "INFO"
        }

        # Install font using Windows Font Installation API
        $result = [FontInstaller]::AddFontResource($fontDestination)
        
        if ($result -eq 0) {
            Write-Log "Failed to install font using Windows API: $fontName" -Level "ERROR"
            return $false
        }

        # Register font in Windows registry
        $registryPath = "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
        Set-ItemProperty -Path $registryPath -Name $properFontName -Value $fontName -Type String -Force -ErrorAction Stop
        
        # Notify Windows that fonts have changed
        $HWND_BROADCAST = [IntPtr]::new(0xFFFF)
        $WM_FONTCHANGE = 0x001D
        [FontInstaller]::SendMessage($HWND_BROADCAST, $WM_FONTCHANGE, [IntPtr]::Zero, [IntPtr]::Zero)
        
        return $true
        
    } catch {
        Write-Log "Failed to install font ${FontPath}: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# Function to download and install a font family
function Install-FontFromAPI {
    param (
        [object]$FontInfo,
        [string]$OutputPath
    )

    try {
        $fontFamily = $FontInfo.family
        Write-Log "Installing font family: $fontFamily" -Level "INFO"
        
        # Create output directory if it doesn't exist
        if (-not (Test-Path -Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        
        $installedVariants = 0
        
        # Download and install each variant of the font
        foreach ($variant in $FontInfo.variants) {
            if ($FontInfo.files.$variant) {
                $fontUrl = $FontInfo.files.$variant
                $fileName = [System.IO.Path]::GetFileName($fontUrl)
                $localPath = Join-Path -Path $OutputPath -ChildPath $fileName
                
                try {
                    Write-Log "Downloading $variant variant: $fileName" -Level "INFO"
                    Invoke-WebRequest -Uri $fontUrl -OutFile $localPath -UseBasicParsing -ErrorAction Stop
                    
                    # Install the font using Windows Font Installation API
                    if (Install-Font -FontPath $localPath -FontFamily $fontFamily -Variant $variant) {
                        $installedVariants++
                        Write-Log "Successfully installed $variant variant of $fontFamily" -Level "INFO"
                    } else {
                        Write-Log "Failed to install $variant variant of $fontFamily" -Level "WARN"
                    }
                    
                } catch {
                    Write-Log "Failed to download/install $variant variant of ${fontFamily}: $($_.Exception.Message)" -Level "WARN"
                }
            }
        }
        
        if ($installedVariants -gt 0) {
            Write-Log "Successfully installed $installedVariants variant(s) of $fontFamily" -Level "INFO"
            return $true
        } else {
            Write-Log "No variants were successfully installed for $fontFamily" -Level "WARN"
            return $false
        }
        
    } catch {
        Write-Log "Failed to install font family ${fontFamily}: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# ================================
# Main Script Logic
# ================================

# Split fonts list into an array
$fontsList = $fonts -split ',' | ForEach-Object { $_.Trim() }

try {
    Write-Log "Installing Google Fonts via API..." -Level "INFO"
    
    if ([string]::IsNullOrWhiteSpace($GoogleFontsAPIKey) -or $GoogleFontsAPIKey -eq "YOUR_API_KEY_HERE") {
        Write-Log "Google Fonts API key not configured. Please set a valid API key." -Level "ERROR"
        throw "Google Fonts API key not configured"
    }

    $successCount = 0
    $totalFonts = $fontsList.Count

    foreach ($fontName in $fontsList) {
        $fontName = Convert-ToTitleCase -Text $fontName
        try {
            # Check if the font is already installed
            $isFontInstalled = Test-FontInstalled -FontName $fontName

            if ($isFontInstalled) {
                Write-Log "Font $fontName is already installed. Skipping." -Level "INFO"
                continue
            }

            # Get font information from API
            $fontInfo = Get-FontInfoFromAPI -FontFamilyName $fontName -APIKey $GoogleFontsAPIKey
            
            if ($fontInfo) {
                # Install the font
                if (Install-FontFromAPI -FontInfo $fontInfo -OutputPath $tempDownloadFolder) {
                    $successCount++
                    Write-Log "Font: $fontName installed successfully" -Level "INFO"
                } else {
                    Write-Log "Failed to install font $fontName" -Level "ERROR"
                }
            } else {
                Write-Log "Could not retrieve font information for $fontName" -Level "WARN"
            }
            
        } catch {
            Write-Log "Error processing font ${fontName}: $($_.Exception.Message)" -Level "ERROR"
        }
    }

    Write-Log "Font installation completed. Installed $successCount out of $totalFonts fonts." -Level "INFO"
    
    if ($successCount -gt 0) {
        Write-Log "Successfully installed $successCount of $totalFonts fonts" -Level "INFO"
    }

} catch {
    Write-Log "Error installing fonts: $($_.Exception.Message)" -Level "ERROR"
} finally {
    # Clean up temporary files
    try {
        if (Test-Path $tempDownloadFolder) {
            Remove-Item -Path $tempDownloadFolder -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
            Write-Log "Temporary font files cleaned up" -Level "INFO"
        }
    } catch {
        Write-Log "Failed to clean up temporary font files: $($_.Exception.Message)" -Level "WARN"
    }
}