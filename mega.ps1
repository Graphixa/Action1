# ================================================
# Squawk Deployment v10.0
# ================================================

$ProgressPreference = 'SilentlyContinue'

# COMPUTER NAME PREFIX - e.g. MSFT for Microsoft
$ComputerNamePrefix = "MSFT"

# QUICKACCESS FOLDER SETTINGS
$PinMyDrive = 1
$SharedFolders = ${Shared Drives To Pin}

# TASK SCHEDULER IMPORT SETTINGS
$XMLTaskFiles = ${Import Task Path} # Split the provided paths/URLs into an array (assumes comma-separated URLs)

# WINGET CONFIG SETTINGS
$wingetConfigPath = ${Winget Configuration Path}  # Provide URL or path for the Winget configuration file
$downloadLatestVersions = ${Download Latest Versions}  # Boolean: 1 to download latest versions or 0 to download version info in configuration file

# GOOGLE FONT SETTINGS
$fonts = ${Font List} # Define the list of fonts to install seperated by a comma ( , )
# Example: "notosans, opensans, firasans, merriweather"

# GCPW SETTINGS
$googleEnrollmentToken = ${Enrollment Token}
$domainsAllowedToLogin = ${Domains Allowed To Login}

# PRINTER SETTINGS
$PrinterIP = ${IP Address}
$PrinterName = ${Printer Name}
$DriverFilesURL = "https://github.com/Graphixa/PCL6-Driver-for-Universal-Print/archive/refs/heads/main.zip"

# BRANDING SETTINGS
$wallpaperUrlOrPath = ${Wallpaper Path}
$lockScreenUrlOrPath = ${LockScreenPath}  


# DEFAULT APP ASSOCAITION SETTINGS
$defaultAppAssocPath = ${XML File Path} # Replace this with the URL or local path to a default app associations XML file

# START MENU DEPLOYMENT SETTINGS
$StartMenuBINFile = "https://github.com/Graphixa/Action1/blob/033eef21639ef91fbed0eabb184ba2ce46b32eb0/Configurations/start2.bin"  # Replace with your actual path or URL to Start2.bin file


# Logging Function
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = "$env:SystemDrive\LST-Action1.log",
        [string]$Level = "INFO"  # INFO, WARN, ERROR
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write log entry to the log file
    Add-Content -Path $LogFilePath -Value $logMessage
}

# RegistryTouch Function to Add or Remove Registry Entries
function RegistryTouch {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("add", "remove")]
        [string]$action,

        [Parameter(Mandatory = $true)]
        [string]$path,

        [Parameter(Mandatory = $true)]
        [string]$name,

        [Parameter()]
        [ValidateSet("String", "ExpandString", "Binary", "DWord", "MultiString", "QWord")]
        [string]$type = "String",  # Default to String

        [Parameter()]
        [object]$value  # Changed to object to handle various data types
    )

    try {
        # Detect the base registry hive
        switch -regex ($path) {
            '^HKLM:\\|^HKEY_LOCAL_MACHINE\\' {
                $baseKey = "HKLM:"
                $pathWithoutHive = $path -replace '^HKLM:\\|^HKEY_LOCAL_MACHINE\\', ''
            }
            '^HKCU:\\|^HKEY_CURRENT_USER\\' {
                $baseKey = "HKCU:"
                $pathWithoutHive = $path -replace '^HKCU:\\|^HKEY_CURRENT_USER\\', ''
            }
            '^HKCR:\\|^HKEY_CLASSES_ROOT\\' {
                $baseKey = "HKCR:"
                $pathWithoutHive = $path -replace '^HKCR:\\|^HKEY_CLASSES_ROOT\\', ''
            }
            '^HKU:\\|^HKEY_USERS\\' {
                $baseKey = "HKU:"
                $pathWithoutHive = $path -replace '^HKU:\\|^HKEY_USERS\\', ''
            }
            '^HKCC:\\|^HKEY_CURRENT_CONFIG\\' {
                $baseKey = "HKCC:"
                $pathWithoutHive = $path -replace '^HKCC:\\|^HKEY_CURRENT_CONFIG\\', ''
            }
            default {
                Write-Log "Unsupported registry hive in path: $path" -Level "WARN"
                return
            }
        }

        # Build the full registry path
        $fullKeyPath = $baseKey
        $subKeyParts = $pathWithoutHive -split "\\"

        # Incrementally create any missing registry keys
        foreach ($part in $subKeyParts) {
            $fullKeyPath = Join-Path $fullKeyPath $part
            if (-not (Test-Path $fullKeyPath)) {
                New-Item -Path $fullKeyPath -Force -ErrorAction Stop | Out-Null
            }
        }

        # Now that all parent keys exist, handle the registry value
        if ($action -eq "add") {
            $itemExists = Get-ItemProperty -Path $fullKeyPath -Name $name -ErrorAction SilentlyContinue

            if (-not $itemExists) {
                Write-Log "Registry item added: $name with value: $value" -Level "INFO"
                New-ItemProperty -Path $fullKeyPath -Name $name -Value $value -PropertyType $type -Force -ErrorAction Stop | Out-Null
            } else {
                # Retrieve the current value with proper data type handling
                $currentValue = (Get-ItemProperty -Path $fullKeyPath -Name $name).$name

                # Convert current value and new value to the appropriate types for comparison
                switch ($type) {
                    "DWord" {
                        $currentValue = [int]$currentValue
                        $value = [int]$value
                    }
                    "QWord" {
                        $currentValue = [long]$currentValue
                        $value = [long]$value
                    }
                    "Binary" {
                        $currentValue = [Byte[]]$currentValue
                        $value = [Byte[]]$value
                    }
                    default {
                        $currentValue = [string]$currentValue
                        $value = [string]$value
                    }
                }

                if (-not ($currentValue -eq $value)) {
                    Write-Log "Registry value differs. Updating item: $name from $currentValue to $value" -Level "INFO"
                    Set-ItemProperty -Path $fullKeyPath -Name $name -Value $value -Force -ErrorAction Stop
                } else {
                    Write-Log "Registry item: $name with value: $value already exists. Skipping." -Level "WARN"
                }
            }
        } elseif ($action -eq "remove") {
            # Check if the registry value exists
            if (Get-ItemProperty -Path $fullKeyPath -Name $name -ErrorAction SilentlyContinue) {
                Write-Log "Removing registry item: $name from path: $fullKeyPath" -Level "INFO"
                Remove-ItemProperty -Path $fullKeyPath -Name $name -Force -ErrorAction Stop
            } else {
                Write-Log "Registry item: $name does not exist at path: $fullKeyPath. Skipping." -Level "WARN"
            }
        }
    } catch {
        Write-Log "Error modifying the registry: $($_.Exception.Message)" -Level "ERROR"
    }
}

function Set-ComputerName {
# Set Computer Name from Device Serial Number
$SerialNumber = (Get-CimInstance -class win32_bios).SerialNumber

# Sanitize the serial number - remove spaces and limit length to 15 characters
$SerialNumber = $SerialNumber -replace ' ', ''
$SerialNumber = $SerialNumber.Substring(0, [Math]::Min(15, $SerialNumber.Length))

# Retrieve system type and model information
$systemTest = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty PCSystemType)
$systemModel = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Model)

# Check if the system is a virtual machine by looking at the model
$isVM = $systemModel -match "Virtual|VMware|KVM|Hyper-V|VirtualBox"

# Ensure there is a dash between the company prefix and computer type if a prefix is provided
if ($ComputerNamePrefix) {
    $ComputerNamePrefix += "-"
}

# Test system type - WS = Workstation, NB = Notebook, VM = Virtual Machine
if ($isVM) {
    $computerName = "${CompanyPrefix}VM-$SerialNumber"
    Write-Log "Detected system as Virtual Machine. New computer name: $computerName"
}
elseif ($systemTest -eq 1) {
    $computerName = "${CompanyPrefix}WS-$SerialNumber"
    Write-Log "Detected system as Workstation. New computer name: $computerName"
}
elseif ($systemTest -eq 2) {
    $computerName = "${CompanyPrefix}NB-$SerialNumber"
    Write-Log "Detected system as Notebook. New computer name: $computerName"
}
elseif ($systemTest -eq 3) {
    $computerName = "${CompanyPrefix}WS-$SerialNumber"
    Write-Log "Detected system as Workstation. New computer name: $computerName"
}
else {
    $computerName = "${CompanyPrefix}$SerialNumber"
    Write-Log "System type unrecognized. Defaulting to new computer name: $computerName"
}

# Attempt to rename the computer
try {
    Rename-Computer -NewName $computerName -Force
    Write-Log "Successfully renamed computer to $computerName"
}
catch {
    Write-Log "Error renaming computer: $_" -Level "ERROR"
    Return
}
}

function Remove-Bloatware {

    ================================
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
    
}

function Add-QuickAccessScript {

    # Define the startup folder path
    $scriptGenerateLocation = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\StartUp\"

    # Initialize foldersToPin array
    $foldersToPin = @()

    # Add My Drive if $PinMyDrive is set to 1
    if ($PinMyDrive -eq 1) {
        $foldersToPin += "My Drive"
    }

    # Split the Folders parameter into an array if provided
    if (-not [string]::IsNullOrWhiteSpace($SharedFolders)) {
        $sharedDriveFolders = $SharedFolders -split ',' | ForEach-Object { $_.Trim() }
        $foldersToPin += $sharedDriveFolders
    }

    # If no folders were provided, log and exit
    if (-not $foldersToPin) {
        Write-Log "No folders specified to pin. Exiting script." -Level "WARN"
        return
    }

    # Generated Script Content (the script that will be created)
    $scriptContent = @"
`$ProgressPreference = 'SilentlyContinue'

`$qa = New-Object -ComObject shell.application
`$quickAccessFolder = `$qa.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}')
`$items = `$quickAccessFolder.Items()

# Array of folders to pin
`$foldersToPin = @(
"@

    # Add each folder to the generated script content
    foreach ($folder in $foldersToPin) {
        if ($folder -eq "My Drive") {
            $scriptContent += "`"G:\My Drive`",`n"
        }
        else {
            $scriptContent += "`"G:\Shared drives\$folder`",`n"
        }
    }

    # Complete the script content
    $scriptContent += @"
)

# Unpin all existing folders
foreach (`$item in `$items) {
    `$item.InvokeVerb('unpinfromhome')
}

# Pin new folders to Quick Access
foreach (`$folder in `$foldersToPin) {
    `$QuickAccess = New-Object -ComObject shell.application
    `$QuickAccess.Namespace(`$folder).Self.InvokeVerb('pintohome')
}
"@


    # Define the path of the new script in the startup folder
    $scriptPath = Join-Path $scriptGenerateLocation "PinFoldersToQuickAccess.ps1"

    # Write the script content to the file in the startup folder
    try {
        Set-Content -Path $scriptPath -Value $scriptContent -Force
        Write-Log "Script created successfully at: $scriptPath" -Level "INFO"
    }
    catch {
        Write-Log "Failed to create script: $($_.Exception.Message)" -Level "ERROR"
    }
    
}

function Import-TaskXML {

    $tempTaskFolder = "$env:TEMP\Tasks"
    $taskFiles = $("$XMLTaskFiles" -split ',').Trim() # Split the provided paths/URLs into an array (assumes comma-separated URLs)

    function Import-Task {
        param (
            [string]$taskFile    # Path or URL to the task file (XML)
        )

        # Ensure temp folder for task files exists
        if (-not (Test-Path $tempTaskFolder)) {
            New-Item -Path $tempTaskFolder -ItemType Directory -Force | Out-Null
        }

        # Extract file name and remove the .xml extension
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($taskFile)
        $tempTaskFile = Join-Path $tempTaskFolder "$fileName.xml"

        try {
            # Check if it's a remote URL or a local/network path
            if ($taskFile -match "^https?://") {
                Write-Log "Downloading task file from remote URL: $taskFile" -Level "INFO"
                Invoke-WebRequest -Uri $taskFile -OutFile $tempTaskFile -ErrorAction Stop
            }
            elseif (Test-Path $taskFile) {
                Write-Log "Copying task file from local/network path: $taskFile" -Level "INFO"
                Copy-Item -Path $taskFile -Destination $tempTaskFile -Force
            }
            else {
                Write-Log "The task file does not exist or is inaccessible: $taskFile" -Level "ERROR"
                return
            }

            # Import the task into Task Scheduler
            Write-Log "Importing task into Task Scheduler with name: $fileName" -Level "INFO"
            Register-ScheduledTask -TaskName $fileName -Xml (Get-Content $tempTaskFile | Out-String) -Force | Out-Null

            Write-Log "Successfully imported task: $fileName" -Level "INFO"
        }
        catch {
            Write-Log "Failed to import task: $($_.Exception.Message)" -Level "ERROR"
        }
    }

    # ================================
    # Main Script Logic
    # ================================
    try {
        foreach ($taskFile in $taskFiles) {
            Import-Task -taskFile $taskFile
        }

        Write-Log "Scheduled task(s) import complete." -Level "INFO"
    }
    catch {
        Write-Log "An error occurred while importing scheduled tasks: $($_.Exception.Message)" -Level "ERROR"
    }
    finally {
        # Clean up the temp folder
        try {
            if (Test-Path $tempTaskFolder) {
                Remove-Item -Path $tempTaskFolder -Recurse -Force
                Write-Log "Temporary task files cleaned up." -Level "INFO"
            }
        }
        catch {
            Write-Log "Failed to clean up temp folder: $($_.Exception.Message)" -Level "ERROR"
        }
    }    
}

function Install-AppsFromConfig {

    $downloadLocation = "$env:temp\winget-import"  # Path to store the Winget config file temporarily, e.g., "$env:temp\winget-import"

    function Get-WinGetExecutable {
        $winget = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq 'Microsoft.DesktopAppInstaller' }

        if ($null -ne $winget) {
            $wingetFilePath = Join-Path -Path $($winget.InstallLocation) -ChildPath 'winget.exe'
            $wingetFile = Get-Item -Path $wingetFilePath
            return $wingetFile
        }
        else {
            Write-Log 'The WinGet executable is not detected, please proceed to Microsoft Store to update the Microsoft.DesktopAppInstaller application.' -Level "WARN"
            return $false
        }
    }

    function Get-WingetConfigFile {
        param (
            [string]$configPath, # URL or path to the configuration file
            [string]$fileName      # File name for saving the configuration file locally
        )

        if (-not (Test-Path $downloadLocation)) {
            New-Item -Path $downloadLocation -ItemType Directory -Force | Out-Null
        }

        $localConfigPath = Join-Path $downloadLocation $fileName

        try {
            if ($configPath -match "^https?://") {
                Write-Log "Downloading configuration file from URL: $configPath" -Level "INFO"
                Invoke-WebRequest -Uri $configPath -OutFile $localConfigPath -ErrorAction Stop
            }
            elseif (Test-Path $configPath) {
                Write-Log "Copying configuration file from local/network path: $configPath" -Level "INFO"
                Copy-Item -Path $configPath -Destination $localConfigPath -Force
            }
            else {
                throw "The configuration file does not exist or is inaccessible: $configPath"
            }

            return $localConfigPath
        }
        catch {
            Write-Log "Failed to download or copy configuration file: $($_.Exception.Message)" -Level "ERROR"
            throw
        }
    }


    function Validate-WingetConfig {
        param (
            [string]$configFile
        )

        try {
            $configContent = Get-Content -Path $configFile -Raw
            $configJson = $configContent | ConvertFrom-Json

            # Check if the file contains the expected "Packages" and "SourceDetails"
            if (-not $configJson.Sources.Packages) {
                throw "The configuration file is missing the 'Packages' section."
            }
            if (-not $configJson.Sources.SourceDetails) {
                throw "The configuration file is incorrect or malformed."
            }

            return $true
        }
        catch {
            Write-Log "Invalid Winget configuration file: $($_.Exception.Message)" -Level "ERROR"
            throw
        }
    }

    # ================================
    # Main Script Logic
    # ================================
    try {
        Write-Log "Starting Winget application installation process." -Level "INFO"

        # Step 1: Check for WinGet executable
        $wingetExe = Get-WinGetExecutable
        if (-not $wingetExe) {
            throw "Winget executable not found. Exiting script."
        }

        # Step 2: Download or access the configuration file
        $wingetConfigFile = Get-WingetConfigFile -configPath $wingetConfigPath -fileName "winget-config.json"

        # Step 3: Validate the Winget configuration file
        if (-not (Validate-WingetConfig -configFile $wingetConfigFile)) {
            throw "The configuration file is invalid. Exiting script."
        }

        # Step 4: Install applications via Winget
        try {
            Write-Log "Installing applications via Winget using the config file." -Level "INFO"
        
            # Set flag for ignoring versions in config file
            $ignoreVersions = ""
            if ($downloadLatestVersions -eq 1) {
                $ignoreVersions = "--ignore-versions"
            }

            # Install applications
            & $wingetExe.FullName import -i $wingetConfigFile --accept-package-agreements --accept-source-agreements $ignoreVersions --ignore-unavailable 

            Write-Log "Applications installed successfully via Winget." -Level "INFO"
        }
        catch {
            Write-Log "Error occurred during Winget application installation: $($_.Exception.Message)" -Level "ERROR"
            throw
        }

        Write-Log "Winget application installation complete." -Level "INFO"
    }
    catch {
        Write-Log "An error occurred during the Winget setup: $($_.Exception.Message)" -Level "ERROR"
    }
    
}

function Install-GoogleFonts {

    $tempDownloadFolder = "$env:TEMP\google_fonts"

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
    }
    catch {
        Write-Log "Error installing fonts: $($_.Exception.Message)" -Level "ERROR"
    }
    
}

function Install-GoogleMDM {
    
    function Test-ProgramInstalled {
        param(
            [string]$ProgramName
        )

        $InstalledSoftware = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" |
        ForEach-Object { [PSCustomObject]@{ 
                DisplayName    = $_.GetValue('DisplayName')
                DisplayVersion = $_.GetValue('DisplayVersion')
            } }

        return $InstalledSoftware | Where-Object { $_.DisplayName -like "*$ProgramName*" }
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
            Write-Log "Google Chrome Enterprise is already installed. Skipping installation." -Level "INFO"
        }
        else {
            Write-Log "Downloading Google Chrome Enterprise..." -Level "INFO"
            Invoke-WebRequest -Uri $chromeUrl -OutFile "$env:TEMP\$chromeFileName" | Out-Null

            try {
                $arguments = "/i `"$env:TEMP\$chromeFileName`" /qn"
                $installProcess = Start-Process msiexec.exe -ArgumentList $arguments -PassThru -Wait

                if ($installProcess.ExitCode -eq 0) {
                    Write-Log "Google Chrome Enterprise installed successfully." -Level "INFO"
                }
                else {
                    Write-Log "Failed to install Google Chrome Enterprise. Exit code: $($installProcess.ExitCode)" -Level "ERROR"
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
            Write-Log "GCPW is already installed. Skipping..." -Level "INFO"
        }
        else {
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
                    }
                    else {
                        Write-Log 'No domains provided. Skipping domain configuration.' -Level "INFO"
                    }
                }
                else {
                    Write-Log "Failed to install GCPW. Exit code: $($installProcess.ExitCode)" -Level "ERROR"
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
            Write-Log 'Google Drive is already installed. Skipping...' -Level "INFO"
        }
        else {
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
            }
            catch {
                Write-Log "Failed to install Google Drive: $($_.Exception.Message)" -Level "ERROR"
            }
            finally {
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
    }
    catch {
        Write-Log "Error during deployment: $($_.Exception.Message)" -Level "ERROR"
    }


}

function Install-Missive {
    
    $softwareName = 'Missive'
    $checkLocation = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
    $tempPath = "$env:SystemDrive\Temp\"
    $jsonFile = $tempPath + "latest.json"

    # Check if Missive is already installed
    if (Get-ChildItem $checkLocation -Recurse -ErrorAction Stop | Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -Match "^$softwareName.*" }) {
        Write-Log "$softwareName is already installed. No action required." -Level "INFO"
        return
    }

    Write-Log "$softwareName is NOT installed. Installing Now..." -Level "INFO"

    # URL for the latest Missive version JSON file (Hosted by Missiveapp.com)
    $jsonUrl = 'https://missiveapp.com/download/latest.json'

    # Create Temp Folder in the Root of the System Drive
    try {
        New-Item -ItemType Directory -Force -Path $tempPath
        Write-Log "Created temporary directory: $tempPath" -Level "INFO"
    }
    catch {
        Write-Log "Failed to create temp directory: $tempPath" -Level "ERROR"
        return
    }

    # Download JSON file to Temp Folder Location
    try {
        Invoke-RestMethod -Method Get -Uri $jsonUrl -OutFile $jsonFile
        Write-Log "Downloaded JSON file for Missive version info." -Level "INFO"
    }
    catch {
        Write-Log "Failed to download JSON file for Missive." -Level "ERROR"
        Write-Log $_.Exception.Message -Level "ERROR"
        return
    }

    # Search JSON file and return the URL of the latest Missive release for Windows
    $json = (Get-Content $jsonFile -Raw) | ConvertFrom-Json
    $version = $json.version
    $windowsDL = $json.downloads.windows.direct
    $missiveFile = "$tempPath\$softwareName-$version.exe"

    # Download latest Missive version for Windows
    try {
        Write-Log "Downloading Missive Installer from $windowsDL" -Level "INFO"
        Invoke-WebRequest -Method Get -Uri $windowsDL -OutFile $missiveFile
        Write-Log "Download completed: $missiveFile" -Level "INFO"
    }
    catch {
        Write-Log "Failed to download Missive installer." -Level "ERROR"
        Write-Log $_.Exception.Message -Level "ERROR"
        return
    }

    # Install Missive silently for all users
    try {
        Write-Log "Installing Missive for all users." -Level "INFO"
        Start-Process -Wait -FilePath $missiveFile -ArgumentList "/S /D=$env:SystemDrive\$softwareName" -PassThru
        Write-Log "Missive installation completed." -Level "INFO"
    }
    catch {
        Write-Log "Missive installation failed." -Level "ERROR"
        Write-Log $_.Exception.Message -Level "ERROR"
        return
    }

    # Create Shortcuts for all users
    $missiveExecutable = "$env:SystemDrive\$softwareName\Missive.exe"
    $startMenuPath = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Missive.lnk"
    $desktopPath = "$env:Public\Desktop\Missive.lnk"
    $WScriptObj = New-Object -ComObject ("WScript.Shell")

    try {
        Write-Log "Creating desktop and start menu shortcuts for all users." -Level "INFO"
    
        # Create Shortcut in All-Users Start Menu
        $shortcutStart = $WscriptObj.CreateShortcut($startMenuPath)
        $shortcutStart.TargetPath = $missiveExecutable
        $shortcutStart.Save()

        # Create Shortcut on All-Users Desktop
        $shortcutDesktop = $WscriptObj.CreateShortcut($desktopPath)
        $shortcutDesktop.TargetPath = $missiveExecutable
        $shortcutDesktop.Save()

        Write-Log "Shortcuts created successfully." -Level "INFO"
    }
    catch {
        Write-Log "Failed to create desktop and start menu shortcuts." -Level "ERROR"
        Write-Log $_.Exception.Message -Level "ERROR"
    }

    # Cleanup temporary files at the end
    try {
        Write-Log "Cleaning up temporary files and folders." -Level "INFO"
        Remove-Item -Path $tempPath -Recurse -Force
        Write-Log "Temporary files and folders cleaned up successfully." -Level "INFO"
    }
    catch {
        Write-Log "Failed to remove temporary files and folders." -Level "ERROR"
        Write-Log $_.Exception.Message -Level "ERROR"
    }

}

function Install-PrinterRicoh {

    # Define Variables
    $PortName = "TCPPort:${PrinterIP}"
    $TempDownloadFolder = "$env:TEMP"
    $DownloadFileName = "$TempDownloadFolder\PrinterDriver.zip"
    $DownloadPath = "$TempDownloadFolder\PrinterDriver"
    $TempExtractPath = "$TempDownloadFolder\TempArchive"



    # Create $DownloadPath directory if it doesn't exist
    if (-not (Test-Path -Path $DownloadPath)) {
        New-Item -Path $DownloadPath -ItemType Directory
        Write-Log "Created directory: $DownloadPath"
    }
    else {
        Write-Log "Directory already exists: $DownloadPath"
    }

    # Create temporary extraction directory if it doesn't exist
    if (-not (Test-Path -Path $TempExtractPath)) {
        New-Item -Path $TempExtractPath -ItemType Directory
        Write-Log "Created temporary extraction directory: $TempExtractPath"
    }
    else {
        Write-Log "Temporary extraction directory already exists: $TempExtractPath"
    }

    # Download the ZIP file
    try {
        Write-Log "Downloading file from: $DriverFilesURL"
        Invoke-WebRequest -Uri $DriverFilesURL -OutFile $DownloadFileName
        Write-Log "Download completed: $DownloadFileName"
    }
    catch {
        Write-Log "Failed to download the ZIP file from $DriverFilesURL. Error: $_" -Level "ERROR"
        return
    }

    # Extract the ZIP file to the temporary extraction folder
    try {
        Expand-Archive -Path $DownloadFileName -DestinationPath $TempExtractPath -Force
        Write-Log "Extraction completed to $TempExtractPath"
    }
    catch {
        Write-Log "Failed to extract the ZIP file. Error: $_" -Level "ERROR"
        return
    }

    # Move the contents from the temporary extraction folder to $DownloadPath
    try {
        $ExtractedFolder = Get-ChildItem -Path $TempExtractPath | Select-Object -First 1
        if ($ExtractedFolder) {
            Get-ChildItem -Path $ExtractedFolder.FullName -Force | Move-Item -Destination $DownloadPath -Force
            Write-Log "Moved contents to $DownloadPath"
        }
        else {
            Write-Log "No extracted folder found in $TempExtractPath" -Level "ERROR"
        }
    }
    catch {
        Write-Log "Failed to move extracted files to $DownloadPath. Error: $_" -Level "ERROR"
        return
    }

    # Clean up the temporary extraction folder
    try {
        Remove-Item -Path $TempExtractPath -Recurse -Force
        Write-Log "Cleaned up temporary extraction directory: $TempExtractPath"
    }
    catch {
        Write-Log "Failed to clean up temporary extraction directory: $TempExtractPath. Error: $_" -Level "ERROR"
    }

    # Clean up the downloaded ZIP file
    try {
        Remove-Item -Path $DownloadFileName -Force
        Write-Log "Cleaned up ZIP file: $DownloadFileName"
    }
    catch {
        Write-Log "Failed to clean up ZIP file: $DownloadFileName. Error: $_" -Level "ERROR"
    }

    # Change directory to $DownloadPath for any further operations
    try {
        Set-Location -Path $DownloadPath
        Write-Log "Changed directory to $DownloadPath"
    }
    catch {
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
        }
        catch {
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
        }
        catch {
            Write-Log "Printer driver $DriverName is already available on this system." -Level "WARN"
        }

        # Add printer port
        try {
            Add-PrinterPort -Name $PortName -PrinterHostAddress $PrinterIP -ErrorAction SilentlyContinue
            Write-Log "Printer port $PortName created."
        }
        catch {
            Write-Log "Printer port $PortName is already installed." -Level "WARN"
        }

        # Add the printer
        try {
            Add-Printer -DriverName $DriverName -Name $PrinterName -PortName $PortName -Verbose -ErrorAction SilentlyContinue
            Write-Log "Printer $PrinterName added successfully."
        }
        catch {
            Write-Log "Printer $PrinterName is already installed." -Level "WARN"
        }
    }

    function Set-PrinterDefaults {
        # Set printer as default
        try {
            $CimInstance = Get-CimInstance -Class Win32_Printer -Filter "Name='$PrinterName'"
            Invoke-CimMethod -InputObject $CimInstance -MethodName SetDefaultPrinter
            Write-Log "Printer $PrinterName set as default."
        }
        catch {
            Write-Log "Could not set $PrinterName as default: $($_.Exception.Message)" -Level "WARN"
        }
    
        # Set paper size to A4
        try {
            Set-PrintConfiguration -PrinterName $PrinterName -PaperSize A4
            Write-Log "Paper size set to A4 for $PrinterName."
        }
        catch {
            Write-Log "Could not set paper size for ${PrinterName}: $($_.Exception.Message)" -Level "WARN"
        }
    
        # Set default color setting (black and white)
        try {
            Set-PrintConfiguration -PrinterName $PrinterName -Color 0
            Write-Log "Color setting set to black and white for $PrinterName."
        }
        catch {
            Write-Log "Could not set color setting for ${PrinterName}: $($_.Exception.Message)" -Level "WARN"
        }
    }

    # FUNCTIONS RUN HERE
    Install-PrinterDrivers
    Set-PrinterDefaults

    # Confirm successful installation
    if (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue) {
        Write-Log "Printer $PrinterName has installed successfully."
    }
    else {
        Write-Log "Failed to install printer $PrinterName." -Level "ERROR"
    }

    Remove-Item -Path "$DownloadPath" -Recurse -Force -ErrorAction SilentlyContinue

}

function Set-Branding {

    $downloadLocation = "$env:SystemDrive\Squawk" # Path to download and keep wallpaper/lockscreen files *if required*

    function Get-Image {
        param (
            [string]$imagePath, # URL or path to the image file
            [string]$fileName     # File name for saving the image locally
        )

        $localImagePath = if ($imagePath -match "^https?://") {
            # Only create download location for URLs
            if (-not (Test-Path $downloadLocation)) {
                New-Item -Path $downloadLocation -ItemType Directory -Force | Out-Null
            }
            Join-Path $downloadLocation $fileName
        }
        else {
            $imagePath
        }

        try {
            if ($imagePath -match "^https?://") {
                Write-Log "Downloading image from URL: $imagePath" -Level "INFO"
                Invoke-WebRequest -Uri $imagePath -OutFile $localImagePath -ErrorAction Stop
            }
            elseif (-not (Test-Path $imagePath)) {
                throw "The image file does not exist or is inaccessible: $imagePath"
            }

            return $localImagePath
        }
        catch {
            Write-Log "Failed to download or copy image: $($_.Exception.Message)" -Level "ERROR"
            throw
        }
    }

    # ================================
    # Main Script Logic
    # ================================
    try {
        Write-Log "Starting Wallpaper and Lock Screen configuration." -Level "INFO"

        # Set Wallpaper
        if ($wallpaperUrlOrPath) {
            try {
                $localWallpaperPath = Get-Image -imagePath $wallpaperUrlOrPath -fileName "wallpaper.jpg"
                Write-Log "Setting wallpaper to: $localWallpaperPath"

                $registryPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
                RegistryTouch -action add -path $registryPath -name "DesktopImagePath" -type "String" -value $localWallpaperPath
                RegistryTouch -action add -path $registryPath -name "DesktopImageUrl" -type "String" -value $localWallpaperPath
                RegistryTouch -action add -path $registryPath -name "DesktopImageStatus" -type "DWord" -value 1
                Write-Log "Wallpaper set successfully." -Level "INFO"
            }
            catch {
                Write-Log "Error setting wallpaper: $($_.Exception.Message)" -Level "ERROR"
            }
        }
        else {
            Write-Log "Wallpaper not set. No image path provided." -Level "WARN"
        }

        # Set Lock Screen
        if ($lockScreenUrlOrPath) {
            try {
                $localLockScreenPath = Get-Image -imagePath $lockScreenUrlOrPath -fileName "lockscreen.jpg"
                Write-Log "Setting lock screen image to: $localLockScreenPath"

                $registryPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
                RegistryTouch -action add -path $registryPath -name "LockScreenImagePath" -type "String" -value $localLockScreenPath
                RegistryTouch -action add -path $registryPath -name "LockScreenImageUrl" -type "String" -value $localLockScreenPath
                RegistryTouch -action add -path $registryPath -name "LockScreenImageStatus" -type "DWord" -value 1
                Write-Log "Lock screen image set successfully." -Level "INFO"
            }
            catch {
                Write-Log "Error setting lock screen image: $($_.Exception.Message)" -Level "ERROR"
            }
        }
        else {
            Write-Log "Lock screen image not set. No image path provided." -Level "WARN"
        }

        # Restart Explorer for settings to take effect
        Stop-Process -Name explorer -Force
        Start-Process explorer
        Write-Log "Explorer restarted successfully." -Level "INFO"

    }
    catch {
        Write-Log "An error occurred during the setup: $($_.Exception.Message)" -Level "ERROR"
    }


}

function Set-RegistrySettings {
    
    # ================================
    # Main Script Logic
    # ================================

    try {
        Write-Log "Setting Registry Items for System" -Level "INFO"

        # Disable OneDrive
        RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -name "DisableFileSyncNGSC" -value 1 -type "DWord"
        Write-Log "OneDrive disabled." -Level "INFO"
        try {
            Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
            Write-Log "OneDrive process stopped." -Level "INFO"
        }
        catch {
            Write-Log "OneDrive process was not running or could not be stopped." -Level "WARN"
        }

        # Disable Cortana
        RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -name "AllowCortana" -value 0 -type "DWord"

        # Disable CoPilot
        RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\WindowsCopilot" -name "CopilotEnabled" -value 0 -type "DWord"

        # Disable Privacy Experience (OOBE)
        RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OOBE" -name "DisablePrivacyExperience" -value 1 -type "DWord"

        # Disable 'Meet Now' in Taskbar
        RegistryTouch -action add -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" -name "HideSCAMeetNow" -value 1 -type "DWord"

        # Disable News and Interests
        RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" -name "AllowNewsAndInterests" -value 0 -type "DWord"

        # Disable Personalized Advertising
        RegistryTouch -action add -path "HKLM:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -name "Enabled" -value 0 -type "DWord"

        # Disable Start Menu Suggestions and Windows Advertising
        RegistryTouch -action add -path "HKLM:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -name "SubscribedContent-338389Enabled" -value 0 -type "DWord"

        # Disable First Logon Animation
        RegistryTouch -action add -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -name "EnableFirstLogonAnimation" -value 0 -type "DWord"

        # Disable Lock Screen App Notifications
        RegistryTouch -action add -path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -name "DisableLockScreenAppNotifications" -value 1 -type "DWord"

        # Restart Explorer to apply changes
        try {
            Stop-Process -Name explorer -Force
            Start-Process explorer
            Write-Log "Explorer restarted successfully." -Level "INFO"
        }
        catch {
            Write-Log "Failed to restart Explorer: $($_.Exception.Message)" -Level "ERROR"
        }

    }
    catch {
        Write-Log "An error occurred while setting registry entries: $($_.Exception.Message)" -Level "ERROR"
    }    
}

function Set-DefaultAppAssociations {

    
    # ================================
    # Function: Download File
    # ================================
    function Download-File {
        param (
            [string]$fileURL, # URL to the file to be downloaded
            [string]$destinationPath  # Local path where the file should be saved
        )

        try {
            if ($fileURL -match "^https?://") {
                Write-Log "Downloading file from remote URL: $fileURL" -Level "INFO"
                Invoke-WebRequest -Uri $fileURL -OutFile $destinationPath -ErrorAction Stop
            }
            elseif (Test-Path $fileURL) {
                Write-Log "Copying file from local/network path: $fileURL" -Level "INFO"
                Copy-Item -Path $fileURL -Destination $destinationPath -Force
            }
            else {
                throw "The file does not exist or is inaccessible: $fileURL"
            }
        }
        catch {
            Write-Log "Failed to download or copy the file: $($_.Exception.Message)" -Level "ERROR"
            throw
        }
    }

    # ================================
    # Main Script Logic
    # ================================
    try {
        # Define the temp folder for the downloaded XML file
        $tempFolder = "$env:TEMP\Action1Files"
        if (-not (Test-Path $tempFolder)) {
            New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null
        }
        $xmlFilePath = Join-Path $tempFolder "DefaultAppAssoc.xml"

        # Download the default app associations XML file
        Download-File -fileURL $defaultAppAssocPath -destinationPath $xmlFilePath

        # Apply the default app associations using DISM
        Write-Log "Applying default app associations using DISM." -Level "INFO"
        try {
            Start-Process dism.exe -ArgumentList "/Online /Import-DefaultAppAssociations:$xmlFilePath" -Wait -NoNewWindow
            Write-Log "Default app associations applied successfully." -Level "INFO"
        }
        catch {
            Write-Log "Failed to apply default app associations: $($_.Exception.Message)" -Level "ERROR"
        }

    }
    catch {
        Write-Log "An error occurred during script execution: $($_.Exception.Message)" -Level "ERROR"
    }
    finally {
        # Clean up the temp folder
        try {
            if (Test-Path $tempFolder) {
                Remove-Item -Path $tempFolder -Recurse -Force
                Write-Log "Temporary files cleaned up." -Level "INFO"
            }
        }
        catch {
            Write-Log "Failed to clean up temp folder: $($_.Exception.Message)" -Level "ERROR"
        }
    }  
    
}

function Install-MSOffice {

    # Set the URL for the Office Deployment Tool (ODT)
    $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16731-20398.exe"

    # Set the path for downloading ODT
    $odtPath = $env:TEMP
    $odtFile = "$odtPath/ODTSetup.exe"

    # Download the Office Deployment Tool
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtFile

    # Set the path for the configuration XML file
    $ConfigurationXMLFile = "$odtPath\configuration.xml"

    # Create the configuration XML file
    @"
<Configuration>
<Add OfficeClientEdition="64" Channel="SemiAnnual">
<Product ID="ProPlus2019Retail">
  <Language ID="en-US" />
</Product>
</Add>
  <Updates Enabled="FALSE" />
  <Display Level="Full" AcceptEULA="TRUE" />
  <Setting Id="SETUP_REBOOT" Value="Never" /> 
  <Setting Id="REBOOT" Value="ReallySuppress"/>
</Configuration>
"@ | Out-File $ConfigurationXMLFile


    #Run the Office Deployment Tool setup
    try {
        Write-Log "Deploying Microsoft Office using the ODT" -Level "INFO"
        Start-Process $odtFile -ArgumentList "/quiet /extract:$odtPath" -Wait
    }
    catch {
        Write-Log "Error running the Office Deployment Tool: :$($_.Exception.Message)" -Level "ERROR"
        return
    }
    #Run the O365 install
    try {
        Write-Log "Installing Microsoft Office" -Level "INFO"
        $Silent = Start-Process "$odtPath\Setup.exe" -ArgumentList "/configure `"$ConfigurationXMLFile`"" -Wait -PassThru
    }
    Catch {
        Write-Log "Error running the Office Installer: :$($_.Exception.Message)" -Level "ERROR"
    }

    # Remove temporary files
    Remove-Item $odtFile
    Remove-Item $ConfigurationXMLFile

    
}

function Set-StartMenuLayout {

    
    $tempBinPath = "$env:TEMP\Start2.bin"  # Temp file path for downloaded .bin file
    $destFolderPath = "$env:SystemDrive\Users\Default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState"

    try {
        Write-Log "Starting the process to copy Start2.bin..." -Level "INFO"
    
        # Check if $StartMenuBINFile is a URL or a local/network path
        if ($StartMenuBINFile -match '^http[s]?://') {
            Write-Log "Downloading Start2.bin from URL: $StartMenuBINFile" -Level "INFO"
            try {
                Invoke-WebRequest -Uri $StartMenuBINFile -OutFile $tempBinPath -UseBasicParsing
                Write-Log "Download completed successfully." -Level "INFO"
            }
            catch {
                Write-Log "Failed to download Start2.bin: $($_.Exception.Message)" -Level "ERROR"
                return
            }
        }
        elseif (Test-Path $StartMenuBINFile) {
            # Use local or network path directly
            Write-Log "Using local/network Start2.bin file: $StartMenuBINFile" -Level "INFO"
            $tempBinPath = $StartMenuBINFile
        }
        else {
            Write-Log "Invalid file path or URL: $StartMenuBINFile" -Level "ERROR"
            return
        }

    }
    catch {
        Write-Log "Pre-check failed: $($_.Exception.Message)" -Level "ERROR"
        return
    }

    # ================================
    # Main Script Logic
    # ================================
    try {
        # Create the destination folder if it doesn't exist
        Write-Log "Creating destination folder if it doesn't exist: $destFolderPath" -Level "INFO"
        New-Item -ItemType "directory" -Path $destFolderPath -Force | Out-Null
    
        # Copy the Start2.bin file to the destination
        Write-Log "Copying Start2.bin to $destFolderPath" -Level "INFO"
        Copy-Item -Path $tempBinPath -Destination "$destFolderPath\Start2.bin" -Force | Out-Null
        Write-Log "Start2.bin copied successfully." -Level "INFO"
    
    }
    catch {
        Write-Log "An error occurred while copying Start2.bin: $($_.Exception.Message)" -Level "ERROR"
        return
    }

    # Cleanup Files
    try {
        Write-Log "Cleaning up temporary files..." -LogFilePath $LogFilePath -Level "INFO"
    
        # If the .bin was downloaded, remove the temp file
        if ($StartMenuBINFile -match '^http[s]?://') {
            Remove-Item -Path $tempBinPath -Force -ErrorAction SilentlyContinue
            Write-Log "Temporary files removed." -LogFilePath $LogFilePath -Level "INFO"
        }

    }
    catch {
        Write-Log "Failed to clean up temporary files: $($_.Exception.Message)" -LogFilePath $LogFilePath -Level "ERROR"
    }

    
}

# Main
Set-ComputerName
Remove-Bloatware
Install-GoogleMDM
Set-Branding
Set-RegistrySettings
Set-StartMenuLayout
Set-DefaultAppAssociations
Install-AppsFromConfig
Install-Missive
Install-MSOffice
Install-GoogleFonts
Import-TaskXML
Add-QuickAccessScript
Install-PrinterRicoh
