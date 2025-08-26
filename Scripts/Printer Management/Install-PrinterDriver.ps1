# ================================================
# Printer Installation Script for Action1
# ================================================
# Description:
#   - This script installs a printer using a TCP/IP port and installs the driver from a specified download.
#   - It sets up the printer with default configurations, including paper size and color settings.
#   - Enhanced with robust driver selection, signature verification, and better error handling.
#   - Generic design to work with any printer brand and driver type.
#
# Requirements:
#   - Admin rights are required.
#   - Internet access is required to download the driver.
#

# ================================================

$ProgressPreference = 'SilentlyContinue'

# Action1 Variables
$PrinterIP = ${IP Address}
$PrinterName = ${Printer Name}
$DriverFilesURL = ${Driver Files URL} # URL to download the printer driver ZIP file
$DriverFilesHash = ${Driver Files Hash} # Optional SHA256 hash for integrity verification (format: sha256:hashvalue)

$PortName = "TCPPort:${PrinterIP}"
$LogFilePath = "$env:SystemDrive\LST\Action1.log" # Default log file path   

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
    Write-Output "$Message"
}

# ================================
# Main Script Logic
# ================================

# Quick sanity check on inputs
if ([string]::IsNullOrWhiteSpace($PrinterIP)) {
    Write-Log "Missing required variables (PrinterIP)" -Level "ERROR"
    return
}
if ([string]::IsNullOrWhiteSpace($PrinterName)) {
    Write-Log "Missing required variables (PrinterName)" -Level "ERROR"
    return
}
if ([string]::IsNullOrWhiteSpace($DriverFilesURL)) {
    Write-Log "Missing required variables (DriverFilesURL)" -Level "ERROR"
    return
}

# Idempotency: skip if already present
try {
    if (Get-Printer -Name $PrinterName -ErrorAction Stop) {
        Write-Log "Printer $PrinterName is already installed. Skipping." -Level "INFO"
        Write-Output "Printer $PrinterName is already installed."
        return
    }
} catch {}

# Prepare temporary paths
$TempDownloadFolder = "$env:TEMP"
$ZipPath = Join-Path $TempDownloadFolder "PrinterDriver.zip"
$ExtractRoot = Join-Path $TempDownloadFolder "PrinterDriver_Extract"

try {
    if (Test-Path $ExtractRoot) { 
        Remove-Item $ExtractRoot -Recurse -Force -ErrorAction SilentlyContinue 
    }
    New-Item -Path $ExtractRoot -ItemType Directory -Force | Out-Null
    Write-Log "Prepared extraction directory: $ExtractRoot" -Level "INFO"
} catch {
    Write-Log "Failed to prepare directories: $($_.Exception.Message)" -Level "ERROR"
    return
}

# Download ZIP with TLS1.2 and retries
try { 
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
} catch {}

$ok = $false
for ($i=1; $i -le 3 -and -not $ok; $i++) {
    Write-Log "Downloading printer driver (attempt $i): $DriverFilesURL" -Level "INFO"
    try {
        Invoke-WebRequest -Uri $DriverFilesURL -OutFile $ZipPath -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" -ErrorAction Stop
        $ok = Test-Path $ZipPath
    } catch {
        Write-Log "Download failed: $($_.Exception.Message)" -Level "WARN"
        Start-Sleep -Seconds 5
    }
}

if (-not $ok) {
    Write-Log "Driver download failed after retries." -Level "ERROR"
    return
}
Write-Log "Download completed: $ZipPath" -Level "INFO"

# Optional SHA-256 check (supports 'sha256:...')
if ($DriverFilesHash) {
    try {
        $expected = ($DriverFilesHash -replace '^(?i)sha256:', '').Trim().ToUpperInvariant()
        $actual = (Get-FileHash -Path $ZipPath -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
        Write-Log "ZIP SHA256: $actual" -Level "INFO"
        if ($actual -ne $expected) {
            Write-Log "Hash mismatch. Expected $expected" -Level "ERROR"
            return
        }
    } catch {
        Write-Log "Hash check failed: $($_.Exception.Message)" -Level "ERROR"
        return
    }
}

# Extract ZIP
try {
    Expand-Archive -Path $ZipPath -DestinationPath $ExtractRoot -Force
    Write-Log "Extraction completed to $ExtractRoot" -Level "INFO"
} catch {
    Write-Log "Failed to extract ZIP file: $($_.Exception.Message)" -Level "ERROR"
    return
}

# Find a signed INF with matching catalog (.cat)
$allInfs = Get-ChildItem -Path $ExtractRoot -Recurse -Filter *.inf -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -ne 'Autorun.inf' }
if (-not $allInfs) {
    Write-Log "No INF files found after extraction" -Level "ERROR"
    return
}

# Generic INF file selection - prefer oemsetup.inf, then any valid INF
$ordered = @()
$oem = $allInfs | Where-Object { $_.Name -ieq 'oemsetup.inf' }
if ($oem) { $ordered += $oem }

# Add remaining INFs
$rest = $allInfs | Where-Object { $_.FullName -notin $ordered.FullName }
if ($rest) { $ordered += $rest }

$chosenInf = $null
foreach ($fi in $ordered) {
    # Read CatalogFile value
    $catName = $null
    try {
        $lines = Get-Content -LiteralPath $fi.FullName -Encoding Default -ErrorAction Stop
        $m = $lines | Select-String -Pattern '^\s*CatalogFile(?:\.[^=]+)?\s*=\s*(.+)$' | Select-Object -First 1
        if ($m) { $catName = $m.Matches[0].Groups[1].Value.Trim().Trim('"') }
    } catch {}

    if (-not $catName) { 
        Write-Log "Skipping $($fi.Name): no CatalogFile entry" -Level "WARN"
        continue 
    }

    $catPath = Join-Path -Path (Split-Path $fi.FullName -Parent) -ChildPath $catName
    if (-not (Test-Path -LiteralPath $catPath)) { 
        Write-Log "Skipping $($fi.Name): missing catalog $catName" -Level "WARN"
        continue 
    }

    try {
        $sig = Get-AuthenticodeSignature -LiteralPath $catPath
        $status = [string]$sig.Status
        $signer = if ($sig.SignerCertificate) { $sig.SignerCertificate.Subject } else { $null }
        Write-Log ("INF {0} → CAT {1} | Signature: {2} | Signer: {3}" -f $fi.Name, [IO.Path]::GetFileName($catPath), $status, $signer) -Level "INFO"
        if ($status -eq 'Valid') { $chosenInf = $fi.FullName; break }
    } catch {
        Write-Log "Skipping $($fi.Name): catalog signature check failed: $($_.Exception.Message)" -Level "WARN"
    }
}

if (-not $chosenInf) {
    Write-Log "Aborting: no signed INF/CAT pair found in package." -Level "ERROR"
    return
}

Write-Log "Using signed INF: $chosenInf" -Level "INFO"

# Stage driver into Driver Store (silent if signed)
try {
    $pnputilOut = & pnputil /add-driver "`"$chosenInf`"" /install 2>&1
    $exit = $LASTEXITCODE
    if ($pnputilOut) { Write-Log ($pnputilOut -join "`r`n") -Level "INFO" }
    if ($exit -ne 0 -and $exit -ne 3010) {
        Write-Log "pnputil failed (exit $exit)" -Level "ERROR"
        return
    }
    Write-Log "Driver staged successfully (exit $exit)" -Level "INFO"
} catch {
    Write-Log "pnputil exception: $($_.Exception.Message)" -Level "ERROR"
    return
}

# Determine driver display name(s) from INF, then bind by name
$candidateNames = @()
try {
    $lines = Get-Content -LiteralPath $chosenInf -Encoding Default -ErrorAction Stop
    $drvLine = $lines | Select-String -Pattern '^\s*DrvName\s*=\s*\"(.+?)\"' | Select-Object -First 1
    if ($drvLine) { $candidateNames += $drvLine.Matches[0].Groups[1].Value }
    $coLine = $lines | Select-String -Pattern '^\s*CoDrvName\s*=\s*\"(.+?)\"' | Select-Object -First 1
    if ($coLine) { $candidateNames += $coLine.Matches[0].Groups[1].Value }

    if (-not $candidateNames -or $candidateNames.Count -eq 0) {
        foreach ($m in ($lines | Select-String -Pattern '^\s*\"(.+?)\"\s*=')) {
            $candidateNames += $m.Matches[0].Groups[1].Value
        }
    }
} catch {}

# If no names found in INF, use a minimal generic fallback
if (-not $candidateNames -or $candidateNames.Count -eq 0) {
    $candidateNames = @('Universal Printing PCL6')
    Write-Log "No driver names found in INF, using generic fallback" -Level "WARN"
} else {
    $candidateNames = $candidateNames | Where-Object { $_ } | Select-Object -Unique
    Write-Log "Found driver names in INF: $($candidateNames -join ', ')" -Level "INFO"
}

$driverName = $null
foreach ($n in $candidateNames) {
    try {
        $exists = Get-PrinterDriver -Name $n -ErrorAction SilentlyContinue
        if (-not $exists) {
            Add-PrinterDriver -Name $n -ErrorAction Stop | Out-Null
            Write-Log "Bound driver by name: $n" -Level "INFO"
            $driverName = $n; break
        } else {
            Write-Log "Driver already available: $n" -Level "INFO"
            $driverName = $n; break
        }
    } catch {
        Write-Log "Could not bind '$n': $($_.Exception.Message)" -Level "WARN"
    }
}

if (-not $driverName) {
    Write-Log "No driver name could be bound after staging." -Level "ERROR"
    return
}

# Create TCP port + printer
try {
    if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) {
        Add-PrinterPort -Name $PortName -PrinterHostAddress $PrinterIP -ErrorAction Stop | Out-Null
        Write-Log "Printer port $PortName created" -Level "INFO"
    } else {
        Write-Log "Printer port $PortName already exists" -Level "INFO"
    }

    if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
        Add-Printer -DriverName $driverName -Name $PrinterName -PortName $PortName -ErrorAction Stop | Out-Null
        Write-Log "Printer $PrinterName added ($driverName on $PortName)" -Level "INFO"
    } else {
        Write-Log "Printer $PrinterName already present" -Level "INFO"
    }
} catch {
    Write-Log "Failed to add printer/port: $($_.Exception.Message)" -Level "ERROR"
    return
}

# Set defaults: A4, B&W, set as default
try {
    Set-PrintConfiguration -PrinterName $PrinterName -PaperSize A4 -ErrorAction SilentlyContinue | Out-Null
    Set-PrintConfiguration -PrinterName $PrinterName -Color 0 -ErrorAction SilentlyContinue | Out-Null
    try { 
        Set-Printer -Name $PrinterName -IsDefault $true -ErrorAction SilentlyContinue | Out-Null 
    } catch {}
    Write-Log "Applied defaults: A4 paper, Black & White, set as default" -Level "INFO"
} catch {
    Write-Log "Could not apply default print settings: $($_.Exception.Message)" -Level "WARN"
}

# Cleanup temporary files
try {
    Remove-Item -Path $ExtractRoot -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    Remove-Item -Path $ZipPath -Force -ErrorAction SilentlyContinue | Out-Null
} catch {}

Write-Log "Printer $PrinterName has been installed successfully." -Level "INFO"
Write-Output "Printer $PrinterName has been installed successfully."