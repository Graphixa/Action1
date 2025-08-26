# ================================================
# Modify Registry Settings Script for Action1
# ================================================
# Description:
#   - This script manages registry keys by creating, modifying, or deleting them
#     on a remote system.
#
# Requirements:
#   - Admin rights required.
#   - Designed to run on Windows systems.
#
# Usage:
#
# - Registry Action: Enter 'add' to create or modify a registry key, or 'remove' to delete one.
# - Registry Path: Provide the full registry path (e.g., HKLM:\SOFTWARE\Action1).
# - Registry Key: Specify the name of the registry value to add, modify, or delete (e.g., EnableFeature).
# - Data Type: Specify the data type (e.g., String, DWord, ExpandString, Binary, MultiString, QWord).
# - Data Value: Enter the data value to set (e.g., 1) (ignored for remove action).
#
# ================================================

$ProgressPreference = 'SilentlyContinue'

# ================================
# Parameters Section
# ================================
# Define registry modifications below. Example:
$registryAction = "${Registry Action}"  # Specify "add" or "remove"
$registryPath =  "${Registry Path}" # "HKLM:\SOFTWARE\Action1"
$registryKey = "${Registry Key}" # Key Value name - Example: "MDM"
$registryType = "${Data Type}" # Data Type - Example: String, ExpandString, Binary, DWord, MultiString, QWord
$registryValue = "${Data Value}" # Value to be set for "add" action. Ignored if "remove" is used.

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

function Test-RegistryParameters {
    param (
        [string]$Action,
        [string]$Path,
        [string]$Key,
        [string]$Type,
        [object]$Value
    )
    
    $errors = @()
    
    # Check if parameters are provided
    if ([string]::IsNullOrWhiteSpace($Action)) {
        $errors += "Registry Action parameter is missing or empty"
    }
    
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $errors += "Registry Path parameter is missing or empty"
    }
    
    if ([string]::IsNullOrWhiteSpace($Key)) {
        $errors += "Registry Key parameter is missing or empty"
    }
    
    # Validate action value
    if ($Action -and $Action -notin @("add", "remove")) {
        $errors += "Registry Action must be either 'add' or 'remove', got: '$Action'"
    }
    
    # Validate path format
    if ($Path -and $Path -notmatch '^(HKLM|HKCU|HKCR|HKU|HKCC|HKEY_LOCAL_MACHINE|HKEY_CURRENT_USER|HKEY_CLASSES_ROOT|HKEY_USERS|HKEY_CURRENT_CONFIG):\\.*') {
        $errors += "Registry Path format is invalid. Must start with valid hive (e.g., HKLM:\SOFTWARE\...)"
    }
    
    # Validate data type for add action
    if ($Action -eq "add" -and $Type -and $Type -notin @("String", "ExpandString", "Binary", "DWord", "MultiString", "QWord")) {
        $errors += "Data Type must be one of: String, ExpandString, Binary, DWord, MultiString, QWord"
    }
    
    # Check if value is provided for add action
    if ($Action -eq "add" -and [string]::IsNullOrWhiteSpace($Value)) {
        $errors += "Data Value is required for 'add' action"
    }
    
    return $errors
}

function Convert-ValueToType {
    param (
        [object]$Value,
        [string]$Type
    )
    
    try {
        switch ($Type) {
            "DWord" {
                if ($Value -match '^0x[0-9A-Fa-f]+$') {
                    return [int]::Parse($Value.Substring(2), [System.Globalization.NumberStyles]::HexNumber)
                }
                return [int]$Value
            }
            "QWord" {
                if ($Value -match '^0x[0-9A-Fa-f]+$') {
                    return [long]::Parse($Value.Substring(2), [System.Globalization.NumberStyles]::HexNumber)
                }
                return [long]$Value
            }
            "Binary" {
                if ($Value -match '^[0-9A-Fa-f\s]+$') {
                    $hexString = $Value -replace '\s', ''
                    $bytes = @()
                    for ($i = 0; $i -lt $hexString.Length; $i += 2) {
                        $bytes += [byte]::Parse($hexString.Substring($i, 2), [System.Globalization.NumberStyles]::HexNumber)
                    }
                    return $bytes
                }
                return [Byte[]]$Value
            }
            "MultiString" {
                if ($Value -is [array]) {
                    return $Value
                }
                if ($Value -is [string] -and $Value.Contains(',')) {
                    return $Value -split ','
                }
                return @($Value)
            }
            default {
                return [string]$Value
            }
        }
    }
    catch {
        Write-Log "Error converting value '$Value' to type '$Type': $_" -Level "ERROR"
        throw "Failed to convert value '$Value' to type '$Type': $_"
    }
}

function Set-RegistryModification {
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
        [string]$type = "String",

        [Parameter()]
        [object]$value
    )

    try {
        Write-Log "Starting registry modification: Action=$action, Path=$path, Key=$name, Type=$type" -Level "INFO"
        
        # Detect the base registry hive
        $baseKey = $null
        $pathWithoutHive = $null
        
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
                throw "Unsupported registry hive in path: $path"
            }
        }

        Write-Log "Detected registry hive: $baseKey" -Level "INFO"
        Write-Log "Path without hive: $pathWithoutHive" -Level "INFO"

        # Build the full registry path
        $fullKeyPath = $baseKey
        $subKeyParts = $pathWithoutHive -split "\\" | Where-Object { $_ -ne "" }

        Write-Log "Creating registry path structure..." -Level "INFO"
        
        # Incrementally create any missing registry keys
        foreach ($part in $subKeyParts) {
            $fullKeyPath = Join-Path $fullKeyPath $part
            Write-Log "Checking path: $fullKeyPath" -Level "INFO"
            
            if (-not (Test-Path $fullKeyPath)) {
                Write-Log "Creating registry key: $fullKeyPath" -Level "INFO"
                New-Item -Path $fullKeyPath -Force -ErrorAction Stop | Out-Null
                Write-Log "Successfully created registry key: $fullKeyPath" -Level "INFO"
            } else {
                Write-Log "Registry key already exists: $fullKeyPath" -Level "INFO"
            }
        }

        # Now that all parent keys exist, handle the registry value
        if ($action -eq "add") {
            Write-Log "Processing 'add' action for key: $name" -Level "INFO"
            
            # Convert value to proper type
            $convertedValue = Convert-ValueToType -Value $value -Type $type
            Write-Log "Converted value '$value' to type '$type': $($convertedValue -join ', ')" -Level "INFO"
            
            $itemExists = Get-ItemProperty -Path $fullKeyPath -Name $name -ErrorAction SilentlyContinue

            if (-not $itemExists) {
                Write-Log "Creating new registry value: $name with value: $($convertedValue -join ', ')" -Level "INFO"
                New-ItemProperty -Path $fullKeyPath -Name $name -Value $convertedValue -PropertyType $type -Force -ErrorAction Stop | Out-Null
                Write-Log "Successfully created registry value: $name" -Level "INFO"
            } else {
                Write-Log "Registry value already exists: $name" -Level "INFO"
                
                # Retrieve the current value
                $currentValue = (Get-ItemProperty -Path $fullKeyPath -Name $name).$name
                Write-Log "Current value: $($currentValue -join ', ')" -Level "INFO"
                Write-Log "New value: $($convertedValue -join ', ')" -Level "INFO"

                # Compare values (handle arrays properly)
                $valuesEqual = $false
                if ($type -eq "MultiString" -or $convertedValue -is [array]) {
                    $valuesEqual = ($currentValue -join ',') -eq ($convertedValue -join ',')
                } else {
                    $valuesEqual = $currentValue -eq $convertedValue
                }

                if (-not $valuesEqual) {
                    Write-Log "Updating registry value: $name from '$($currentValue -join ', ')' to '$($convertedValue -join ', ')'" -Level "INFO"
                    Set-ItemProperty -Path $fullKeyPath -Name $name -Value $convertedValue -Force -ErrorAction Stop
                    Write-Log "Successfully updated registry value: $name" -Level "INFO"
                } else {
                    Write-Log "Registry value unchanged: $name already has the correct value" -Level "INFO"
                }
            }
        } elseif ($action -eq "remove") {
            Write-Log "Processing 'remove' action for key: $name" -Level "INFO"
            
            # Check if the registry value exists
            if (Get-ItemProperty -Path $fullKeyPath -Name $name -ErrorAction SilentlyContinue) {
                Write-Log "Removing registry value: $name from path: $fullKeyPath" -Level "INFO"
                Remove-ItemProperty -Path $fullKeyPath -Name $name -Force -ErrorAction Stop | Out-Null
                Write-Log "Successfully removed registry value: $name" -Level "INFO"
            } else {
                Write-Log "Registry value does not exist: $name at path: $fullKeyPath" -Level "INFO"
            }
        }
        
        Write-Log "Registry modification completed successfully" -Level "INFO"
        return $true
        
    } catch {
        Write-Log "Error in Set-RegistryModification: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
        return $false
    }
}

# ================================
# Main Script Logic
# ================================
try {
    Write-Log "=== Starting Registry Modification Script ===" -Level "INFO"
    Write-Log "Script parameters received:" -Level "INFO"
    Write-Log "  Action: '$registryAction'" -Level "INFO"
    Write-Log "  Path: '$registryPath'" -Level "INFO"
    Write-Log "  Key: '$registryKey'" -Level "INFO"
    Write-Log "  Type: '$registryType'" -Level "INFO"
    Write-Log "  Value: '$registryValue'" -Level "INFO"
    
    # Validate parameters
    $validationErrors = Test-RegistryParameters -Action $registryAction -Path $registryPath -Key $registryKey -Type $registryType -Value $registryValue
    
    if ($validationErrors.Count -gt 0) {
        foreach ($validationError in $validationErrors) {
            Write-Log "Parameter validation error: $validationError" -Level "ERROR"
        }
        throw "Parameter validation failed. Please check the script parameters."
    }
    
    Write-Log "Parameter validation passed" -Level "INFO"
    
    # Set default type if not specified
    if ([string]::IsNullOrWhiteSpace($registryType)) {
        $registryType = "String"
        Write-Log "Using default data type: String" -Level "INFO"
    }
    
    # Execute registry modification
    $success = Set-RegistryModification -action $registryAction -path $registryPath -name $registryKey -type $registryType -value $registryValue
    
    if ($success) {
        Write-Log "=== Registry modification completed successfully ===" -Level "INFO"
        Write-Host "Registry modification completed successfully."
        exit 0
    } else {
        Write-Log "=== Registry modification failed ===" -Level "ERROR"
        Write-Error "Registry modification failed. Check the logs for details."
        exit 1
    }
    
} catch {
    Write-Log "=== Critical error in main script ===" -Level "ERROR"
    Write-Log "Error: $($_.Exception.Message)" -Level "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
    Write-Error "Critical error occurred during registry modification: $($_.Exception.Message)"
    exit 1
}