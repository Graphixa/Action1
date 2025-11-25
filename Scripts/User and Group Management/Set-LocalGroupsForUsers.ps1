# ================================================
# Set Local Group Membership for Users script for Action1
# ================================================
# Description:
#   - Adds or removes users from a specified local Windows group
#   - Supports multiple users (comma-separated list)
#   - Handles any valid local group (Administrators, Users, etc.)
#   - Verifies user and group existence before modification
#
# Requirements:
#   - Must run with elevated privileges (SYSTEM or Administrator)
# ================================================

$ProgressPreference = 'SilentlyContinue'

# ====================
# Parameters Section
# ====================

$UserList = ${User List}         # Enter a comma-separated list of usernames. Example: "john_contoso, tim_contoso"
$GroupMemberships = ${Group Memberships} # Enter comma-separated list of groups. Example: "Administrators, Remote Desktop Users"

$LogFilePath = "$env:SystemDrive\LST\Action1.log" # Default log file path

# ================================
# Logging Function: Write-Log
# ================================
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath = $LogFilePath,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    # Ensure the directory for the log file exists
    $logFileDirectory = Split-Path -Path $LogFilePath -Parent
    if (!(Test-Path -Path $logFileDirectory)) {
        try {
            New-Item -Path $logFileDirectory -ItemType Directory -Force | Out-Null
        } catch {
            Write-Error "Failed to create log directory: $logFileDirectory. $_"
            return
        }
    }

    # Check log file size and recreate if too large
    if (Test-Path -Path $LogFilePath) {
        $logSize = (Get-Item -Path $LogFilePath -ErrorAction Stop).Length
        if ($logSize -ge 5242880) {
            Remove-Item -Path $LogFilePath -Force -ErrorAction Stop | Out-Null
            Out-File -FilePath $LogFilePath -Encoding utf8 -ErrorAction Stop
            Add-Content -Path $LogFilePath -Value "[$timestamp] [INFO] Log file exceeded 5MB and was recreated."
        }
    }

    Add-Content -Path $LogFilePath -Value $logMessage
    Write-Output "$Message"
}

# ================================
# Pre-Check Section
# ================================

# Input validation for UserList
if ([string]::IsNullOrWhiteSpace($UserList)) {
    Write-Log "You must specify users to add to the local group" -Level "ERROR"
    exit 1
}

# Input validation for GroupMemberships
if ([string]::IsNullOrWhiteSpace($GroupMemberships)) {
    Write-Log "You must specify a local groups to add the user(s) to" -Level "ERROR"
    exit 1
}

# Convert comma-separated strings to arrays and trim whitespace
$UserList = $UserList.Split(',').Trim()
$DesiredGroups = $GroupMemberships.Split(',').Trim()

# Validate we have users after splitting
if ($UserList.Count -eq 0 -or ($UserList.Count -eq 1 -and [string]::IsNullOrWhiteSpace($UserList[0]))) {
    Write-Log "No valid users provided in User List" -Level "ERROR"
    exit 1
}

# Validate all specified groups exist
$InvalidGroups = @()
foreach ($GroupName in $DesiredGroups) {
    try {
        $null = Get-LocalGroup -Name $GroupName -ErrorAction Stop
    }
    catch {
        $InvalidGroups += $GroupName
    }
}

if ($InvalidGroups.Count -gt 0) {
    Write-Log "The following groups do not exist: $($InvalidGroups -join ', ')" -Level "ERROR"
    exit 1
}

# ================================
# Main Function: Set-UserAdminRights
# ================================
function Set-UserAdminRights {
    Write-Log "Starting user group membership synchronization" -Level "INFO"
    
    try {
        # Get all local users once
        $localUsers = Get-WmiObject -Class "Win32_UserAccount" -Filter "LocalAccount=True" -ErrorAction Stop
        
        # Process each user
        foreach ($targetUser in $UserList) {
            Write-Log "Processing user: $targetUser" -Level "INFO"
            
            # Check if user exists
            $user = $localUsers | Where-Object { $_.Name -eq $targetUser }
            
            if ($user) {
                Write-Log "Found user $targetUser on system" -Level "INFO"
                
                # Get all groups the user is currently a member of
                $currentGroups = (Get-LocalGroupMember -Group * -ErrorAction SilentlyContinue | 
                    Where-Object { $_.Name -like "*\$targetUser" }).Group
                
                # Add user to desired groups they're not already in
                foreach ($groupName in $DesiredGroups) {
                    if ($groupName -notin $currentGroups) {
                        try {
                            Add-LocalGroupMember -Group $groupName -Member $targetUser -ErrorAction Stop
                            Write-Log "Added $targetUser to $groupName group" -Level "INFO"
                        }
                        catch {
                            Write-Log "Failed to add $targetUser to $groupName group: $($_.Exception.Message)" -Level "ERROR"
                        }
                    }
                }
                
                # Remove user from groups they shouldn't be in
                foreach ($currentGroup in $currentGroups) {
                    if ($currentGroup -notin $DesiredGroups) {
                        try {
                            Remove-LocalGroupMember -Group $currentGroup -Member $targetUser -ErrorAction Stop
                            Write-Log "Removed $targetUser from $currentGroup group" -Level "INFO"
                        }
                        catch {
                            Write-Log "Failed to remove $targetUser from $currentGroup group: $($_.Exception.Message)" -Level "ERROR"
                        }
                    }
                }
            }
            else {
                Write-Log "User $targetUser not found on system" -Level "WARN"
            }
        }
    }
    catch {
        Write-Log "An error occurred during the group membership process: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# ================================
# Main Script Logic
# ================================
try {
    Write-Log "Starting user group membership script" -Level "INFO"
    Set-UserAdminRights
    Write-Log "Script completed successfully" -Level "INFO"
}
catch {
    Write-Log "Script failed: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}