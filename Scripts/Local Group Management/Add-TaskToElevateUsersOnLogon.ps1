<#
Title: Add Scheduled Task to Elevate Users to Administrator Group

.SYNOPSIS
    Creates or removes a scheduled task that automatically elevates specified users to administrators.

.DESCRIPTION
    Creates a scheduled task that runs on any user logon to check for specified users
    and automatically add them to the Administrators group if they exist. The task:
    - Runs in SYSTEM context with highest privileges
    - Is hidden from users and runs in background
    - Handles non-existent users gracefully
    - Supports multiple users via comma-separated list

.PARAMETER User List
    Comma-separated list of users to elevate (e.g. "john_doe, jane_smith")
    Required: Yes
    Validation: Cannot be empty, must contain valid usernames

.PARAMETER Task Action
    Action to perform. Must be "Add" or "Remove"
    Required: Yes
    Validation: Must be one of ["Add", "Remove"]

.NOTES
    Required Action1 Permissions:
        - Run as System/Admin
        - Task Scheduler Access
        - Local User Management Access

    Action1 Configuration:
        Required Parameters:
            - Name: "User List"
              Type: String
              Format: Comma-separated usernames
            
            - Name: "Task Action"
              Type: String
              Options: ["Add", "Remove"]

.OUTPUTS
    Success (exit 0): Task created/removed successfully
    Error (exit 1): Script failed, check Action1 logs
#>

# Action1 Parameters
$UserList = ${User List}         # Comma-separated list of users to elevate
$TaskAction = ${Task Action}     # "Add" to create task, "Remove" to delete it

$LogFilePath = "$env:SystemDrive\LST\Action1.log"

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
# Parameter Validation
# ================================
if ([string]::IsNullOrWhiteSpace($UserList)) {
    Write-Log "Missing required parameter: User List" -Level "ERROR"
    exit 1
}

if ([string]::IsNullOrWhiteSpace($TaskAction)) {
    Write-Log "Missing required parameter: Task Action" -Level "ERROR"
    exit 1
}

if ($TaskAction -notin @("Add", "Remove")) {
    Write-Log "Invalid Task Action: $TaskAction. Must be 'Add' or 'Remove'" -Level "ERROR"
    exit 1
}

# Convert comma-separated string to array and trim whitespace
$UsersToElevate = $UserList.Split(',').Trim()

# Validate we have users after splitting
if ($UsersToElevate.Count -eq 0 -or ($UsersToElevate.Count -eq 1 -and [string]::IsNullOrWhiteSpace($UsersToElevate[0]))) {
    Write-Log "No valid users provided in User List" -Level "ERROR"
    exit 1
}

try {
    $taskName = "ElevateUsersToAdmin_Logon"

    # Handle task removal if specified
    if ($TaskAction -eq "Remove") {
        Write-Log "Attempting to remove auto-elevation task" -Level "INFO"
        
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
            Write-Log "Successfully removed auto-elevation task" -Level "INFO"
        } else {
            Write-Log "No existing auto-elevation task found" -Level "INFO"
        }
        exit 0
    }

    # Create the PowerShell command for elevation
    $userList = $UsersToElevate -join "','"
    $userList = "'$userList'"
    
    $elevationCommand = '$UsersToElevate = @(' + $userList + '); foreach ($user in $UsersToElevate) { $userAccount = Get-LocalUser -Name $user -ErrorAction SilentlyContinue; if ($userAccount) { $isAdmin = (Get-LocalGroupMember -Group ''Administrators'' -Member $user -ErrorAction SilentlyContinue); if (-not $isAdmin) { Add-LocalGroupMember -Group ''Administrators'' -Member $user -ErrorAction SilentlyContinue } } }'
    
    Write-Log "Creating auto-elevation task for users: $($UsersToElevate -join ', ')" -Level "INFO"

    $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>$(Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffffff")</Date>
    <Author>Action1</Author>
    <Description>Elevates specified users to administrator group on any user logon</Description>
    <URI>\$taskName</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>StopExisting</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>true</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT5M</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT1M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass -Command "$elevationCommand"</Arguments>
      <WorkingDirectory>C:\Windows\System32</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

    # Remove existing task if present
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existingTask) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Log "Removed existing task" -Level "INFO"
    }

    # Create and register new task
    $tempXmlPath = "$env:TEMP\$taskName.xml"
    try {
        $taskXml | Out-File -FilePath $tempXmlPath -Encoding Unicode
        Register-ScheduledTask -Xml (Get-Content $tempXmlPath -Raw) -TaskName $taskName -Force
        Write-Log "Successfully created auto-elevation task" -Level "INFO"
    }
    finally {
        if (Test-Path $tempXmlPath) {
            Remove-Item $tempXmlPath -Force
        }
    }

    exit 0
}
catch {
    Write-Log "Failed to manage task: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}