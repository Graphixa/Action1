# ================================================
# Add Task in Task Scheduler to Elevate User to Administrator upon their logon
# ================================================
# Description:
#   - Creates a scheduled task that elevates a specific user to administrator when that user logs in
#   - Task triggers on that user's logon
#   - Runs in SYSTEM context
#   - Runs in the background and is hidden from the user.
# ================================================



$UserToElevate = ${Username} # Enter a user name (must be a valid user on the system)
$TaskAction = ${Task Action} # Add or Remove to remove the scheduled task form the system

$LogFilePath = "$env:SystemDrive\LST\Action1.log"

# ================================
# Logging Function: Write-Log
# ================================
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

# Guard Clauses to check if the required variables are set

if ([string]::IsNullOrWhiteSpace($UserToElevate)) {
    Write-Log "Missing required variable: UserToElevate" -Level "ERROR"
    exit 1
}

if ([string]::IsNullOrWhiteSpace($TaskAction)) {
    Write-Log "Missing required variable: TaskAction" -Level "ERROR"
    exit 1
}

if ($TaskAction -ne "Add" -and $TaskAction -ne "Remove") {
    Write-Log "Invalid TaskAction: $TaskAction" -Level "ERROR"
    exit 1
}

try {
    $taskName = "ElevateUser_$UserToElevate"

    # Handle task removal if specified
    if ($TaskAction -eq "Remove") {
        Write-Log "Attempting to remove auto-elevation task for user: $UserToElevate" -Level "INFO"
        
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
            Write-Log "Successfully removed auto-elevation task" -Level "INFO"
        } else {
            Write-Log "No existing auto-elevation task found for $UserToElevate" -Level "INFO"
        }
        exit 0
    }

    # Proceed with task creation
    Write-Log "Preparing to create auto-elevation task for user: $UserToElevate" -Level "INFO"
    Write-Log "Task will grant administrator rights when $UserToElevate logs in" -Level "INFO"

    $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2025-08-26T00:07:51.6000783</Date>
    <Author>Action1</Author>
    <Description>Elevates $UserToElevate to administrator group on their login</Description>
    <URI>\ElevateUser_$UserToElevate</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>$UserToElevate</UserId>
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
      <Arguments>-ExecutionPolicy Bypass -Command "Add-LocalGroupMember -Group 'Administrators' -Member '$UserToElevate'"</Arguments>
      <WorkingDirectory>C:\Windows\System32</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

    # Remove existing task if present
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existingTask) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Log "Cleaned up previous elevation task" -Level "INFO"
    }

    # Register new task
    $taskXml | Out-File "$env:TEMP\$taskName.xml" -Encoding Unicode
    Register-ScheduledTask -Xml (Get-Content "$env:TEMP\$taskName.xml" -Raw) -TaskName $taskName -Force
    Remove-Item "$env:TEMP\$taskName.xml" -Force
    
    Write-Log "Auto-elevation task installed successfully" -Level "INFO"
    exit 0
} catch {
    Write-Log "Failed to create task: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}