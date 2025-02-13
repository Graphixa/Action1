# ================================================
# Play Notification Function
# ================================================
# Description:
#   - This script downloads a notification sound file from a specified URL and plays it locally.
#   - This function can be incorporated into other scripts to play a notification sound upon completion.
#
# Requirements:
#   - Admin rights are required.
# ================================================ 

function Play-Notification {
    param (
        [string]$AudioUrl = "https://github.com/Graphixa/Action1/raw/refs/heads/main/notification.wav",  # URL to WAV file
        [string]$LocalFilePath = "$env:windir\Temp\notification.wav"  # Local path to save the file
    )

    # Step 1: Download the file if it doesn't exist locally
    if (-not (Test-Path $LocalFilePath)) {
        try {
            Write-Output "Downloading audio file from $AudioUrl..."
            Invoke-WebRequest -Uri $AudioUrl -OutFile $LocalFilePath
            Write-Output "Audio file downloaded successfully to $LocalFilePath."
        } catch {
            Write-Output "Failed to download audio file. Error: $($_.Exception.Message)"
            return
        }
    } else {
        Write-Output "Audio file already exists locally at $LocalFilePath. Skipping download."
    }

    # Step 2: Validate the downloaded file
    if (-not (Test-Path $LocalFilePath)) {
        Write-Output "Audio file not found at $LocalFilePath. Cannot play audio."
        return
    }

    # Step 3: Play the audio file
    try {
        Write-Output "Playing audio file: $LocalFilePath"

        $PlayWav = New-Object System.Media.SoundPlayer
        $PlayWav.SoundLocation=$LocalFilePath
        
        # Play the .WAV file
        
        $PlayWav.playsync()

        # Wait for playback to complete
        Start-Sleep -Milliseconds 1000

        Write-Output "Audio playback completed."
    } catch {
        Write-Output "Failed to play audio file. Error: $($_.Exception.Message)"
    }
}

Play-Notification