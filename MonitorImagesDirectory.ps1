 # Set up folder paths, file paths, and runtime configurations
$UserHome = [Environment]::GetFolderPath("UserProfile")
$downloadsPath = Join-Path $UserHome "Images"
$copyImagesScriptPath = Join-Path $UserHome "ImageConversionScripts\HEIC2PNG.ps1"
$logPath = Join-Path $UserHome "DownloadsMonitor.log"
$waitTimeInSeconds = 30  # Interval between folder checks
$timeoutPerFileInSeconds = 120  # Timeout per file processing
$startTime = Get-Date

# Load System.Windows.Forms for MessageBox support
Add-Type -AssemblyName System.Windows.Forms

# Logging function
function Write-Action {
    param ([string]$Message)
    Add-Content -Path $logPath -Value "$(Get-Date) - $Message"
}

# Preliminary checks to ensure proper environment setup
if (-not $env:USERDOMAIN) {
    Write-Host "Script must run under logged-in user context. Exiting..."
    Write-Action "Script must run under logged-in user context. Exiting..."
    exit
}

if (-not (Test-Path $downloadsPath)) {
    New-Item -Path $downloadsPath -ItemType Directory | Out-Null
    Write-Action "Images folder created at $downloadsPath"
}

if (-not (Test-Path $copyImagesScriptPath)) {
    Write-Action "Conversion script not found at $copyImagesScriptPath"
    [System.Windows.Forms.MessageBox]::Show(
        "Conversion script not found at: $copyImagesScriptPath",
        "Critical Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

Write-Action "Script started at $(Get-Date)"

# Function to run HEIC2PNG.ps1 script with a timeout
function Run-HEIC2PNGWithTimeout {
    param (
        [string]$FilePath,  # File to process
        [int]$TimeoutInSeconds = 120
    )

    try {
        # Run the HEIC2PNG.ps1 script
        $process = Start-Process -FilePath "powershell.exe" `
                                 -ArgumentList "-File `"$copyImagesScriptPath`" -FilePath `"$FilePath`"" `
                                 -NoNewWindow -PassThru
        if ($process.WaitForExit($TimeoutInSeconds * 1000)) {
            Write-Action "Successfully processed file: $FilePath using HEIC2PNG.ps1"
        } else {
            Write-Warning "Timeout reached for file: $FilePath. Terminating process..."
            Write-Action "Timeout reached for file: $FilePath. Terminating process..."
            Stop-Process -Id $process.Id -Force

            # Display error popup
            [System.Windows.Forms.MessageBox]::Show(
                "Timeout reached for file: $FilePath. Process terminated.",
                "Processing Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    } catch {
        Write-Action "Error occurred while processing file: $FilePath using HEIC2PNG.ps1. $_"

        # Display error popup
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred while processing file: $FilePath. Error details: $_",
            "Processing Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# Folder monitoring function
function Start-MonitorDownloads {
    while ($true) {
        # Refresh file list
        $filesInDownloads = Get-ChildItem $downloadsPath -File 

        if ($filesInDownloads.Count -gt 0) {
            Write-Host "Files detected. Beginning processing..."
            Write-Action "Files detected. Beginning processing..."

            foreach ($file in $filesInDownloads) {
                if (Test-Path -Path $file.FullName) {
                    try {
                        # Process the file with timeout protection
                        Run-HEIC2PNGWithTimeout -FilePath $file.FullName -TimeoutInSeconds $timeoutPerFileInSeconds
                    } catch {
                        Write-Action "Error processing file $($file.FullName): $_"
                        Write-Host "Error processing file $($file.FullName)."
                    }
                } else {
                    Write-Action "File $($file.FullName) no longer exists. Skipping..."
                    Write-Host "File $($file.FullName) no longer exists. Skipping..."
                }
            }

            Write-Host "File processing complete. Returning to monitoring state..."
            Write-Action "File processing complete. Returning to monitoring state..."
        } else {
            Write-Host "No files detected. Monitoring continues..."
            Write-Action "No files detected. Monitoring continues..."
        }

        # Wait before checking the folder again
        Start-Sleep -Seconds $waitTimeInSeconds
    }
}

# Start folder monitoring
Start-MonitorDownloads 
