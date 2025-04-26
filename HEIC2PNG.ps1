# Initialize variables
$UserHome = [Environment]::GetFolderPath("UserProfile")
$Desktop = [Environment]::GetFolderPath("Desktop")
$sourceFoldersExist = 0
$resizePercentage = "85%" # Configurable resize percentage
$LogFile = "$UserHome\ImageConverstionLog.txt"
$heicFiles = $null
$supportedExtensions = @(".heic", ".heif")

# Define paths and requirements
$MagickPath = "$UserHome\ImageConversionScripts\ImageMagick\"
$RequiredBinary = "$MagickPath\magick.exe"
$PublisherLink = "https://download.imagemagick.org/archive/binaries/ImageMagick-7.1.1-47-portable-Q16-HDRI-x64.zip"

# Folder structure
$Folders = @("Images", "ImageConversionScripts")
$SubFolders = @("Images\Converted", "ImageConversionScripts\ImageMagick")
$sourceFolder = "$UserHome\Images"
$destinationFolder = "$UserHome\Images\Converted"
$MagickCmd = "$MagickPath\magick.exe"

# Log action function
function Log-Action {
    param ([string]$Message)
    Add-Content -Path $LogFile -Value "$(Get-Date) - $Message"
}

# Folder existence check function
function Search-FolderExists {
    param (
        [Parameter(Mandatory = $true)] [string]$FolderPath,
        [string]$ShortcutPath,
        [string]$Description
    )

    # Ensure parent folder exists
    $parentFolder = [System.IO.Path]::GetDirectoryName($FolderPath)
    if ($parentFolder -and -not (Test-Path $parentFolder)) {
        New-Item -ItemType Directory -Path $parentFolder -Force | Out-Null
        Log-Action "Parent folder created: $parentFolder"
    }

    # Create folder
    if (-not (Test-Path $FolderPath)) {
        try {
            New-Item -ItemType Directory -Path $FolderPath -Force | Out-Null
            Log-Action "Folder created: $FolderPath"
        } catch {
            Log-Action "Failed to create folder: $FolderPath. Error: $_"
        }
    } else {
        Log-Action "Folder exists: $FolderPath"
    }

    # Create shortcut if needed
    if ($ShortcutPath -and $Description) {
        try {
            $WScriptShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
            $Shortcut.TargetPath = $FolderPath
            $Shortcut.Description = $Description
            $Shortcut.Save()
            Add-Content -Path $LogFile -Value "$(Get-Date) - Shortcut created: $ShortcutPath"
        } catch {
            Log-Action "Failed to create shortcut: $ShortcutPath. Error: $_"
        }
    }
}

# Create folders
foreach ($Folder in $Folders) { 
    $ShortcutPath = "$Desktop\$($Folder.Split('\')[-1]).lnk"
    $Description = "Open $Folder"
    Search-FolderExists -FolderPath "$UserHome\$Folder" -ShortcutPath $ShortcutPath -Description $Description
}

foreach ($SubFolder in $SubFolders) { 
    $ShortcutPath = "$Desktop\$($SubFolder.Split('\')[-1]).lnk"
    $Description = "Open $SubFolder"
    Search-FolderExists -FolderPath "$UserHome\$SubFolder" -ShortcutPath $ShortcutPath -Description $Description
}

# Check for ImageMagick binary
if (-not (Test-Path $RequiredBinary)) {
    Add-Type -AssemblyName System.Windows.Forms
    $missingMessage = @"
    ImageMagick binaries are missing from $MagickPath.
    
    Please follow these steps to resolve the issue:
    1. Download the required files here: $PublisherLink
    2. Extract the contents of the ZIP file.
    3. Copy the extracted files into the following folder:
       $MagickPath
    4. Rerun this script again once the files are in place.
"@
    [System.Windows.Forms.MessageBox]::Show($missingMessage, "Missing Required Files", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    Log-Action "Error: Required binaries missing from $MagickPath."
    exit
} else {
    Log-Action "Required binaries found in $MagickPath."
}

# Get .HEIC files from source folder
$heicFiles = Get-ChildItem -Path $sourceFolder -File | Where-Object { $supportedExtensions -contains $_.Extension.ToLower() }

# If no files found
if ($heicFiles.Count -eq 0) {
    $noFilesMessage = "No HEIC files found in $sourceFolder. Please add HEIC files and try again."
    Log-Action $noFilesMessage
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($noFilesMessage, "No Files Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    exit
} else {
    # Process each HEIC file
    foreach ($file in $heicFiles) {
        $baseName = $file.BaseName
        $destinationFile = Join-Path $destinationFolder "$($baseName).png"
        
        # Try to convert and resize image
        try {
            & "$MagickCmd" "$($file.FullName)" -resize $resizePercentage "$destinationFile"
            Log-Action "Converted $($file.Name) to $destinationFile"
        } catch {
            Log-Action "Failed to convert $($file.Name). Error: $_"
        }
    }

    # Delete original HEIC files
    foreach ($heicFile in $heicFiles) {
        try {
            Remove-Item $heicFile.FullName -Force
            Log-Action "Deleted original file: $($heicFile.FullName)"
        } catch {
            Log-Action "Failed to delete $($heicFile.FullName). Error: $_"
        }
    }

    # Summary popup
    $convertedFiles = Get-ChildItem -Path $destinationFolder -Filter "*.png"
    Add-Type -AssemblyName System.Windows.Forms
    $resultsMessage = "Conversion Summary:`n----------------------------------------"
    foreach ($convertedFile in $convertedFiles) {
        $resultsMessage += "`nConverted: $($convertedFile.Name)"
    }
    [System.Windows.Forms.MessageBox]::Show($resultsMessage, "Conversion Results", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
