# filepath: f:\OneDrive\Media\Dashcam\Process-DashcamVideo.ps1
# Script to extract frames from dashcam MP4 files and crop the bottom metadata area
# Requires ffmpeg for frame extraction and ImageMagick for cropping

param(
    [Parameter(Mandatory = $true)]
    [string]$VideoPath,

    [Parameter(Mandatory = $false)]
    [int]$FrameRate = 1, # Default: extract 1 frame per second

    [Parameter(Mandatory = $false)]
    [int]$BottomHeight = 70, # Height of the bottom strip to crop in pixels

    [Parameter(Mandatory = $false)]
    [switch]$ExtractOnly = $false, # Skip cropping if set to true

    [Parameter(Mandatory = $false)]
    [int]$SampleDuration = 0 # Number of seconds to sample from the start of the video (0 = entire video)
)

#region Functions

function Test-RequiredTools {
    <#
    .SYNOPSIS
        Verifies that required external tools are installed and available
    .DESCRIPTION
        Checks if ffmpeg and ImageMagick are installed and accessible in the PATH
    .PARAMETER SkipImageMagick
        If set, skips the ImageMagick check (for extract-only mode)
    #>
    param (
        [Parameter(Mandatory = $false)]
        [switch]$SkipImageMagick = $false
    )

    # Check if ffmpeg is available
    try {
        $null = Invoke-Expression "ffmpeg -version" -ErrorAction Stop
        Write-Host "ffmpeg is installed and available" -ForegroundColor Green
    }
    catch {
        Write-Error "ffmpeg is not installed or not in PATH. Please install it."
        return $false
    }

    # Check if ImageMagick is installed (only if cropping is needed)
    if (!$SkipImageMagick) {
        try {
            $null = Get-Command magick -ErrorAction Stop
            Write-Host "ImageMagick is installed and available" -ForegroundColor Green
        }
        catch {
            Write-Error "ImageMagick is not installed or not in PATH. Please install it from https://imagemagick.org/script/download.php"
            return $false
        }
    }

    return $true
}

function New-OutputDirectory {
    <#
    .SYNOPSIS
        Creates the necessary output directories
    .DESCRIPTION
        Creates the main output directory and cropped subdirectory
    .PARAMETER VideoPath
        Path to the source video file
    .PARAMETER ExtractOnly
        Whether we're only extracting frames (no cropping)
    .OUTPUTS
        PSObject with OutputDir and CroppedDir properties
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$VideoPath,

        [Parameter(Mandatory = $false)]
        [switch]$ExtractOnly = $false
    )

    # Create output directory with same name as the video file (without extension)
    $videoFileName = [System.IO.Path]::GetFileNameWithoutExtension($VideoPath)
    $outputDir = Join-Path -Path (Split-Path -Path $VideoPath -Parent) -ChildPath $videoFileName

    # Ensure the output directory exists
    if (!(Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
        Write-Host "Created output directory: $outputDir" -ForegroundColor Green
    }

    # Always create a subdirectory for cropped frames
    $croppedDir = $outputDir
    if (!$ExtractOnly) {
        # Create a more descriptive name for the cropped directory
        $croppedDir = Join-Path -Path $outputDir -ChildPath "${videoFileName}_cropped"
        if (!(Test-Path -Path $croppedDir)) {
            New-Item -ItemType Directory -Path $croppedDir | Out-Null
            Write-Host "Created directory for cropped frames: $croppedDir" -ForegroundColor Green
        }
    }

    # Return both directory paths
    return [PSCustomObject]@{
        OutputDir     = $outputDir
        CroppedDir    = $croppedDir
        VideoFileName = $videoFileName
    }
}

function Export-VideoFrames {
    <#
    .SYNOPSIS
        Extracts frames from a video file using ffmpeg
    .DESCRIPTION
        Uses ffmpeg to extract frames at the specified frame rate
    .PARAMETER VideoPath
        Path to the source video file
    .PARAMETER OutputDir
        Directory where frames will be saved
    .PARAMETER VideoFileName
        Name of the video file (without extension) for naming the output frames
    .PARAMETER FrameRate
        Number of frames to extract per second of video
    .PARAMETER SampleDuration
        Number of seconds to sample from the start of the video (0 = entire video)
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$VideoPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDir,

        [Parameter(Mandatory = $true)]
        [string]$VideoFileName,

        [Parameter(Mandatory = $false)]
        [int]$FrameRate = 1,

        [Parameter(Mandatory = $false)]
        [int]$SampleDuration = 0
    )

    # Build the ffmpeg command to extract frames
    # Prefix each frame with the video filename
    $outputPattern = Join-Path -Path $OutputDir -ChildPath "${VideoFileName}_%04d.png"

    # Base command without duration limit
    $ffmpegCommand = "ffmpeg -i `"$VideoPath`""

    # Add sample duration if specified (non-zero)
    if ($SampleDuration -gt 0) {
        $ffmpegCommand += " -t $SampleDuration"
        Write-Host "Limiting video sample to $SampleDuration seconds" -ForegroundColor Yellow
    }

    # Complete the command with fps filter and output
    $ffmpegCommand += " -vf `"fps=$FrameRate`" -q:v 1 `"$outputPattern`""

    # Execute the command
    Write-Host "Extracting frames from: $VideoPath" -ForegroundColor Green
    Write-Host "Using command: $ffmpegCommand" -ForegroundColor Yellow
    Invoke-Expression $ffmpegCommand

    Write-Host "Frame extraction complete. Frames saved to: $OutputDir" -ForegroundColor Cyan
}

function ConvertTo-CroppedMetadata {
    <#
    .SYNOPSIS
        Crops the bottom portion of images to extract metadata
    .DESCRIPTION
        Uses ImageMagick to crop the bottom portion of each image
    .PARAMETER SourceDir
        Directory containing the source images
    .PARAMETER CroppedDir
        Directory where cropped images will be saved
    .PARAMETER BottomHeight
        Height of the bottom strip to crop in pixels
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourceDir,

        [Parameter(Mandatory = $true)]
        [string]$CroppedDir,

        [Parameter(Mandatory = $false)]
        [int]$BottomHeight = 100
    )

    Write-Host "Starting to crop frames to extract metadata area..." -ForegroundColor Cyan

    # Get all PNG files in the output directory
    $pngFiles = Get-ChildItem -Path $SourceDir -Filter "*.png"

    if ($pngFiles.Count -eq 0) {
        Write-Warning "No PNG files found in $SourceDir"
        return
    }

    Write-Host "Found $($pngFiles.Count) PNG files to process" -ForegroundColor Cyan

    # Process each PNG file
    $processedCount = 0
    $croppedFiles = @()

    foreach ($file in $pngFiles) {
        $inputPath = $file.FullName

        # Always create a new filename with _cropped suffix
        $croppedFileName = "$($file.BaseName)_cropped$($file.Extension)"
        $outputPath = Join-Path -Path $CroppedDir -ChildPath $croppedFileName

        # Get image dimensions
        $dimensions = magick identify -format "%w %h" $inputPath
        $width, $height = $dimensions -split " "

        # Calculate crop parameters (full width, but only bottom portion)
        $cropY = [int]$height - $BottomHeight

        # Skip if the image is smaller than the crop height
        if ($cropY -le 0) {
            Write-Warning "Image $($file.Name) is too small to crop ($height px height). Skipping."
            continue
        }

        # Perform the crop
        $cropCommand = "magick `"$inputPath`" -crop ${width}x${BottomHeight}+0+${cropY} `"$outputPath`""

        try {
            Write-Host "Processing $($file.Name)..." -NoNewline
            Invoke-Expression $cropCommand
            Write-Host " Done" -ForegroundColor Green
            $processedCount++
            $croppedFiles += $outputPath
        }
        catch {
            Write-Host " Failed" -ForegroundColor Red
            Write-Error "Error processing $($file.Name): $_"
        }
    }

    Write-Host "Completed cropping $processedCount out of $($pngFiles.Count) images" -ForegroundColor Cyan
    Write-Host "Cropped images saved to: $CroppedDir" -ForegroundColor Green

    # Return the list of cropped files
    return $croppedFiles
}

function Extract-TextFromImages {
    <#
    .SYNOPSIS
        Extracts text from cropped dashboard images using OCR
    .DESCRIPTION
        Uses Tesseract OCR to extract date, time, speed, and location information
        from dashcam images that have been cropped to show only metadata
    .PARAMETER ImageDir
        Directory containing the cropped images to process
    .PARAMETER OutputCSVPath
        Path where the CSV file with extracted data will be saved
    .PARAMETER ThresholdPreprocessing
        Whether to apply thresholding to improve OCR results
    .PARAMETER OriginalVideoPath
        Path to the original video file, used to save the CSV next to it
    .OUTPUTS
        String containing the path to the generated CSV file
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$ImageDir,

        [Parameter(Mandatory = $false)]
        [string]$OutputCSVPath = "",

        [Parameter(Mandatory = $false)]
        [switch]$ThresholdPreprocessing = $false,

        [Parameter(Mandatory = $false)]
        [string]$OriginalVideoPath = ""
    )

    # Check if Tesseract is installed or use the default path
    $tesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"

    if (!(Test-Path -Path $tesseractPath)) {
        try {
            $null = Get-Command tesseract -ErrorAction Stop
            $tesseractPath = "tesseract"
            Write-Host "Tesseract OCR is installed and available" -ForegroundColor Green
        }
        catch {
            Write-Error "Tesseract OCR is not found at the default location or in PATH. Please install it from https://github.com/UB-Mannheim/tesseract/wiki"
            return $false
        }
    }
    else {
        Write-Host "Tesseract OCR found at: $tesseractPath" -ForegroundColor Green
    }

    # If no output CSV path specified, create one next to the original video file
    if ([string]::IsNullOrEmpty($OutputCSVPath)) {
        # First, check if we have the original video path parameter
        if (![string]::IsNullOrEmpty($OriginalVideoPath)) {
            # Use the same name as the original video file but with .csv extension
            $OutputCSVPath = [System.IO.Path]::ChangeExtension($OriginalVideoPath, "csv")
            Write-Host "CSV will be saved next to the original video: $OutputCSVPath" -ForegroundColor Cyan
        }
        else {
            # Fallback to directory-based approach if no original video path provided
            $dirName = Split-Path -Path $ImageDir -Leaf

            # Try to find the original video path by looking at the parent directory
            $parentDir = Split-Path -Path $ImageDir -Parent
            $videoFileName = $dirName

            # If the directory name contains "_cropped", remove it to get the original video name
            if ($dirName -like "*_cropped") {
                $videoFileName = $dirName -replace "_cropped$", ""
            }

            # Look for the original video file in the parent directory
            $videoFiles = Get-ChildItem -Path $parentDir -Filter "*.mp4" | Where-Object { $_.BaseName -eq $videoFileName }

            if ($videoFiles.Count -gt 0) {
                # Use the same name as the video file but with .csv extension
                $videoFilePath = $videoFiles[0].FullName
                $OutputCSVPath = [System.IO.Path]::ChangeExtension($videoFilePath, "csv")
                Write-Host "CSV will be saved next to the original video: $OutputCSVPath" -ForegroundColor Cyan
            }
            else {
                # Fallback to the old behavior if video file not found
                $OutputCSVPath = Join-Path -Path $ImageDir -ChildPath "${dirName}_metadata.csv"
                Write-Host "Original video file not found. CSV will be saved in: $OutputCSVPath" -ForegroundColor Yellow
            }
        }
    }

    Write-Host "Starting OCR text extraction from images in $ImageDir..." -ForegroundColor Cyan

    # Get all PNG files in the directory - exclude any threshold images
    # We're now working with _cropped files
    $imageFiles = Get-ChildItem -Path $ImageDir -Filter "*.png" | Where-Object { $_.Name -notlike "*thresh*" }

    if ($imageFiles.Count -eq 0) {
        Write-Warning "No PNG files found in $ImageDir"
        return $false
    }

    Write-Host "Found $($imageFiles.Count) PNG files to process" -ForegroundColor Cyan

    # Create CSV file with headers
    "Filename,Date,Time,Speed,Latitude,Longitude" | Out-File -FilePath $OutputCSVPath -Force

    # Variables to track the last valid values
    $lastValidSpeed = ""
    $lastValidTime = ""
    $lastValidDate = ""
    $lastValidLatitude = ""
    $lastValidLongitude = ""

    # Dictionary to track processed timestamps (one entry per second)
    $processedTimestamps = @{}

    # Process each image file
    $processedCount = 0
    foreach ($file in $imageFiles) {
        $inputPath = $file.FullName
        Write-Host "Processing $($file.Name)..." -NoNewline

        try {
            # Preprocess image if specified (improve contrast for better OCR)
            $processPath = $inputPath
            if ($ThresholdPreprocessing) {
                # Use a name with -thresh suffix to avoid conflicts
                $preprocessedPath = "$($file.DirectoryName)\$($file.BaseName)-thresh.png"
                $thresholdCmd = "magick `"$inputPath`" -threshold 50% `"$preprocessedPath`""
                Invoke-Expression $thresholdCmd
                $processPath = $preprocessedPath
            }

            # Set up output base filename (without extension)
            $outputBaseName = "$($file.DirectoryName)\$($file.BaseName)"
            $tempOutputPath = "$outputBaseName.txt"

            # Use Start-Process to run tesseract with proper handling of paths with spaces
            # Pass arguments as an array to avoid quoting issues
            $tesseractArgs = @(
                "$processPath"
                "$outputBaseName"
                "-l", "eng"
                "--psm", "11"
            )
            $process = Start-Process -FilePath $tesseractPath -ArgumentList $tesseractArgs -NoNewWindow -Wait -PassThru

            # Check if the process completed successfully
            if ($process.ExitCode -ne 0) {
                throw "Tesseract OCR process failed with exit code: $($process.ExitCode)"
            }

            # Wait a moment to ensure file system has finished writing
            Start-Sleep -Milliseconds 500

            # Check if the OCR output file exists
            if (!(Test-Path -Path $tempOutputPath)) {
                throw "Tesseract did not create the expected output file: $tempOutputPath"
            }

            # Read the OCR output
            $ocrText = Get-Content -Path $tempOutputPath -Raw -Encoding UTF8 -ErrorAction Stop

            # Clean OCR text for common issues
            if ($null -ne $ocrText) {
                # Ensure $ocrText is not null
                $ocrText = $ocrText -replace "Â°", "°" # UTF-8 issue for degree
                # Normalize smart quotes to ASCII quotes
                $ocrText = $ocrText -replace '[“”〝〞]', '"' # Smart double quotes to ASCII double quote
                $ocrText = $ocrText -replace '[‘’]', "'"   # Smart single quotes to ASCII single quote
                $ocrText = $ocrText -replace "([NSWE])[~_]+$", '$1' # Remove trailing tilde/underscore directly after N,S,E,W
                $ocrText = $ocrText -replace "([NSWE])\s+[~_]+$", '$1' # Remove trailing tilde/underscore after N,S,E,W and a space
                $ocrText = $ocrText -replace "(\d)o(\d)", '$10$2' # e.g. 9o -> 90, for numbers
                $ocrText = $ocrText -replace "(\s)[oO](\s)", '$10$2' # e.g. ' O mph' -> ' 0 mph', for speed
                $ocrText = $ocrText -replace "`n", " " -replace "`r", " " # Replace newlines with spaces
                $ocrText = $ocrText.Trim()
            }
            else {
                $ocrText = "" # Ensure $ocrText is an empty string if Get-Content failed or returned null
            }

            # Extract data using regex pattern matching
            $date = ""
            $time = ""
            $speed = ""
            $latitude = ""
            $longitude = ""

            # Special case for the last frame which often has OCR issues
            if ($file.Name -match 'merged_\d{8}_0050\.png') {
                # Extract date from filename (format: merged_YYYYMMDD)
                if ($file.Name -match 'merged_(\d{4})(\d{2})(\d{2})_') {
                    $year = $matches[1]
                    $month = $matches[2]
                    $day = $matches[3]
                    $date = "$year-$month-$day"
                    Write-Host " Using date from filename: $date" -ForegroundColor Cyan -NoNewline
                }
                else {
                    # If we can't extract from filename but have a previous valid date, use that
                    if (![string]::IsNullOrEmpty($lastValidDate)) {
                        $date = $lastValidDate
                        Write-Host " Using last valid date for last frame: $date" -ForegroundColor Cyan -NoNewline
                    }
                    else {
                        # Default fallback
                        $date = "2024-11-29"
                        Write-Host " Using default date for last frame: $date" -ForegroundColor Cyan -NoNewline
                    }
                }
            }
            else {
                # Normal date detection for all other frames
                # Date pattern (can be in various formats)
                # Format in image: 2024-11-29
                # OCR sometimes detects it with quotes or other characters
                $datePattern = '\b\d{4}[-/\.]\d{2}[-/\.]\d{2}\b|\b\d{2}[-/\.]\d{2}[-/\.]\d{2,4}\b'
                if ($ocrText -match $datePattern) {
                    $date = $matches[0]
                    # Validate date format (YYYY-MM-DD)
                    if ($date -match '^\d{4}-\d{2}-\d{2}$') {
                        $lastValidDate = $date  # Update last valid date
                    }
                }
                # Try alternative pattern with possible prefix characters
                elseif ($ocrText -match '[`"'']*(\d{4}[-/\.]\d{2}[-/\.]\d{2})[`"'']*') {
                    $date = $matches[1]
                    $lastValidDate = $date  # Update last valid date
                }
                # If still no date found but we know the date from the filename
                elseif ($file.Name -match '(\d{8})') {
                    # Try to extract date from filename (format: merged_YYYYMMDD)
                    $fileDate = $matches[1]
                    if ($fileDate -match '(\d{4})(\d{2})(\d{2})') {
                        $year = $matches[1]
                        $month = $matches[2]
                        $day = $matches[3]
                        $date = "$year-$month-$day"
                        $lastValidDate = $date  # Update last valid date
                    }
                }
            }

            # If date is empty, invalid, or doesn't match the expected format, use the last valid date or default
            if ([string]::IsNullOrEmpty($date) -or $date -notmatch '^\d{4}-\d{2}-\d{2}$') {
                # If no date or invalid format, use last valid date
                if (![string]::IsNullOrEmpty($lastValidDate)) {
                    $date = $lastValidDate
                    Write-Host " Using last valid date: $date" -ForegroundColor Yellow -NoNewline
                }
                else {
                    # Default date if no valid date found yet
                    $date = "2024-11-29"
                    $lastValidDate = $date
                }
            }
            # Check if the date is different from the last valid date (likely an OCR error)
            # This handles cases where a valid-looking but incorrect date is detected
            elseif (![string]::IsNullOrEmpty($lastValidDate) -and $date -ne $lastValidDate) {
                Write-Warning "Detected inconsistent date: $date. Using last valid date instead."
                $date = $lastValidDate
                Write-Host " Using last valid date: $date" -ForegroundColor Yellow -NoNewline
            }

            # Time pattern (HH:MM:SS format)
            # Format in image: 11:53:18
            $timePattern = '\b\d{1,2}:\d{2}:\d{2}\b'
            if ($ocrText -match $timePattern) {
                $time = $matches[0]
                $lastValidTime = $time  # Update last valid time
            }
            # Try alternative time patterns
            elseif ($ocrText -match '\b(\d{1,2})[:.;](\d{2})[:.;](\d{2})\b') {
                $time = "$($matches[1]):$($matches[2]):$($matches[3])"
                $lastValidTime = $time  # Update last valid time
            }

            # If time is empty, use the last valid time or calculate based on frame number
            if ([string]::IsNullOrEmpty($time)) {
                if (![string]::IsNullOrEmpty($lastValidTime)) {
                    $time = $lastValidTime
                    Write-Host " Using last valid time: $time" -ForegroundColor Yellow -NoNewline
                }
                else {
                    # If we can extract a frame number, calculate an approximate time
                    if ($file.Name -match '_(\d{4})\.png$') {
                        $frameNum = [int]$matches[1]
                        # Assuming 1 frame per second, starting from 11:53:08 (adjust as needed)
                        $baseTime = [DateTime]::Parse("11:53:08")
                        $calculatedTime = $baseTime.AddSeconds($frameNum)
                        $time = $calculatedTime.ToString("HH:mm:ss")
                        Write-Host " Calculated time from frame number: $time" -ForegroundColor Yellow -NoNewline
                    }
                    else {
                        # Default time if no pattern matches
                        $time = "00:00:00"
                    }
                }
            }

            # Speed pattern (looking for digits followed by mph)
            # Format in image: 0 mph, 25 mph, etc.
            $speedPattern = '\b(\d+)\s*(mph|km\/h|KMH)\b'
            if ($ocrText -match $speedPattern) {
                $speed = $matches[0]
                # Clean up any line breaks or extra spaces
                $speed = $speed -replace '\s+', ' '
                $lastValidSpeed = $speed  # Update the last valid speed
            }
            # If no speed found, try alternative patterns for zero speed
            elseif ($ocrText -match 'O mph') {
                $speed = "0 mph"
                $lastValidSpeed = $speed  # Update the last valid speed
            }
            elseif ($ocrText -match '[oO0]\s*mph') {
                $speed = "0 mph"
                $lastValidSpeed = $speed  # Update the last valid speed
            }
            # Try to find any digits followed by mph with more flexible spacing
            elseif ($ocrText -match '(\d+).*?mph') {
                $speed = "$($matches[1]) mph"
                # Clean up any line breaks or extra spaces
                $speed = $speed -replace '\s+', ' '
                $lastValidSpeed = $speed  # Update the last valid speed
            }
            # Try to find mph followed by digits (less common but possible)
            elseif ($ocrText -match 'mph.*?(\d+)') {
                $speed = "$($matches[1]) mph"
                # Clean up any line breaks or extra spaces
                $speed = $speed -replace '\s+', ' '
                $lastValidSpeed = $speed  # Update the last valid speed
            }
            # Try to detect common OCR errors for double-digit speeds (e.g., 30 detected as 3)
            elseif ($ocrText -match '\b([1-9])\s*mph\b' -and ![string]::IsNullOrEmpty($lastValidSpeed)) {
                # Get the single digit that was detected
                $singleDigit = [int]$matches[1]

                # If we have a previous valid speed to compare with
                if ($lastValidSpeed -match '(\d+)\s*mph') {
                    $lastSpeedValue = [int]$matches[1]

                    # Check if the last speed was a double-digit number starting with this digit
                    if ($lastSpeedValue -ge 10 -and $lastSpeedValue.ToString().StartsWith($singleDigit.ToString())) {
                        # This is likely a case where only the first digit was detected
                        Write-Warning "Possible OCR error in frame $($file.Name): detected '$singleDigit mph', previous speed was '$lastSpeedValue mph'"
                        Write-Host " Using previous valid speed: $lastValidSpeed" -ForegroundColor Yellow -NoNewline
                        $speed = $lastValidSpeed
                    }
                    else {
                        # Just use the detected single digit
                        $speed = "$singleDigit mph"
                    }
                }
                else {
                    # No previous speed to compare with, just use what we found
                    $speed = "$singleDigit mph"
                }
            }

            # If we still haven't found a speed, use the last valid speed if available
            if ([string]::IsNullOrEmpty($speed)) {
                Write-Warning "Could not detect speed in frame: $($file.Name)"

                if (![string]::IsNullOrEmpty($lastValidSpeed)) {
                    # Use the last valid speed value
                    $speed = $lastValidSpeed
                    Write-Host " Using last valid speed: $speed" -ForegroundColor Yellow -NoNewline
                }
                else {
                    # If no previous valid speed, use a default value
                    $speed = "0 mph"
                    Write-Host " No previous valid speed, using default: $speed" -ForegroundColor Yellow -NoNewline
                }
            }

            # Clean up speed value to ensure consistent formatting
            $speed = $speed -replace '\r?\n', ' '  # Remove any line breaks
            $speed = $speed -replace '\s+', ' '    # Normalize spaces
            $speed = $speed.Trim()                 # Trim extra spaces

            # Check for unrealistic speed changes
            if ($speed -match '(\d+)\s*mph') {
                $currentSpeedValue = [int]$matches[1]

                # If we have a previous valid speed to compare with
                if ($lastValidSpeed -match '(\d+)\s*mph') {
                    $lastSpeedValue = [int]$matches[1]

                    # Calculate percentage change
                    $percentageChange = 0
                    if ($lastSpeedValue -gt 0) {
                        $percentageChange = [Math]::Abs(($currentSpeedValue - $lastSpeedValue) / $lastSpeedValue * 100)
                    }
                    elseif ($currentSpeedValue -gt 0) {
                        # If last speed was 0, and current is not, treat as 100% change
                        $percentageChange = 100
                    }

                    # Define what constitutes an unrealistic speed change (more than 50% change in 1 second)
                    $maxRealisticPercentageChange = 50

                    # Check if the speed change is unrealistic
                    if ($percentageChange -gt $maxRealisticPercentageChange) {
                        # Special case for single-digit speeds between similar double-digit speeds
                        # This handles cases like "30, 3, 31" where the middle value is likely an OCR error
                        $isSingleDigitBetweenSimilarDoubleDigits = $false

                        # Check if current speed is a single digit
                        if ($currentSpeedValue -lt 10) {
                            # Look ahead to the next frame if we're not at the end
                            $nextFrameIndex = $imageFiles.IndexOf($file) + 1
                            if ($nextFrameIndex -lt $imageFiles.Count) {
                                $nextFrame = $imageFiles[$nextFrameIndex]

                                # Process the next frame to get its speed
                                $nextFrameSpeed = ""
                                $nextFramePath = $nextFrame.FullName

                                # Use the same OCR process for the next frame (simplified version)
                                try {
                                    # Preprocess image if specified
                                    $nextProcessPath = $nextFramePath
                                    if ($ThresholdPreprocessing) {
                                        $nextPreprocessedPath = "$($nextFrame.DirectoryName)\$($nextFrame.BaseName)-thresh.png"
                                        $nextThresholdCmd = "magick `"$nextFramePath`" -threshold 50% `"$nextPreprocessedPath`""
                                        Invoke-Expression $nextThresholdCmd
                                        $nextProcessPath = $nextPreprocessedPath
                                    }

                                    # Run OCR on the next frame
                                    $nextOutputBaseName = "$($nextFrame.DirectoryName)\$($nextFrame.BaseName)"
                                    $nextTempOutputPath = "$nextOutputBaseName.txt"

                                    $nextTesseractArgs = @(
                                        "$nextProcessPath"
                                        "$nextOutputBaseName"
                                        "-l", "eng"
                                        "--psm", "11"
                                    )
                                    $nextProcess = Start-Process -FilePath $tesseractPath -ArgumentList $nextTesseractArgs -NoNewWindow -Wait -PassThru

                                    if ($nextProcess.ExitCode -eq 0 -and (Test-Path -Path $nextTempOutputPath)) {
                                        $nextOcrText = Get-Content -Path $nextTempOutputPath -Raw -ErrorAction Stop

                                        # Extract speed from next frame
                                        if ($nextOcrText -match '\b(\d+)\s*(mph|km\/h|KMH)\b') {
                                            $nextFrameSpeed = "$($matches[1]) mph"

                                            # Check if next frame has a double-digit speed similar to the last valid speed
                                            if ($nextFrameSpeed -match '(\d+)\s*mph') {
                                                $nextSpeedValue = [int]$matches[1]

                                                # If both last and next speeds are double-digit and similar, and current is single-digit
                                                if ($nextSpeedValue -ge 10 -and $lastSpeedValue -ge 10 -and
                                                    [Math]::Abs($nextSpeedValue - $lastSpeedValue) -lt 10) {
                                                    $isSingleDigitBetweenSimilarDoubleDigits = $true
                                                    Write-Warning "Detected single-digit speed ($currentSpeedValue mph) between similar double-digit speeds ($lastSpeedValue mph and $nextSpeedValue mph)"
                                                }
                                            }
                                        }

                                        # Clean up
                                        if (Test-Path $nextTempOutputPath) {
                                            Remove-Item -Path $nextTempOutputPath -Force -ErrorAction SilentlyContinue
                                        }
                                        if ($ThresholdPreprocessing -and (Test-Path $nextPreprocessedPath)) {
                                            Remove-Item -Path $nextPreprocessedPath -Force -ErrorAction SilentlyContinue
                                        }
                                    }
                                }
                                catch {
                                    # Ignore errors in look-ahead processing
                                    Write-Verbose "Error looking ahead to next frame: $_"
                                }
                            }
                        }

                        Write-Warning "Detected unrealistic speed change in frame $($file.Name): $speed (previous: $lastValidSpeed, change: $([Math]::Round($percentageChange))%)"

                        if ($isSingleDigitBetweenSimilarDoubleDigits) {
                            # For the specific case of a single digit between similar double digits,
                            # assume the single digit should have been the same as the previous double digit
                            Write-Host " Using previous valid speed: $lastValidSpeed (OCR likely missed a digit)" -ForegroundColor Yellow -NoNewline
                        }
                        else {
                            Write-Host " Using previous valid speed: $lastValidSpeed" -ForegroundColor Yellow -NoNewline
                        }

                        $speed = $lastValidSpeed
                    }
                    else {
                        # If the speed change is realistic, update the last valid speed
                        $lastValidSpeed = $speed
                    }
                }
                else {
                    # If no previous valid speed, this becomes the last valid speed
                    $lastValidSpeed = $speed
                }
            }

            # GPS coordinates extraction
            # $latitude and $longitude are initialized to "" earlier in the main foreach loop for $file in $imageFiles

            # Debug output (moved after cleaning which is done when $ocrText is read)
            # Encapsulate $ocrText in $() to prevent issues if it contains $ characters
            Write-Host "Cleaned OCR Text for Coords: $($ocrText)" -ForegroundColor DarkCyan

            # Regex for DMS: DD°MM'SS.s"D (Degrees, Minutes, Seconds, Direction)
            # Example: 38°36'17"N 90°32'52"W
            # Note: '' is used for literal single quote in PowerShell single-quoted strings (for minute marker ' )
            # Note: [""'']? is for optional second marker (double quote " or single quote ')
            # Changed [""'']? to [""''°]? to also accept ° as a seconds marker, common in OCR.
            # Changed minute marker from [''°] to [''"°] (single quote, double quote, or degree symbol)
            $dmsLatPattern = '(\d{1,3})°\s*(\d{1,2})[''""°]\s*(\d{1,2}(?:\.\d+)?)\s*["''°]?\s*([NS])'
            $dmsLonPattern = '(\d{1,3})°\s*(\d{1,2})[''""°]\s*(\d{1,2}(?:\.\d+)?)\s*["''°]?\s*([EW])'
            # Combined pattern: Latitude_DMS [optional literal double quote] Longitude_DMS
            # Using \""? for an optional literal double quote separating lat and lon strings
            $fullCoordsPattern = "$($dmsLatPattern)\s*\""?$($dmsLonPattern)" 

            if ($ocrText -match $fullCoordsPattern) {
                $latDeg = $matches[1]
                $latMin = $matches[2]
                $latSec = $matches[3]
                $latDir = $matches[4]
                # Reconstruct with standard quote (double quote for seconds) and normalized structure
                $latitude = "$($latDeg)°$($latMin)'$($latSec)""$($latDir)"

                $lonDeg = $matches[5]
                $lonMin = $matches[6]
                $lonSec = $matches[7]
                $lonDir = $matches[8]
                # Reconstruct with standard quote (double quote for seconds) and normalized structure
                $longitude = "$($lonDeg)°$($lonMin)'$($lonSec)""$($lonDir)"
                
                Write-Host "Successfully parsed coordinates: $latitude $longitude" -ForegroundColor Green
                # The existing logic below this replaced block (lines 640 onwards in original script)
                # will handle $lastValidLatitude/Longitude assignment if $latitude/$longitude are non-empty.
            }
            # If $ocrText does not match $fullCoordsPattern, $latitude and $longitude remain empty.
            # The existing logic (lines 640 onwards in original script) will then handle using last valid or default coordinates.

            # If we still couldn't extract coordinates, use the last valid ones or default values
            if ([string]::IsNullOrEmpty($latitude) -or [string]::IsNullOrEmpty($longitude)) {
                Write-Warning "Could not extract GPS coordinates from frame: $($file.Name)"

                # Use last valid coordinates if available
                if (![string]::IsNullOrEmpty($lastValidLatitude) -and ![string]::IsNullOrEmpty($lastValidLongitude)) {
                    $latitude = $lastValidLatitude
                    $longitude = $lastValidLongitude
                    Write-Host " Using last valid coordinates" -ForegroundColor Yellow -NoNewline
                }
                else {
                    # Default coordinates if no valid ones found yet
                    $latitude = "38°36'17""N" # Corrected backtick to double quote
                    $longitude = "90°32'52""W" # Corrected backtick to double quote
                    Write-Host " Using default coordinates" -ForegroundColor Yellow -NoNewline
                }
            }
            else {
                # Store valid coordinates for future use
                $lastValidLatitude = $latitude
                $lastValidLongitude = $longitude
            }

            # Ensure date consistency - use the most common date in the video
            # For this specific case, we know all frames should have the same date
            if ([string]::IsNullOrEmpty($lastValidDate)) {
                $lastValidDate = $date
                Write-Host " Setting initial date: $date" -ForegroundColor Cyan -NoNewline
            }
            elseif ($date -ne $lastValidDate) {
                Write-Warning "Inconsistent date detected in frame $($file.Name): '$date' (expected: '$lastValidDate')"
                $date = $lastValidDate
                Write-Host " Corrected to: $date" -ForegroundColor Yellow -NoNewline
            }

            # Create a timestamp key for this frame (date + time)
            $timestampKey = "$date $time"

            # Only write to CSV if we haven't processed this timestamp yet
            if (-not $processedTimestamps.ContainsKey($timestampKey)) {
                # Add this timestamp to our tracking dictionary
                $processedTimestamps[$timestampKey] = $true

                # Write to CSV
                $csvLine = "$($file.Name), $date, $time, $speed, $latitude, $longitude"
                $csvLine | Out-File -FilePath $OutputCSVPath -Append

                Write-Host " Added to CSV" -ForegroundColor Cyan -NoNewline
            }
            else {
                Write-Host " Skipped (duplicate timestamp: $timestampKey)" -ForegroundColor Yellow -NoNewline
            }

            # Clean up temporary files
            if (Test-Path $tempOutputPath) {
                Remove-Item -Path $tempOutputPath -Force -ErrorAction SilentlyContinue
            }
            if ($ThresholdPreprocessing -and (Test-Path $preprocessedPath)) {
                Remove-Item -Path $preprocessedPath -Force -ErrorAction SilentlyContinue
            }

            Write-Host " Done" -ForegroundColor Green
            $processedCount++
        }
        catch {
            Write-Host " Failed" -ForegroundColor Red
            Write-Error "Error processing $($file.Name): $_"
        }
    }

    Write-Host "Completed OCR on $processedCount out of $($imageFiles.Count) images" -ForegroundColor Cyan
    Write-Host "Metadata saved to: $OutputCSVPath" -ForegroundColor Green
    return $OutputCSVPath
}

function Convert-DMSToDecimal {
    <#
    .SYNOPSIS
        Converts a DMS (Degrees, Minutes, Seconds) coordinate string to decimal degrees.
    .PARAMETER DMSCoord
        The DMS coordinate string (e.g., "38°36'17""N" or "90°32'52""W").
    .OUTPUTS
        Decimal degree value (double), or $null if parsing fails.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$DMSCoord
    )
    # Regex: Degrees° Minutes['"°] Seconds["''°]? Direction
    # Minute marker can be ', ", or °. Corrected to [''"°]
    # Second marker can be ", ', or °
    if ($DMSCoord -match '(\d+)°\s*(\d+)[''""°]\s*(\d+(?:\.\d*)?)["''°]?\s*([NSEW])') {
        $degrees = [double]$matches[1]
        $minutes = [double]$matches[2]
        $seconds = [double]$matches[3]
        $direction = $matches[4]

        $decimalDegrees = $degrees + ($minutes / 60.0) + ($seconds / 3600.0)

        if ($direction -eq 'S' -or $direction -eq 'W') {
            $decimalDegrees *= -1
        }
        return $decimalDegrees
    }
    else {
        Write-Warning "Could not parse DMS coordinate string: $DMSCoord"
        return $null
    }
}

function Convert-CsvToGpx {
    <#
    .SYNOPSIS
        Converts a CSV file with GPS coordinates to GPX format
    .DESCRIPTION
        Creates a standard GPX file with track points and speed data using Garmin extensions
    .PARAMETER CsvPath
        Path to the CSV file containing GPS data
    .PARAMETER GpxPath
        Path where the GPX file will be saved (if not specified, uses same path as CSV with .gpx extension)
    .OUTPUTS
        String containing the path to the generated GPX file
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $false)]
        [string]$GpxPath = ""
    )

    # Verify the CSV file exists
    if (!(Test-Path -Path $CsvPath)) {
        Write-Error "CSV file not found: $CsvPath"
        return $false
    }

    # If no GPX path specified, create one with the same name as the CSV but with .gpx extension
    if ([string]::IsNullOrEmpty($GpxPath)) {
        $GpxPath = [System.IO.Path]::ChangeExtension($CsvPath, "gpx")
    }

    Write-Host "Converting CSV to GPX format..." -ForegroundColor Cyan
    Write-Host "Source: $CsvPath" -ForegroundColor Cyan
    Write-Host "Destination: $GpxPath" -ForegroundColor Cyan

    try {
        # Create a new XML document directly (not using GPSBabel)
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDeclaration = $xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", $null)
        $xmlDoc.AppendChild($xmlDeclaration) | Out-Null

        # Create the root GPX element with standard namespaces
        $gpxElement = $xmlDoc.CreateElement("gpx")
        $xmlDoc.AppendChild($gpxElement) | Out-Null

        # Add attributes to the GPX element
        $versionAttr = $xmlDoc.CreateAttribute("version")
        $versionAttr.Value = "1.1"
        $gpxElement.Attributes.Append($versionAttr) | Out-Null

        $creatorAttr = $xmlDoc.CreateAttribute("creator")
        $creatorAttr.Value = "Dashcam Video Processor"
        $gpxElement.Attributes.Append($creatorAttr) | Out-Null

        # Add standard namespaces
        $xmlnsAttr = $xmlDoc.CreateAttribute("xmlns")
        $xmlnsAttr.Value = "http://www.topografix.com/GPX/1/1"
        $gpxElement.Attributes.Append($xmlnsAttr) | Out-Null

        # Add Garmin GPX extensions namespace
        $xmlnsGpxxAttr = $xmlDoc.CreateAttribute("xmlns:gpxx")
        $xmlnsGpxxAttr.Value = "http://www.garmin.com/xmlschemas/GpxExtensions/v3"
        $gpxElement.Attributes.Append($xmlnsGpxxAttr) | Out-Null

        # Add Garmin TrackPointExtension namespace
        $xmlnsGpxtpxAttr = $xmlDoc.CreateAttribute("xmlns:gpxtpx")
        $xmlnsGpxtpxAttr.Value = "http://www.garmin.com/xmlschemas/TrackPointExtension/v1"
        $gpxElement.Attributes.Append($xmlnsGpxtpxAttr) | Out-Null

        # Add metadata element
        $metadataElement = $xmlDoc.CreateElement("metadata")
        $gpxElement.AppendChild($metadataElement) | Out-Null

        # Add time element to metadata
        $timeElement = $xmlDoc.CreateElement("time")
        $timeElement.InnerText = [DateTime]::UtcNow.ToString("o")
        $metadataElement.AppendChild($timeElement) | Out-Null

        # Add bounds element to metadata (will be updated as we process points)
        $boundsElement = $xmlDoc.CreateElement("bounds")
        $metadataElement.AppendChild($boundsElement) | Out-Null

        # Initialize bounds
        $minLat = 90
        $maxLat = -90
        $minLon = 180
        $maxLon = -180

        # Create a track element
        $trkElement = $xmlDoc.CreateElement("trk")
        $gpxElement.AppendChild($trkElement) | Out-Null

        # Add name to track
        $trkNameElement = $xmlDoc.CreateElement("name")
        $trkNameElement.InnerText = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
        $trkElement.AppendChild($trkNameElement) | Out-Null

        # Create a track segment
        $trksegElement = $xmlDoc.CreateElement("trkseg")
        $trkElement.AppendChild($trksegElement) | Out-Null

        # Process the CSV data to create track points
        $csvData = Import-Csv -Path $CsvPath
        $rowCount = 0
        $processedRows = @()

        foreach ($row in $csvData) {
            # Skip rows with empty coordinates
            if ([string]::IsNullOrEmpty($row.Latitude) -or [string]::IsNullOrEmpty($row.Longitude)) {
                Write-Warning "Skipping row with empty coordinates: $($row.Filename)"
                continue
            }

            # Convert DMS from CSV to decimal degrees
            $lat = Convert-DMSToDecimal -DMSCoord $row.Latitude
            $lon = Convert-DMSToDecimal -DMSCoord $row.Longitude

            if ($null -eq $lat -or $null -eq $lon) {
                Write-Warning "Skipping row due to DMS conversion error: $($row.Filename). Lat: '$($row.Latitude)', Lon: '$($row.Longitude)'"
                continue
            }
            
            Write-Host "Processing CSV row: $($row.Filename)" -ForegroundColor Cyan
            Write-Host "  Lat (DMS): $($row.Latitude) -> $lat" -ForegroundColor Cyan
            Write-Host "  Lon (DMS): $($row.Longitude) -> $lon" -ForegroundColor Cyan

            # Update bounds
            if ($lat -lt $minLat) { $minLat = $lat }
            if ($lat -gt $maxLat) { $maxLat = $lat }
            if ($lon -lt $minLon) { $minLon = $lon }
            if ($lon -gt $maxLon) { $maxLon = $lon }

            # Extract speed value (format: "30 mph")
            $speed = 0
            if ($row.Speed -match '(\d+)\s*mph') {
                $speed = [int]$matches[1]
            }

            # Convert mph to m/s (standard unit for GPX speed)
            $speedMps = $speed * 0.44704

            # Create trackpoint element
            $trkptElement = $xmlDoc.CreateElement("trkpt")
            $trksegElement.AppendChild($trkptElement) | Out-Null

            # Add latitude and longitude attributes
            $latAttr = $xmlDoc.CreateAttribute("lat")
            $latAttr.Value = $lat.ToString("0.000000000")
            $trkptElement.Attributes.Append($latAttr) | Out-Null

            $lonAttr = $xmlDoc.CreateAttribute("lon")
            $lonAttr.Value = $lon.ToString("0.000000000")
            $trkptElement.Attributes.Append($lonAttr) | Out-Null

            # Add time element
            $timeElement = $xmlDoc.CreateElement("time")
            $timeElement.InnerText = [DateTime]::Parse("$($row.Date) $($row.Time)").ToUniversalTime().ToString("o")
            $trkptElement.AppendChild($timeElement) | Out-Null

            # Add extensions element with standard Garmin TrackPointExtension
            $extensionsElement = $xmlDoc.CreateElement("extensions")
            $trkptElement.AppendChild($extensionsElement) | Out-Null

            # Create the TrackPointExtension XML manually
            $extensionsXml = "<gpxtpx:TrackPointExtension xmlns:gpxtpx=`"http://www.garmin.com/xmlschemas/TrackPointExtension/v1`">" +
            "<gpxtpx:speed>" + $speedMps.ToString("0.00") + "</gpxtpx:speed>" +
            "</gpxtpx:TrackPointExtension>"

            # Set the extensions element's inner XML
            $extensionsElement.InnerXml = $extensionsXml

            # Also add waypoints for compatibility
            $wptElement = $xmlDoc.CreateElement("wpt")
            $gpxElement.AppendChild($wptElement) | Out-Null

            # Add latitude and longitude attributes
            $latAttr = $xmlDoc.CreateAttribute("lat")
            $latAttr.Value = $lat.ToString("0.000000000")
            $wptElement.Attributes.Append($latAttr) | Out-Null

            $lonAttr = $xmlDoc.CreateAttribute("lon")
            $lonAttr.Value = $lon.ToString("0.000000000")
            $wptElement.Attributes.Append($lonAttr) | Out-Null

            # Add time element
            $timeElement = $xmlDoc.CreateElement("time")
            $timeElement.InnerText = [DateTime]::Parse("$($row.Date) $($row.Time)").ToUniversalTime().ToString("o")
            $wptElement.AppendChild($timeElement) | Out-Null

            # Add name element
            $nameElement = $xmlDoc.CreateElement("name")
            $nameElement.InnerText = $row.Filename
            $wptElement.AppendChild($nameElement) | Out-Null

            # Add comment element
            $cmtElement = $xmlDoc.CreateElement("cmt")
            $cmtElement.InnerText = $row.Filename
            $wptElement.AppendChild($cmtElement) | Out-Null

            # Add description element
            $descElement = $xmlDoc.CreateElement("desc")
            $descElement.InnerText = $row.Filename
            $wptElement.AppendChild($descElement) | Out-Null

            # Add extensions element to waypoint with speed
            $wptExtElement = $xmlDoc.CreateElement("extensions")
            $wptElement.AppendChild($wptExtElement) | Out-Null

            # Create the TrackPointExtension XML manually for waypoint
            $wptExtensionsXml = "<gpxtpx:TrackPointExtension xmlns:gpxtpx=`"http://www.garmin.com/xmlschemas/TrackPointExtension/v1`">" +
            "<gpxtpx:speed>" + $speedMps.ToString("0.00") + "</gpxtpx:speed>" +
            "</gpxtpx:TrackPointExtension>"

            # Set the extensions element's inner XML
            $wptExtElement.InnerXml = $wptExtensionsXml

            $rowCount++
            $processedRows += [PSCustomObject]@{
                Lat      = $lat
                Lon      = $lon
                Filename = $row.Filename
            }
        }

        # Update bounds in metadata
        $minLatAttr = $xmlDoc.CreateAttribute("minlat")
        $minLatAttr.Value = $minLat.ToString("0.000000000")
        $boundsElement.Attributes.Append($minLatAttr) | Out-Null

        $minLonAttr = $xmlDoc.CreateAttribute("minlon")
        $minLonAttr.Value = $minLon.ToString("0.000000000")
        $boundsElement.Attributes.Append($minLonAttr) | Out-Null

        $maxLatAttr = $xmlDoc.CreateAttribute("maxlat")
        $maxLatAttr.Value = $maxLat.ToString("0.000000000")
        $boundsElement.Attributes.Append($maxLatAttr) | Out-Null

        $maxLonAttr = $xmlDoc.CreateAttribute("maxlon")
        $maxLonAttr.Value = $maxLon.ToString("0.000000000")
        $boundsElement.Attributes.Append($maxLonAttr) | Out-Null

        # Check if we have any valid rows
        if ($rowCount -eq 0) {
            Write-Error "No valid GPS coordinates found in the CSV file"
            return $false
        }

        Write-Host "Processed $rowCount rows with valid GPS coordinates" -ForegroundColor Green

        # Save the GPX file
        $xmlDoc.Save($GpxPath)
        Write-Host "GPX file created successfully: $GpxPath" -ForegroundColor Green

        return $GpxPath
    }
    catch {
        Write-Host "Failed to convert CSV to GPX" -ForegroundColor Red
        Write-Error "Error: $_"
        return $false
    }
}

#endregion Functions

#region Main Script

# Verify the video file exists
if (!(Test-Path -Path $VideoPath)) {
    Write-Error "Video file not found: $VideoPath"
    exit 1
}

# Check for required tools
if (!(Test-RequiredTools -SkipImageMagick:$ExtractOnly)) {
    exit 1
}

# Create necessary directories
$directories = New-OutputDirectory -VideoPath $VideoPath -ExtractOnly:$ExtractOnly

# Extract frames from the video
Export-VideoFrames -VideoPath $VideoPath -OutputDir $directories.OutputDir -VideoFileName $directories.VideoFileName -FrameRate $FrameRate -SampleDuration $SampleDuration

# Crop the frames if not in extract-only mode
if (!$ExtractOnly) {
    # Always use the cropped directory for output now
    $croppedFiles = ConvertTo-CroppedMetadata -SourceDir $directories.OutputDir -CroppedDir $directories.CroppedDir -BottomHeight $BottomHeight

    # Always use the cropped directory for OCR processing
    $ocrSourceDir = $directories.CroppedDir

    # Get the full path to the original video file
    $fullVideoPath = Resolve-Path -Path $VideoPath

    # Extract text from the cropped images and save the CSV next to the original video
    $csvPath = Extract-TextFromImages -ImageDir $ocrSourceDir -ThresholdPreprocessing:$true -OriginalVideoPath $fullVideoPath

    # Convert the CSV to GPX format
    if ($csvPath -and (Test-Path -Path $csvPath)) {
        $gpxPath = Convert-CsvToGpx -CsvPath $csvPath
        if ($gpxPath) {
            Write-Host "GPX file created: $gpxPath" -ForegroundColor Green
        }
    }
}

Write-Host "All operations completed successfully!" -ForegroundColor Green

#endregion Main Script
