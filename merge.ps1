# Define the directory containing your dashcam videos
$videoDir = "E:\Dashcam"
#$outputDir = Join-Path -Path $videoDir -ChildPath "Daily"
$outputDir = "F:\OneDrive\Media\Dashcam\2025"

# Ensure the output directory exists
if (!(Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Get all video files in the directory
$videoFiles = Get-ChildItem -Path $videoDir -Filter "dm_*.mp4"

# Group files by date (extract date from filename: dm_YYYYMMDD_HHMMSS.mp4)
$filesByDate = $videoFiles | Group-Object {
    ($_ -match "dm_(\d{8})_\d{6}" | Out-Null) 
    $Matches[1]
}

# Create a UTF8 encoding without BOM
$utf8NoBOM = New-Object System.Text.UTF8Encoding($false)

# Process each date group
foreach ($group in $filesByDate) {
    $date = $group.Name
    $files = $group.Group

    # Create a temporary file list for ffmpeg
    $fileListPath = Join-Path -Path $videoDir -ChildPath "filelist_$date.txt"

    # Generate properly formatted file list lines
    $fileListContent = $files | ForEach-Object {
        "file '" + $_.FullName.Replace('\', '/') + "'"
    }

    # Write out the lines without BOM using .NET directly
    [System.IO.File]::WriteAllLines($fileListPath, $fileListContent, $utf8NoBOM)

    # Define the output file name
    $outputFile = Join-Path -Path $outputDir -ChildPath "merged_$date.mp4"

    # Run ffmpeg to concatenate files
    $ffmpegCommand = "ffmpeg -f concat -safe 0 -i `"$fileListPath`" -c copy `"$outputFile`""
    Write-Host "Processing files for date: $date" -ForegroundColor Green
    Write-Host $ffmpegCommand -ForegroundColor Yellow
    Invoke-Expression $ffmpegCommand

    # Clean up the temporary file list
    # Remove-Item -Path $fileListPath
}

Write-Host "All videos have been processed and merged into: $outputDir" -ForegroundColor Cyan
