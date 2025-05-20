# Dashcam Video Processor

A PowerShell toolkit for processing dashcam footage, extracting metadata, and analyzing captured information.

## Features

- Extract frames from dashcam videos at specified frame rates
- Crop frames to isolate the metadata section (speed, coordinates, timestamp)
- Perform OCR on metadata sections to extract text information
- Generate CSV files with extracted data (date, time, speed, coordinates)
- Optional frame sample mode for quick testing
- Support for maintaining original and cropped images

## Requirements

- PowerShell 5.1+
- [FFmpeg](https://ffmpeg.org/download.html) for video frame extraction
- [ImageMagick](https://imagemagick.org/script/download.php) for image cropping and processing
- [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) for text extraction

## Usage

```powershell
.\Process-DashcamVideo.ps1 -VideoPath "path\to\video.mp4" [options]
```

### Parameters

- `-VideoPath` (Required): Path to the dashcam video file
- `-FrameRate` (Optional): Number of frames to extract per second (default: 1)
- `-BottomHeight` (Optional): Height in pixels of the bottom metadata strip (default: 100)
- `-ExtractOnly` (Optional): Extract frames only without cropping
- `-KeepOriginalFrames` (Optional): Keep both original and cropped frames
- `-SampleDuration` (Optional): Process only the first X seconds of the video (default: 0 = entire video)

### Example

```powershell
# Process 30 seconds of a dashcam video, keeping original frames
.\Process-DashcamVideo.ps1 -VideoPath ".\2024\my_video.mp4" -KeepOriginalFrames -SampleDuration 30
```

## Output

- Extracted frames in a folder named after the video file
- Cropped metadata frames in a subfolder (if `-KeepOriginalFrames` is used)
- CSV file with extracted metadata (filename, date, time, speed, coordinates)

## Notes

- Default OCR settings are optimized for common dashcam metadata formats
- Use the `-SampleDuration` parameter to test processing on a short segment first
- Adjust the `-BottomHeight` parameter based on your dashcam's metadata strip size
