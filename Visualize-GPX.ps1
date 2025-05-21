# Visualize-GPX.ps1
# Script to visualize GPX files using GPXSee

param(
    [Parameter(Mandatory = $false)]
    [string]$GpxFile = "",
    
    [Parameter(Mandatory = $false)]
    [switch]$ListAll = $false,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowHelp = $false
)

# GPXSee executable path - adjust if installed elsewhere
$gpxseeExe = "C:\Program Files\GPXSee\gpxsee.exe"

if (!(Test-Path $gpxseeExe)) {
    # Try to find GPXSee in common locations
    $possiblePaths = @(
        "C:\Program Files\GPXSee\gpxsee.exe",
        "C:\Program Files (x86)\GPXSee\gpxsee.exe",
        "${env:LocalAppData}\Programs\GPXSee\gpxsee.exe"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $gpxseeExe = $path
            break
        }
    }
    
    if (!(Test-Path $gpxseeExe)) {
        Write-Host "GPXSee executable not found. Please install GPXSee or specify the correct path." -ForegroundColor Red
        exit 1
    }
}

function Show-Help {
    Write-Host "Visualize-GPX.ps1 - Tool to visualize GPX files using GPXSee" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  .\Visualize-GPX.ps1 [-GpxFile <path>] [-ListAll] [-ShowHelp]" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Parameters:" -ForegroundColor Yellow
    Write-Host "  -GpxFile     : Specify a GPX file to visualize" -ForegroundColor Green
    Write-Host "  -ListAll     : List all GPX files in the current directory and subdirectories" -ForegroundColor Green
    Write-Host "  -ShowHelp    : Display this help message" -ForegroundColor Green
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Yellow
    Write-Host "  .\Visualize-GPX.ps1 -GpxFile '.\2024\merged_20241129.gpx'" -ForegroundColor Green
    Write-Host "  .\Visualize-GPX.ps1 -ListAll" -ForegroundColor Green
    Write-Host ""
}

if ($ShowHelp) {
    Show-Help
    exit 0
}

if ($ListAll) {
    $gpxFiles = @(Get-ChildItem -Path "." -Filter "*.gpx" -Recurse | Select-Object -ExpandProperty FullName)
    
    if ($gpxFiles.Count -eq 0) {
        Write-Host "No GPX files found in the current directory and subdirectories." -ForegroundColor Yellow
        exit 0
    }
    
    Write-Host "Found $($gpxFiles.Count) GPX files:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $gpxFiles.Count; $i++) {
        Write-Host "  $($i+1): $($gpxFiles[$i])" -ForegroundColor Green
    }
      $selected = Read-Host "Enter the number of the file to visualize (1-$($gpxFiles.Count)), or press Enter to exit"
    
    if ($selected -match '^\d+$' -and [int]$selected -ge 1 -and [int]$selected -le $gpxFiles.Count) {
        $GpxFile = $gpxFiles[[int]$selected - 1]
        Write-Host "Opening $GpxFile..." -ForegroundColor Cyan
    } else {
        Write-Host "No file selected." -ForegroundColor Yellow
        exit 0
    }
}

if (-not [string]::IsNullOrEmpty($GpxFile)) {
    if (Test-Path $GpxFile) {
        # Start GPXSee with the specified GPX file
        $absolutePath = (Resolve-Path $GpxFile).Path
        Start-Process -FilePath $gpxseeExe -ArgumentList "`"$absolutePath`""
        Write-Host "Opened $GpxFile in GPXSee" -ForegroundColor Green
    } else {
        Write-Host "GPX file not found: $GpxFile" -ForegroundColor Red
        exit 1
    }
} else {
    Show-Help
}
