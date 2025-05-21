# Dashcam Processing Development Container

This directory contains the configuration for a development container that can be used with Visual Studio Code or GitHub Codespaces to create a consistent development environment for working with dashcam video processing scripts.

## Features

The dev container includes:

- PowerShell Core
- FFmpeg for video processing
- GPXSee for GPX file visualization
- Git and GitHub CLI
- PowerShell Script Analyzer for linting

## Usage

### With Visual Studio Code

1. Install the [Remote - Containers](https://marketplace.visualstudio.com/items?itemName=ms-vscode-remote.remote-containers) extension
2. Open this repository in VS Code
3. When prompted, click "Reopen in Container"
4. Alternatively, use the Command Palette (F1) and select "Remote-Containers: Reopen in Container"

### With GitHub Codespaces

1. Navigate to the GitHub repository
2. Click the "Code" button
3. Select the "Codespaces" tab
4. Click "New codespace"

## GPXSee Access

To run GPXSee with GUI inside the container, you'll need to:

1. Uncomment the X11 display lines in docker-compose.yml
2. Configure X11 forwarding on your host system
3. Use the gpxsee-wrapper command or modify the Visualize-GPX.ps1 script to use the correct path

## Notes

- The PowerShell scripts are automatically copied to the workspace directory
- The container includes all necessary dependencies for processing dashcam videos and GPX files
- For the best experience, ensure VS Code is running with administrator privileges when working with external devices
