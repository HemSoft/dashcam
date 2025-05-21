FROM mcr.microsoft.com/powershell:latest

# Install necessary tools
RUN apt-get update && apt-get install -y \
    ffmpeg \
    wget \
    unzip \
    curl \
    software-properties-common \
    python3-pip \
    && rm -rf /var/lib/apt/lists/*

# Create directory for GPXSee
RUN mkdir -p /opt/gpxsee

# Download and install GPXSee
RUN wget -q https://github.com/tumic0/GPXSee/releases/download/12.9/GPXSee-12.9_jammy_amd64.deb -O /tmp/gpxsee.deb \
    && dpkg -i /tmp/gpxsee.deb || apt-get -f install -y \
    && rm /tmp/gpxsee.deb

# Install PowerShell modules
RUN pwsh -c "Install-Module -Name PSScriptAnalyzer -Force -Scope AllUsers"

# Set up working directory
WORKDIR /workspace

# Create wrapper script for GPXSee
RUN echo '#!/bin/bash\ngpxsee "$@"' > /usr/local/bin/gpxsee-wrapper \
    && chmod +x /usr/local/bin/gpxsee-wrapper

# Copy dashboard processing scripts to container
COPY *.ps1 /workspace/
