# Define paths for ODT installation
$odtUrl = "https://download.microsoft.com/download/1/1/2/1124D1F9-DF9A-4C82-919F-26772AAE9D7E/OfficeDeploymentTool_4321.20218.404.0.exe"
$odtInstaller = "C:\ODT\ODT_installer.exe"
$odtSetupPath = "C:\ODT\setup.exe"
$configXmlPath = "C:\ODT\config.xml"

# Download ODT installer if not exists
if (-not (Test-Path $odtInstaller)) {
    Write-Host "Downloading ODT installer..."
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtInstaller
}

# Install ODT silently
Write-Host "Installing Office Deployment Tool..."
Start-Process -FilePath $odtInstaller -ArgumentList "/quiet /install" -Wait

# Verify the presence of setup.exe and config.xml
if (Test-Path $odtSetupPath -and Test-Path $configXmlPath) {
    Write-Host "Starting Office installation..."

    # Run Office installation silently with EULA acceptance
    Start-Process -FilePath $odtSetupPath -ArgumentList "/configure $configXmlPath" -Wait

    Write-Host "Office installation completed."
} else {
    Write-Host "Error: setup.exe or config.xml not found."
}
