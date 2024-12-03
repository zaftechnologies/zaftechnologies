# Paths to the ODT setup and config file
$odtSetupPath = "C:\ODT\setup.exe"   # Path to setup.exe from ODT
$configXmlPath = "C:\ODT\config.xml" # Path to your config.xml

# Ensure setup.exe and config.xml exist
if (Test-Path $odtSetupPath -and Test-Path $configXmlPath) {
    Write-Host "Starting Office installation..."

    # Run ODT setup.exe to install Office silently and accept the EULA
    Start-Process -FilePath $odtSetupPath -ArgumentList "/configure $configXmlPath" -Wait

    Write-Host "Office installation completed."
} else {
    Write-Host "Error: setup.exe or config.xml not found."
}
