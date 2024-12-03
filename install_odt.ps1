# Define paths
$odtPath = "C:\ODT\setup.exe"
$configXmlPath = "C:\ODT\config.xml"

# Check if ODT setup.exe exists
if (Test-Path $odtPath) {
    Write-Host "Starting Office 2016 installation..."
    Start-Process -FilePath $odtPath -ArgumentList "/configure $configXmlPath" -Wait
    Write-Host "Office 2016 installation completed."
} else {
    Write-Host "Error: ODT setup.exe not found!"
}
