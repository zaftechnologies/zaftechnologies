# Define paths and parameters
$installerPath = "C:\Office2016\setup.exe"
$licenseKey = "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
$logFile = "C:\Office2016\install_log.txt"

# Check if the installer exists
if (Test-Path $installerPath) {
    Write-Host "Starting Office 2016 MSI installation..."
    Start-Process -FilePath $installerPath -ArgumentList "/quiet /norestart PIDKEY=$licenseKey /log $logFile" -Wait
    Write-Host "Office 2016 installation completed. Check the log at $logFile"
} else {
    Write-Host "Error: Office 2016 installer not found!"
}
