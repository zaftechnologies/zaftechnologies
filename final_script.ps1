Invoke-WebRequest -Uri "https://zaftech-terraform.s3.us-east-1.amazonaws.com/SW_DVD5_Office_2016_64Bit_English_MLF_X20-42479.ISO" -OutFile "C:\office.iso"

$isoPath = "C:\office.iso"

# Mount the ISO file
Mount-DiskImage -ImagePath $isoPath

# Wait a moment for the ISO to mount and get the drive letter
Start-Sleep -Seconds 5

# Get the drive letter assigned to the mounted ISO
$driveLetter = (Get-DiskImage -ImagePath $isoPath | Get-Volume).DriveLetter

# Check if the ISO is mounted and assign it to F: using subst
if ($driveLetter) {
    Write-Host "ISO mounted to $driveLetter"
    
    # Remove existing F: mapping if needed
    if (Test-Path "F:\") {
        Write-Host "Removing existing F: drive mapping."
        subst F: /d
    }
    
    # Use subst to map the ISO to F:
    subst F: "$($driveLetter):\"
    Write-Host "ISO is now mapped to F:"
} else {
    Write-Host "Failed to mount the ISO."
}


# Create admin.msp file 
#.\setup.exe /admin

$installerPath = "E:\setup.exe"
$licenseKey = "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
$logFile = "C:\path\to\logfile.txt"
