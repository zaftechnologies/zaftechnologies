# Silent MS Office Installation Script

# Elevate script to admin privileges if not already running as admin
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs
        Exit
    }
}

# Set variables
$isoPath = "C:\office.iso"
$setupPath = "F:\setup.exe"
$configXmlPath = "C:\OfficeInstall-config.xml"

# Create configuration XML for silent installation
$configXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current" AllowCdnFallback="TRUE">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
    </Product>
  </Add>
  <Display Level="Silent" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="%temp%\OfficeInstallLogs" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
</Configuration>
"@

# Save configuration XML
$configXml | Out-File -FilePath $configXmlPath -Encoding UTF8

try {
    # Mount ISO
    $mountResult = Mount-DiskImage -ImagePath $isoPath -PassThru
    $driveLetter = ($mountResult | Get-Volume).DriveLetter + ":"

    # Run Office Setup silently with correct command-line parameters
    Start-Process "$setupPath" -ArgumentList "/configure", "$configXmlPath" -Wait -NoNewWindow -PassThru

    Write-Host "Office installation completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "Error during Office installation: $_" -ForegroundColor Red
}
finally {
    # Dismount ISO
    Dismount-DiskImage -ImagePath $isoPath
}

# Optional: Clean up configuration XML after installation
Remove-Item -Path $configXmlPath -Force -ErrorAction SilentlyContinue

# Optional: Restart computer to complete installation
# Uncomment the following line if you want to automatically restart
# Restart-Computer -Force
