# Silent MS Office Installation Script with /adminfile

# Ensure running as Administrator
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Explicitly defined paths
$setupPath = "F:\setup.exe"
$configXmlPath = "C:\OfficeInstall-config.xml"
$adminfilePath = "C:\OfficeInstall-adminfile.mst"

# Create Admin Transform File (MST)
$adminfileContent = @"
[Setup]
Lang=english
Dir=C:\Program Files\Microsoft Office
SetupCompanyName=YourCompanyName
"@
$adminfileContent | Out-File -FilePath $adminfilePath -Encoding Unicode

# Detailed Configuration XML
$configXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
    </Product>
  </Add>
  <Display Level="Silent" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="%temp%\OfficeInstallLogs" />
  <Updates Enabled="TRUE" />
</Configuration>
"@

# Save Configuration XML
$configXml | Out-File -FilePath $configXmlPath -Encoding UTF8

# Detailed Logging Function
function Log-Message {
    param([string]$Message, [string]$Type = "Info")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    Write-Host $logMessage
    
    # Optional: Log to file
    Add-Content -Path "C:\OfficeInstallLog.txt" -Value $logMessage
}

try {
    # Verify Paths
    if (-Not (Test-Path $setupPath)) {
        throw "Setup executable not found at $setupPath"
    }

    if (-Not (Test-Path $configXmlPath)) {
        throw "Configuration XML not created successfully"
    }

    # Verbose Installation Command
    Log-Message "Starting Office Installation"
    
    # Comprehensive Setup Execution with Multiple Parameters
    $process = Start-Process "$setupPath" -ArgumentList `
        "/config", "$configXmlPath", `
        "/adminfile", "$adminfilePath", `
        "/quiet", `
        "/norestart" `
        -Wait -PassThru -NoNewWindow

    # Check Exit Code
    if ($process.ExitCode -eq 0) {
        Log-Message "Office installation completed successfully" -Type "Success"
    } else {
        throw "Installation failed with exit code $($process.ExitCode)"
    }
}
catch {
    Log-Message "Installation Error: $_" -Type "Error"
    
    # Additional Diagnostics
    Log-Message "Attempting to get setup.exe help information" -Type "Diagnostic"
    $helpOutput = & $setupPath /?
    Log-Message "$helpOutput" -Type "Diagnostic"
}
finally {
    # Clean up configuration and adminfile
    Remove-Item -Path $configXmlPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $adminfilePath -Force -ErrorAction SilentlyContinue
}

# Optional system restart
# Uncomment if you want automatic restart after installation
# Restart-Computer -Force
