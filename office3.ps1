# Silent MS Office Installation Script

# Ensure running as Administrator
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Explicitly defined paths
$setupPath = "F:\setup.exe"
$configXmlPath = "C:\OfficeInstall-config.xml"

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
    # Verbose Checking of Paths
    if (-Not (Test-Path $setupPath)) {
        throw "Setup executable not found at $setupPath"
    }

    if (-Not (Test-Path $configXmlPath)) {
        throw "Configuration XML not created successfully"
    }

    # Precise Setup Execution
    Log-Message "Starting Office Installation"
    
    # Direct Call with Full Parameters
    $process = Start-Process "$setupPath" -ArgumentList "/configure", "$configXmlPath", "/quiet" -Wait -PassThru

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
    # Optional: Clean up configuration file
    Remove-Item -Path $configXmlPath -Force -ErrorAction SilentlyContinue
}

# Optional system restart
# Uncomment if you want automatic restart after installation
# Restart-Computer -Force
