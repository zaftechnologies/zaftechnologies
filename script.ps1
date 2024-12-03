<powershell>

# Ensure PowerShell runs as Administrator
if (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Start-Process powershell.exe "-File $PSCommandPath" -Verb RunAs
    exit
}

# Set an admin username and password
$adminUsername = "data-terminal"
$adminPassword = ConvertTo-SecureString "Welcome@123!" -AsPlainText -Force

# Create the admin user
New-LocalUser -Name $adminUsername -Password $adminPassword -FullName "Admin User" -Description "Administrator account"
Add-LocalGroupMember -Group "Administrators" -Member $adminUsername

$credential = [PSCredential]::New($adminUsername, $adminPassword)
Start-Process "cmd.exe" -Credential $credential -ArgumentList "/C"
New-Item -ItemType Directory -Force -Path "C:\Users\$adminUsername\.aws"
New-Item -ItemType Directory -Force -Path "C:\Users\$adminUsername\Desktop"


# Create AWS Config for authenticating as EC2 Role
$awsConfig = @"
[default]
credential_source=Ec2InstanceMetadata
"@
Out-File -FilePath "C:\Users\$adminUsername\.aws\config" -InputObject $awsConfig -Encoding ascii

# Allow the user to log on via RDP
$rdpUsers = Get-WmiObject -Class Win32_Group -Filter "Name='Remote Desktop Users'"
$rdpUsers.Add("WinNT://./$adminUsername")

# Set the RDP configuration to allow connections
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name "fDenyTSConnections" -Value 0

# Enable firewall rule for RDP
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"

# Install Chocolatey
Set-ExecutionPolicy Bypass -Scope Process -Force
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

# Install Python
choco install -y python --version=3.9.7

# Install AWSCLIV2
Start-Process msiexec.exe -Wait -ArgumentList '/i https://awscli.amazonaws.com/AWSCLIV2.msi /qn'
[Environment]::SetEnvironmentVariable("Path", [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::Machine) + ";C:\Program Files\Amazon\AWSCLIV2", [EnvironmentVariableTarget]::Machine)

# Install R
choco install -y r.project --version=4.2.0

# Refresh the environment variables
Import-Module $env:ChocolateyInstall\helpers\chocolateyProfile.psm1
refreshenv

# asdf
python -m pip install --upgrade pip setuptools --user
$packages = @(
    'pandas',
    'numpy',
    'random',
    'seaborn',
    'matplotlib',
    'scikit-learn',
    'scipy',
    'tensorflow',
    'pyspark',
    'pyodbc',
    'shap',
    'sqlalchemy'
)
foreach ($package in $packages) {
    python -m pip install $package --user
}

choco install -y pycharm-community vscode
python -m pip install jupyter --user

$WshShell = New-Object -ComObject WScript.Shell

$Shortcut = $WshShell.CreateShortcut("c:\Users\$adminUsername\Desktop\Jupyter Notebook.lnk")
$Shortcut.TargetPath = [System.Environment]::ExpandEnvironmentVariables('%LOCALAPPDATA%') + "\\Programs\\Python\\Python39\\Scripts\\jupyter-notebook.exe"
$Shortcut.Save()

 PowerShell Script and Shortcut for starting RDS
$startupScript = @"
# Ensure RDS is running
aws lambda invoke --function-name ${rds_lambda_name} --payload "{\``"action\``": \``"start\``"}" --cli-binary-format raw-in-base64-out \Temp\aws_startup.txt
write-host -nonewline "`n"
get-content \Temp\aws_startup.txt
write-host -nonewline "`n`nPress any key to continue..."
`$null = `$Host.UI.RawUI.ReadKey('Noecho,IncludeKeyDown');
"@
Out-File -FilePath C:\startup_rds.ps1 -InputObject $startupScript -Encoding ascii
$Shortcut = $WshShell.CreateShortcut("c:\Users\$adminUsername\Desktop\Start Database.lnk")
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-ExecutionPolicy Bypass -File c:\\startup_rds.ps1"
$Shortcut.Save()

# PowerShell Script and Shortcut for getting RDS status
$statusScript = @"
# Ensure RDS is running
aws lambda invoke --function-name ${rds_lambda_name} --payload "{\``"action\``": \``"status\``"}" --cli-binary-format raw-in-base64-out \Temp\aws_status.txt
write-host -nonewline "`n"
get-content \Temp\aws_status.txt
write-host -nonewline "`n`nPress any key to continue..."
`$null = `$Host.UI.RawUI.ReadKey('Noecho,IncludeKeyDown');
"@
Out-File -FilePath C:\status_rds.ps1 -InputObject $statusScript -Encoding ascii
$Shortcut = $WshShell.CreateShortcut("c:\Users\$adminUsername\Desktop\Database Status.lnk")
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-ExecutionPolicy Bypass -File c:\\status_rds.ps1"
$Shortcut.Save()

# PowerShell Script and Shortcut for stopping RDS Server
$shutdownScript = @"
# Ensure RDS is stopped
aws lambda invoke --function-name ${rds_lambda_name} --payload "{\``"action\``": \``"stop\``"}" --cli-binary-format raw-in-base64-out \Temp\aws_shutdown.txt
write-host -nonewline "`n"
get-content \Temp\aws_shutdown.txt
write-host -nonewline "`n`nPress any key to continue..."
`$null = `$Host.UI.RawUI.ReadKey('Noecho,IncludeKeyDown');
"@
Out-File -FilePath C:\shutdown_rds.ps1 -InputObject $shutdownScript -Encoding ascii
$Shortcut = $WshShell.CreateShortcut("c:\Users\$adminUsername\Desktop\Stop Database.lnk")
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-ExecutionPolicy Bypass -File c:\\shutdown_rds.ps1"
$Shortcut.Save()

# Install R packages
$Rscript = @'
install.packages("dplyr", repos="http://cran.r-project.org")
install.packages("tidyr", repos="http://cran.r-project.org")
install.packages("tibble", repos="http://cran.r-project.org")
install.packages("ggplot2", repos="http://cran.r-project.org")
install.packages("caret", repos="http://cran.r-project.org")
install.packages("glmnet", repos="http://cran.r-project.org")
install.packages("ggpubr", repos="http://cran.r-project.org")
install.packages("corrplot", repos="http://cran.r-project.org")
install.packages("caTools", repos="http://cran.r-project.org")
install.packages("MASS", repos="http://cran.r-project.org")
install.packages("lmtest", repos="http://cran.r-project.org")
install.packages("kernlab", repos="http://cran.r-project.org")
install.packages("lubridate", repos="http://cran.r-project.org")
install.packages("RODBC", repos="http://cran.r-project.org")
install.packages("odbc", repos="http://cran.r-project.org")
install.packages("DBI", repos="http://cran.r-project.org")
'@

Out-File -FilePath C:\InstallRPackages.R -InputObject $Rscript -Encoding ascii
& "C:\Program Files\R\R-4.2.0\bin\Rscript.exe" C:\InstallRPackages.R


# Define the registry paths
$rdpPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"

# Ensure the policy path exists
if (!(Test-Path $rdpPolicyPath)) {
    New-Item -Path $rdpPolicyPath -Force
}
Set-ItemProperty -Path $rdpPolicyPath -Name "fDisableClip" -Value 1
Set-ItemProperty -Path $rdpPolicyPath -Name "fDisableClipTransfer" -Value 1
Set-ItemProperty -Path $rdpPolicyPath -Name "fDisableClipTransferSrv" -Value 1
Set-ItemProperty -Path $rdpPolicyPath -Name "fDisableCdm" -Value 1
Restart-Service -Name TermService -Force

# Download and install Microsoft Excel (part of Microsoft Office)
$ProgressPreference = "SilentlyContinue"
$officeSetupPath = "C:\\temp\\OfficeSetup.exe"
$configurationXMLPath = "C:\\temp\\Configuration.xml"

# Ensure temp directory exists
New-Item -ItemType Directory -Path "C:\\temp" -Force

# Download Office setup executable
Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/?linkid=525135" -OutFile $officeSetupPath

# Configuration XML for Office installation
$configXML = @"
<Configuration>
  <Add OfficeClientEdition="64">
    <Product ID="ProPlusRetail">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
"@

# Write the configuration XML to a file
$configXML | Out-File -FilePath $configurationXMLPath

# Install Office
Start-Process -FilePath $officeSetupPath -ArgumentList "/configure $configurationXMLPath" -Wait

# Cleanup
Remove-Item -Path $officeSetupPath
Remove-Item -Path $configurationXMLPath

# Install SSMS #
$filepath="\windows\temp\SSMS-Setup-ENU.exe"

$URL = "https://aka.ms/ssmsfullsetup"
$clnt = New-Object System.Net.WebClient
$clnt.DownloadFile($url,$filepath)

$params = " /Install /Passive"
Start-Process -FilePath $filepath -ArgumentList $params -Wait

$ssms = Get-ChildItem -Path "c:\Program Files (x86)" -Filter ssms.exe -Recurse -ErrorAction SilentlyContinue -Force | % {$_.FullName}
$Shortcut = $WshShell.CreateShortcut("c:\Users\$adminUsername\Desktop\SSMS.lnk")
$Shortcut.TargetPath = $ssms
$Shortcut.Save()
</powershell>
