# Path to your ISO file
$isoPath = "C:\office"

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
Write-host "Creating xml file"
# Define the path where the configuration file will be saved
$configFilePath = "C:\temp\configuration.xml"  # Replace with your desired path

# Create the XML structure
[xml]$xmlContent = New-Object -TypeName System.Xml.XmlDocument

# Create the root <Configuration> element
$configurationElement = $xmlContent.CreateElement("Configuration")

# Create the <Add> element with attributes
$addElement = $xmlContent.CreateElement("Add")
$addElement.SetAttribute("OfficeClientEdition", "64")
$addElement.SetAttribute("Channel", "Monthly")

# Create the <Product> element with nested <Language> element
$productElement = $xmlContent.CreateElement("Product")
$productElement.SetAttribute("ID", "O365ProPlusRetail")
$languageElement = $xmlContent.CreateElement("Language")
$languageElement.SetAttribute("ID", "en-us")

# Append the <Language> element to the <Product> element
$productElement.AppendChild($languageElement)

# Append the <Product> element to the <Add> element
$addElement.AppendChild($productElement)

# Append the <Add> element to the root <Configuration> element
$configurationElement.AppendChild($addElement)

# Create the <Display> element with attributes
$displayElement = $xmlContent.CreateElement("Display")
$displayElement.SetAttribute("Level", "None")
$displayElement.SetAttribute("AcceptEULA", "TRUE")

# Append the <Display> element to the root <Configuration> element
$configurationElement.AppendChild($displayElement)

# Create the <Property> element
$propertyElement = $xmlContent.CreateElement("Property")
$propertyElement.SetAttribute("Name", "ForceAppShutdown")
$propertyElement.SetAttribute("Value", "TRUE")

# Append the <Property> element to the root <Configuration> element
$configurationElement.AppendChild($propertyElement)

# Append the <Configuration> element to the document
$xmlContent.AppendChild($configurationElement)

# Save the XML content to the file
$xmlContent.Save($configFilePath)

Write-Host "Configuration file created at: $configFilePath"

Write-host "Office installation"

$officeSetupPath = "f:/setup.exe"  

# Check if the setup.exe exists
if (Test-Path $officeSetupPath) {
    Write-Host "Starting Office installation..."

    # Run the setup.exe with silent installation arguments
    Start-Process -FilePath $officeSetupPath -ArgumentList "/quiet", "/norestart" -Wait

    Write-Host "Office installation completed silently."
} else {
    Write-Host "Setup file not found: $officeSetupPath"
}

