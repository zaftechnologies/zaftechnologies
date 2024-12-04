# Pip install 
# Comprehensive Python Package Installation Script for Windows

# Ensure running with administrative privileges
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Logging function
function Log-Message {
    param([string]$Message, [string]$Type = "Info")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    Write-Host $logMessage
    
    # Log to file
    Add-Content -Path "C:\PythonPackageInstall.log" -Value $logMessage
}

# Check Python installation
function Test-PythonInstallation {
    try {
        $pythonVersion = & python --version
        Log-Message "Python version found: $pythonVersion" -Type "Success"
        return $true
    }
    catch {
        Log-Message "Python is not installed or not in PATH" -Type "Error"
        return $false
    }
}

# Upgrade pip and setuptools
function Upgrade-PipSetuptools {
    try {
        Start-Process python -ArgumentList "-m", "pip", "install", "--upgrade", "pip", "setuptools", "--user" -Wait -NoNewWindow -PassThru | Out-Null
        Log-Message "Pip and setuptools upgraded successfully" -Type "Success"
    }
    catch {
        Log-Message "Failed to upgrade pip and setuptools: $_" -Type "Error"
    }
}

# Install packages with retry mechanism
function Install-PythonPackages {
    param([string[]]$Packages)

    foreach ($package in $Packages) {
        $attempts = 0
        $maxAttempts = 3
        $success = $false

        while ($attempts -lt $maxAttempts -and -not $success) {
            $attempts++
            try {
                Log-Message "Installing $package (Attempt $attempts)" -Type "Info"
                
                # Special handling for specific packages
                $installArgs = if ($package -eq "tensorflow") {
                    @("-m", "pip", "install", $package, "--user", "--no-cache-dir")
                } else {
                    @("-m", "pip", "install", $package, "--user")
                }

                $process = Start-Process python -ArgumentList $installArgs -Wait -NoNewWindow -PassThru
                
                # Check exit code
                if ($process.ExitCode -eq 0) {
                    Log-Message "$package installed successfully" -Type "Success"
                    $success = $true
                }
                else {
                    Log-Message "Failed to install $package with exit code $($process.ExitCode)" -Type "Error"
                }
            }
            catch {
                Log-Message "Error installing $package: $_" -Type "Error"
            }

            # Wait between attempts
            if (-not $success) {
                Start-Sleep -Seconds (5 * $attempts)
            }
        }

        # Final check
        if (-not $success) {
            Log-Message "Could not install $package after $maxAttempts attempts" -Type "Critical"
        }
    }
}

# Main script execution
try {
    # List of packages to install
    $packages = @(
        'pandas',
        'numpy',
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

    # Verify Python installation
    if (-not (Test-PythonInstallation)) {
        throw "Python is not installed. Please install Python first."
    }

    # Upgrade pip and setuptools
    Upgrade-PipSetuptools

    # Install packages
    Install-PythonPackages -Packages $packages

    Log-Message "Package installation process completed" -Type "Success"
}
catch {
    Log-Message "Critical error in installation process: $_" -Type "Critical"
}

# Optional: Pause to view results (comment out in production)
# Pause
