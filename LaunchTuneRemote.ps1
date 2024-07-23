# Load Windows Forms assembly for message boxes
Add-Type -AssemblyName System.Windows.Forms

# Set up logging
$logFile = "$env:TEMP\LaunchTuneRemote_log.txt"
function Write-Log {
    param($message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -FilePath $logFile -Append
}

Write-Log "Script started"

function Is-Installed($programName) {
    Write-Log "Checking if $programName is installed"
    $x86 = ((Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$programName*" } ).Length -gt 0
    $x64 = ((Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$programName*" } ).Length -gt 0
    Write-Log "$programName installation status: $($x86 -or $x64)"
    return $x86 -or $x64
}

# Check if Chocolatey is installed, if not, install it
Write-Log "Checking Chocolatey installation"
if (!(Get-Command choco.exe -ErrorAction SilentlyContinue)) {
    Write-Log "Chocolatey not found, attempting to install"
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
        Write-Log "Chocolatey installed successfully"
    }
    catch {
        Write-Log "Error installing Chocolatey: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to install Chocolatey. Script cannot continue.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}

# Check if Glary Utilities and O&O ShutUp10 are installed
$glaryInstalled = Is-Installed "Glary Utilities"
$shutupInstalled = Is-Installed "O&O ShutUp10"

if ($glaryInstalled -or $shutupInstalled) {
    $uninstallPrompt = [System.Windows.Forms.MessageBox]::Show(
        "Glary Utilities and/or O&O ShutUp10 are already installed. Do you want to uninstall them?",
        "Uninstall Software",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    Write-Log "User response to uninstall prompt: $uninstallPrompt"

    if ($uninstallPrompt -eq 'Yes') {
        if ($glaryInstalled) {
            Write-Log "Attempting to uninstall Glary Utilities"
            choco uninstall glaryutilities-free -y
            Write-Log "Glary Utilities uninstall command completed"
        }

        if ($shutupInstalled) {
            Write-Log "Attempting to uninstall O&O ShutUp10"
            choco uninstall shutup10 -y
            Write-Log "O&O ShutUp10 uninstall command completed"
        }

        [System.Windows.Forms.MessageBox]::Show("Uninstallation complete.", "Uninstall Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
} else {
    $installPrompt = [System.Windows.Forms.MessageBox]::Show(
        "Glary Utilities and O&O ShutUp10 are not installed. Do you want to install them?",
        "Install Software",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    Write-Log "User response to install prompt: $installPrompt"

    if ($installPrompt -eq 'Yes') {
        Write-Log "Attempting to install Glary Utilities"
        choco install glaryutilities-free -y
        Write-Log "Glary Utilities install command completed"

        Write-Log "Attempting to install O&O ShutUp10"
        choco install shutup10 -y
        Write-Log "O&O ShutUp10 install command completed"

        [System.Windows.Forms.MessageBox]::Show("Installation complete.", "Install Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
}

Write-Log "Script completed"