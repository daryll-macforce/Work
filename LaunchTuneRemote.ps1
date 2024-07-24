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

function Is-OOShutUp10Present {
    $shutupPath = "$env:ProgramData\chocolatey\lib\shutup10\tools\OOSU10.exe"
    $exists = Test-Path $shutupPath
    Write-Log "O&O ShutUp10 executable presence: $exists"
    return $exists
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
$shutupInstalled = Is-OOShutUp10Present

$installPrompt = [System.Windows.Forms.MessageBox]::Show(
    "Do you want to install/update Glary Utilities and O&O ShutUp10 to the latest versions?",
    "Install/Update Software",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

Write-Log "User response to install/update prompt: $installPrompt"

if ($installPrompt -eq 'Yes') {
    Write-Log "Attempting to install/update latest version of Glary Utilities"
    choco upgrade glaryutilities-free -y --force
    Write-Log "Glary Utilities install/update command completed"

    Write-Log "Attempting to install/update O&O ShutUp10"
    choco upgrade shutup10 -y --force
    Write-Log "O&O ShutUp10 install/update command completed"

    if (Is-OOShutUp10Present) {
        [System.Windows.Forms.MessageBox]::Show("O&O ShutUp10 has been successfully installed/updated. You can find it at: $env:ProgramData\chocolatey\lib\shutup10\tools\OOSU10.exe", "O&O ShutUp10 Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show("O&O ShutUp10 installation could not be verified. Please check the log file for more information.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }

    [System.Windows.Forms.MessageBox]::Show("Installation/Update complete.", "Install/Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

Write-Log "Script completed"