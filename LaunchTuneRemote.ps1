# Load Windows Forms assembly for message boxes
Add-Type -AssemblyName System.Windows.Forms

function Launch-SystemRestore {
    try {
        Start-Process "rstrui.exe"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to launch System Restore. Please run it manually from the Control Panel.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Is-Installed($programName) {
    $x86 = ((Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$programName*" } ).Length -gt 0
    $x64 = ((Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$programName*" } ).Length -gt 0
    return $x86 -or $x64
}

# Check if Chocolatey is installed, if not, install it
if (!(Get-Command choco.exe -ErrorAction SilentlyContinue)) {
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
}

# Check if Glary Utilities and O&O ShutUp10 are installed
$glaryInstalled = Is-Installed "Glary Utilities"
$shutupInstalled = Is-Installed "O&O ShutUp10"

$restorePrompt = [System.Windows.Forms.MessageBox]::Show(
    "Do you want to create a restore point before making changes?",
    "Create Restore Point",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($restorePrompt -eq 'Yes') {
    Launch-SystemRestore
    [System.Windows.Forms.MessageBox]::Show("Please create a restore point in the opened window before continuing.", "Create Restore Point", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

if ($glaryInstalled -or $shutupInstalled) {
    $uninstallPrompt = [System.Windows.Forms.MessageBox]::Show(
        "Glary Utilities and/or O&O ShutUp10 are already installed. Do you want to uninstall them?",
        "Uninstall Software",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($uninstallPrompt -eq 'Yes') {
        # Uninstall Glary Utilities if installed
        if ($glaryInstalled) {
            choco uninstall glaryutilities-free -y
        }

        # Uninstall O&O ShutUp10 if installed
        if ($shutupInstalled) {
            choco uninstall shutup10 -y
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

    if ($installPrompt -eq 'Yes') {
        # Install Glary Utilities and O&O ShutUp10
        choco install glaryutilities-free -y
        choco install shutup10 -y

        [System.Windows.Forms.MessageBox]::Show("Installation complete.", "Install Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
}

# Prompt for system reboot
$rebootPrompt = [System.Windows.Forms.MessageBox]::Show(
    "Do you want to restart your computer now to ensure all changes take effect?",
    "Restart Computer",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($rebootPrompt -eq 'Yes') {
    Restart-Computer -Force
} else {
    [System.Windows.Forms.MessageBox]::Show("Please remember to restart your computer later to complete the process.", "Reminder", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}