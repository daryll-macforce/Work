# Load Windows Forms assembly for message boxes
Add-Type -AssemblyName System.Windows.Forms

function Manage-SystemRestore {
    $systemRestoreEnabled = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore" -Name "RPSessionInterval").RPSessionInterval -ne 0
    
    if (-not $systemRestoreEnabled) {
        $enablePrompt = [System.Windows.Forms.MessageBox]::Show(
            "System Restore is currently disabled. Would you like to enable it?",
            "Enable System Restore",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($enablePrompt -eq 'Yes') {
            try {
                Enable-ComputerRestore -Drive "C:\"
                [System.Windows.Forms.MessageBox]::Show("System Restore has been enabled.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $systemRestoreEnabled = $true
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to enable System Restore. Continuing without creating a restore point.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("System Restore will remain disabled. Continuing without creating a restore point.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
    }
    
    if ($systemRestoreEnabled) {
        try {
            Checkpoint-Computer -Description "Before installing/uninstalling software" -RestorePointType "MODIFY_SETTINGS"
            [System.Windows.Forms.MessageBox]::Show("Restore point created successfully.", "Restore Point", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to create restore point. Continuing without creating one.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
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

if ($glaryInstalled -or $shutupInstalled) {
    $uninstallPrompt = [System.Windows.Forms.MessageBox]::Show(
        "Glary Utilities and/or O&O ShutUp10 are already installed. Do you want to uninstall them?",
        "Uninstall Software",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($uninstallPrompt -eq 'Yes') {
        Manage-SystemRestore

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
        Manage-SystemRestore

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