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

function Download-FromGitHub {
    param($fileName)
    $url = "https://github.com/daryll-macforce/Work/raw/main/Apps/$fileName"
    $outPath = Join-Path $env:TEMP $fileName
    Write-Log "Downloading $fileName from GitHub"
    try {
        Invoke-WebRequest -Uri $url -OutFile $outPath
        Write-Log "$fileName downloaded successfully"
        return $outPath
    }
    catch {
        Write-Log ("Error downloading " + $fileName + ": " + $_.Exception.Message)
        return $null
    }
}

function Install-Program {
    param($filePath, $programName)
    if ($filePath) {
        Write-Log "Starting installation of $programName"
        try {
            Start-Process -FilePath $filePath -Wait
            Write-Log "$programName installation completed"
        }
        catch {
            Write-Log ("Error installing " + $programName + ": " + $_.Exception.Message)
            [System.Windows.Forms.MessageBox]::Show("Failed to install $programName. Please check the log for details.", "Installation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    else {
        Write-Log "$programName installation file not found"
        [System.Windows.Forms.MessageBox]::Show("$programName installation file not found. Please check your internet connection and try again.", "File Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

$installPrompt = [System.Windows.Forms.MessageBox]::Show(
    "Do you want to install O&O ShutUp10 and Glary Utilities?",
    "Install Software",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

Write-Log "User response to install prompt: $installPrompt"

if ($installPrompt -eq 'Yes') {
    # Download and install O&O ShutUp10
    $oosuPath = Download-FromGitHub "OOSU10.exe"
    Install-Program $oosuPath "O&O ShutUp10"

    # Download and install Glary Utilities
    $glaryPath = Download-FromGitHub "gusetup.exe"
    Install-Program $glaryPath "Glary Utilities"

    # Cleanup
    if ($oosuPath) { Remove-Item $oosuPath -ErrorAction SilentlyContinue }
    if ($glaryPath) { Remove-Item $glaryPath -ErrorAction SilentlyContinue }

    [System.Windows.Forms.MessageBox]::Show("Installation process completed. Please check the log for any issues.", "Installation Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

Write-Log "Script completed"