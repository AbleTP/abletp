# Set working directory
$workDir = "C:\ProgramData\AppDeploy"
New-Item -ItemType Directory -Path $workDir -Force | Out-Null
Set-Location $workDir

# Function to download files
function Download-File {
    param (
        [string]$url,
        [string]$destination
    )
    if (-not (Test-Path $destination)) {
        Invoke-WebRequest -Uri $url -OutFile $destination -UseBasicParsing
    } else {
        Write-Output "$destination already exists. Skipping download."
    }
}

# Function to check if app is installed
function Is-AppInstalled {
    param([string]$name)
    $keys = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($key in $keys) {
        $installed = Get-ItemProperty $key -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*$name*" }
        if ($installed) { return $true }
    }
    return $false
}

# 1. Adobe Reader DC
if (-not (Is-AppInstalled "Adobe Acrobat Reader")) {
    Write-Output "Installing Adobe Reader..."
    $adobeUrl = "https://ardownload3.adobe.com/pub/adobe/acrobat/win/AcrobatClassic/2400130123/Acrobat_Classic_Web_x64_WWMUI.exe"
    $adobeExe = "$workDir\AdobeReader.exe"
    Download-File $adobeUrl $adobeExe
    Start-Process $adobeExe -ArgumentList "/sAll /rs /rps /msi EULA_ACCEPT=YES" -Wait
} else {
    Write-Output "✅ Adobe Reader is already installed. Skipping."
}

# 2. Google Chrome
if (-not (Is-AppInstalled "Google Chrome")) {
    Write-Output "Installing Google Chrome..."
    $chromeUrl = "https://dl.google.com/chrome/install/375.126/chrome_installer.exe"
    $chromeExe = "$workDir\Chrome.exe"
    Download-File $chromeUrl $chromeExe
    Start-Process $chromeExe -ArgumentList "/silent /install" -Wait
} else {
    Write-Output "✅ Google Chrome is already installed. Skipping."
}

# 3. Zoom
if (-not (Is-AppInstalled "Zoom")) {
    Write-Output "Installing Zoom..."
    $zoomUrl = "https://zoom.us/client/latest/ZoomInstallerFull.msi"
    $zoomMsi = "$workDir\Zoom.msi"
    Download-File $zoomUrl $zoomMsi
    Start-Process msiexec.exe -ArgumentList "/i `"$zoomMsi`" /quiet /norestart" -Wait
} else {
    Write-Output "✅ Zoom is already installed. Skipping."
}

Write-Host "`n✅ Finished. Reboot may be required."
