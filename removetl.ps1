# =========================================================
# ThreatLocker Uninstall Script
# =========================================================

$ErrorActionPreference = 'SilentlyContinue'

$StubPath = "C:\ThreatLockerStub.exe"

# Determine architecture
if ([Environment]::Is64BitOperatingSystem) {
    $DownloadUrl = "https://api.threatlocker.com/updates/installers/threatlockerstubx64.exe"
}
else {
    $DownloadUrl = "https://api.threatlocker.com/updates/installers/threatlockerstubx86.exe"
}

Write-Host "Downloading ThreatLocker uninstall stub..."

try {
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $StubPath -UseBasicParsing
}
catch {
    Write-Host "Invoke-WebRequest failed, attempting BITS download..."

    Start-BitsTransfer `
        -Source $DownloadUrl `
        -Destination $StubPath
}

if (-not (Test-Path $StubPath)) {
    Write-Host "Failed to download ThreatLocker stub."
    exit 1
}

Write-Host "Running ThreatLocker uninstall..."

$process = Start-Process `
    -FilePath $StubPath `
    -ArgumentList "uninstall" `
    -Wait `
    -PassThru

Start-Sleep -Seconds 10

Write-Host "Verifying uninstall..."

$Service = Get-Service -Name "ThreatLockerService" -ErrorAction SilentlyContinue

if ($null -eq $Service) {
    Write-Host "ThreatLocker uninstalled successfully"

    if (Test-Path $StubPath) {
        Remove-Item $StubPath -Force
    }

    exit 0
}
else {
    Write-Host "ThreatLocker failed to uninstall"

    if (Test-Path $StubPath) {
        Remove-Item $StubPath -Force
    }

    exit 1
}
