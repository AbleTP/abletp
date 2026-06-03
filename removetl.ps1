# =========================================================
# ThreatLocker Full Uninstall Script
# =========================================================

$ErrorActionPreference = 'SilentlyContinue'

$StubPath = "C:\ThreatLockerStub.exe"
$LogPath  = "C:\ThreatLocker_Uninstall_Log.txt"

function Write-Log {
    param([string]$Message)
    $Line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Write-Host $Line
    Add-Content -Path $LogPath -Value $Line
}

Write-Log "Starting ThreatLocker removal..."

# Determine architecture
if ([Environment]::Is64BitOperatingSystem) {
    $DownloadUrl = "https://api.threatlocker.com/updates/installers/threatlockerstubx64.exe"
}
else {
    $DownloadUrl = "https://api.threatlocker.com/updates/installers/threatlockerstubx86.exe"
}

# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Log "Downloading ThreatLocker uninstall stub..."

try {
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $StubPath -UseBasicParsing
}
catch {
    Write-Log "Invoke-WebRequest failed. Trying BITS..."
    Start-BitsTransfer -Source $DownloadUrl -Destination $StubPath
}

if (-not (Test-Path $StubPath)) {
    Write-Log "Failed to download ThreatLocker stub."
    exit 1
}

Write-Log "Running ThreatLocker uninstall..."

$Process = Start-Process `
    -FilePath $StubPath `
    -ArgumentList "uninstall" `
    -Wait `
    -PassThru

Start-Sleep -Seconds 20

Write-Log "Stopping remaining ThreatLocker services if present..."

Get-Service | Where-Object {
    $_.Name -like "*ThreatLocker*" -or $_.DisplayName -like "*ThreatLocker*"
} | ForEach-Object {
    Stop-Service $_.Name -Force
}

Write-Log "Removing remaining ThreatLocker services if present..."

Get-Service | Where-Object {
    $_.Name -like "*ThreatLocker*" -or $_.DisplayName -like "*ThreatLocker*"
} | ForEach-Object {
    sc.exe delete $_.Name | Out-Null
}

Write-Log "Removing ThreatLocker scheduled tasks..."

Get-ScheduledTask | Where-Object {
    $_.TaskName -like "*ThreatLocker*" -or $_.TaskPath -like "*ThreatLocker*"
} | ForEach-Object {
    Unregister-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -Confirm:$false
}

Write-Log "Removing leftover folders..."

$Paths = @(
    "C:\Program Files\ThreatLocker",
    "C:\Program Files (x86)\ThreatLocker",
    "C:\ProgramData\ThreatLocker",
    "C:\ThreatLockerStub.exe"
)

foreach ($Path in $Paths) {
    if (Test-Path $Path) {
        Remove-Item $Path -Recurse -Force
        Write-Log "Removed: $Path"
    }
}

Write-Log "Checking final status..."

$RemainingServices = Get-Service | Where-Object {
    $_.Name -like "*ThreatLocker*" -or $_.DisplayName -like "*ThreatLocker*"
}

$RemainingFolders = $Paths | Where-Object { Test-Path $_ }

if (-not $RemainingServices -and -not $RemainingFolders) {
    Write-Log "ThreatLocker removed successfully."
    exit 0
}
else {
    Write-Log "ThreatLocker removal completed, but leftovers were detected."
    exit 1
}
