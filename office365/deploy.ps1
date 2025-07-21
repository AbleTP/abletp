# Define variables
$zipFile = "$PSScriptRoot\officePackage.zip"
$extractPath = "C:\ProgramData\officePackage"

# Create extract folder
if (!(Test-Path $extractPath)) {
    New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
}

# Extract ZIP
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile, $extractPath)

# Run the BAT as SYSTEM/admin
$batFile = Join-Path $extractPath "install.bat"

if (Test-Path $batFile) {
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batFile`"" -Verb RunAs -Wait
} else {
    Write-Warning "BAT file not found: $batFile"
}