
# Zoom policy settings
$zoomPolicyPath = "HKLM:\SOFTWARE\Policies\Zoom\Zoom Meetings\General"
New-Item -Path $zoomPolicyPath -Force | Out-Null

# Disable auto-update
Set-ItemProperty -Path $zoomPolicyPath -Name "EnableAutoUpdate" -Value 0 -Type DWord

# Enable HD video
Set-ItemProperty -Path $zoomPolicyPath -Name "EnableHDVideo" -Value 1 -Type DWord

# Ensure execution policy allows running scripts
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Define working directory
$workDir = "$env:TEMP"

# Function to download a file
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

# Function to uninstall Zoom Workplace (user-based install)
function Uninstall-ZoomWorkplace {
    $users = Get-ChildItem "C:\Users" -Exclude "Public","Default","Default User","All Users","Administrator"
    foreach ($user in $users) {
        $zoomUninstaller = "C:\Users\$($user.Name)\AppData\Roaming\Zoom\uninstall\Installer.exe"
        if (Test-Path $zoomUninstaller) {
            Write-Output "Uninstalling Zoom Workplace for $($user.Name)..."
            Start-Process -FilePath $zoomUninstaller -ArgumentList "/uninstall" -Wait
        }
    }
}

# Function to get Zoom uninstall string (system-wide install)
function Get-ZoomUninstallString {
    $keys = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($key in $keys) {
        $apps = Get-ItemProperty $key -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Zoom*" }
        foreach ($app in $apps) {
            if ($app.UninstallString) {
                return $app.UninstallString
            }
        }
    }
    return $null
}

# Uninstall Zoom Workplace
Uninstall-ZoomWorkplace

# Uninstall Zoom system-wide
$uninstallString = Get-ZoomUninstallString
if ($uninstallString) {
    Write-Output "Uninstalling Zoom system-wide..."
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallString /quiet" -Wait
    Start-Sleep -Seconds 5
} else {
    Write-Output "Zoom system-wide install not found."
}

# Download and install Zoom
$zoomUrl = "https://zoom.us/client/6.3.11.60501/ZoomInstallerFull.msi"
$zoomMsi = "$workDir\Zoom.msi"
Download-File $zoomUrl $zoomMsi
Write-Output "Installing Zoom..."
Start-Process msiexec.exe -ArgumentList "/i `"$zoomMsi`" /quiet /norestart" -Wait
Write-Output "✅ Zoom installation complete."