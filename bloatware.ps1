<#
.SYNOPSIS
Silent removal of Dell bloatware, Microsoft 365, OneDrive, and OneNote
Run as administrator. Creates a log file at C:\ProgramData\Debloat\Debloat.log
#>

param (
    [string[]]$customwhitelist
)

# Elevate if needed
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
    Write-Host "This script must run as administrator. Relaunching..."
    Start-Process powershell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`" -customwhitelist {1}" -f $PSCommandPath, ($customwhitelist -join ',')) -Verb RunAs
    Exit
}

# No errors throughout
$ErrorActionPreference = 'SilentlyContinue'

# Create log folder
$DebloatFolder = "C:\ProgramData\Debloat"
If (-not (Test-Path $DebloatFolder)) {
    New-Item -Path "$DebloatFolder" -ItemType Directory | Out-Null
}
Start-Transcript -Path "$DebloatFolder\Debloat.log"

# Region: Functions
function Force-UninstallApp {
    param ([string]$displayName)

    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $paths) {
        $apps = Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*$displayName*" }

        foreach ($app in $apps) {
            $name = $app.DisplayName
            $uninstallCmd = $app.UninstallString

            if ($uninstallCmd) {
                Write-Output "Uninstalling: ${name}"
                try {
                    if ($uninstallCmd -match "msiexec") {
                        $silentArgs = $uninstallCmd -replace "msiexec\.exe", ""
                        $silentArgs = $silentArgs -replace "/I", "/x"
                        $silentArgs += " /qn"
                        Start-Process "msiexec.exe" -ArgumentList $silentArgs -Wait -NoNewWindow
                    }
                    else {
                        if ($uninstallCmd -notmatch "/quiet|/qn|/silent|/s") {
                            $uninstallCmd += " /quiet /norestart"
                        }
                        Start-Process "cmd.exe" -ArgumentList "/c", "$uninstallCmd" -Wait -NoNewWindow
                    }
                    Write-Output "Uninstall command executed for: ${name}"
                }
                catch {
                    Write-Warning "Failed to uninstall ${name}: $_"
                }
            }
        }
    }
}

function Remove-AppxEverywhere {
    param([string]$partialName)

    $appx = Get-AppxPackage -AllUsers | Where-Object { $_.Name -like "*$partialName*" }
    foreach ($pkg in $appx) {
        try {
            Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
            Write-Output "Removed Appx: $($pkg.Name)"
        }
        catch {
            Write-Warning "Failed to remove Appx $($pkg.Name): $_"
        }
    }

    $provisioned = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "*$partialName*" }
    foreach ($prov in $provisioned) {
        try {
            Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop
            Write-Output "Removed provisioned Appx: $($prov.DisplayName)"
        }
        catch {
            Write-Warning "Failed to remove provisioned Appx $($prov.DisplayName): $_"
        }
    }
}
# EndRegion

# Locale settings
$locale = (Get-WinSystemLocale).Name
switch ($locale) {
    default {
        $everyone = "Everyone"
        $builtin = "Builtin"
    }
}

# Manufacturer check
$manufacturer = (Get-WmiObject -Class Win32_ComputerSystem).Manufacturer
if ($manufacturer -like "*Dell*") {
    Write-Output "Dell system detected..."

    $UninstallPrograms = @(
        "Dell Command | Update for Windows Universal",
        "Dell SupportAssist OS Recovery",
        "Dell SupportAssist",
        "DellInc.DellSupportAssistforPCs",
        "Dell SupportAssist Remediation",
        "SupportAssist Recovery Assistant",
        "Dell SupportAssist OS Recovery Plugin for Dell Update",
        "Dell SupportAssistAgent",
        "Dell Update - SupportAssist Update Plugin"
        "Dell Pair"
    )

    $WhitelistedApps = @(
        "Dell Core Services", "Dell Trusted Device", "Dell Watchdog Timer"
    )

    # Apply custom whitelist
    if ($customwhitelist) {
        $WhitelistedApps += $customwhitelist
    }

    $UninstallPrograms = $UninstallPrograms | Where-Object { $WhitelistedApps -notcontains $_ }

    foreach ($app in $UninstallPrograms) {
        Remove-AppxEverywhere $app
        Force-UninstallApp $app
    }

    # Belt & suspenders uninstall
    foreach ($program in $UninstallPrograms) {
        Get-CimInstance -ClassName Win32_Product | Where-Object Name -Match $program | Invoke-CimMethod -MethodName Uninstall
    }

    # Specific uninstalls via QuietUninstallString
    foreach ($target in @(
            "Dell SupportAssist Remediation",
            "Dell SupportAssist OS Recovery Plugin for Dell Update"
        )) {
        $entries = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, `
            HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
        Get-ItemProperty | Where-Object { $_.DisplayName -match $target }

        foreach ($entry in $entries) {
            if ($entry.QuietUninstallString) {
                try {
                    cmd.exe /c $entry.QuietUninstallString
                    Write-Output "Executed quiet uninstall for $target"
                }
                catch {
                    Write-Warning "Failed to uninstall $target"
                }
            }
        }
    }

    # Dell Optimizer special uninstall
    $dellOptimizerPath = "C:\Program Files (x86)\InstallShield Installation Information\{CC40119D-6ADF-4832-8025-4808195E41D5}\DellOptimizer.exe"
    if (Test-Path $dellOptimizerPath) {
        Start-Process $dellOptimizerPath -ArgumentList "-remove", "-runfromtemp", "-silent" -Wait -NoNewWindow
        Write-Output "Dell Optimizer removed silently"
    }
}

# OneDrive uninstall
$oneDrivePath = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
if (Test-Path $oneDrivePath) {
    Start-Process $oneDrivePath "/uninstall" -NoNewWindow -Wait
    Write-Output "OneDrive uninstalled."
}

# OneNote Appx removal
Remove-AppxEverywhere "OneNote"

# Must run as Administrator
# Set system time zone to Eastern Standard Time
Write-Output "Setting time zone to Eastern Standard Time..."
Set-TimeZone -Id "Eastern Standard Time"

# Optionally resync with time server
Write-Output "Forcing time sync with NTP server..."
w32tm /resync /force

# Confirm result
Write-Output "✅ Time zone and current time updated."
Get-TimeZone
Get-Date

# Get Service Tag (Dell) or Serial Number (fallback)
$serviceTag = (Get-WmiObject -Class Win32_BIOS).SerialNumber.Trim()

# Get current computer name
$currentName = $env:COMPUTERNAME

# Check if renaming is necessary
if ($currentName -ne $serviceTag) {
    Write-Output "Renaming computer from $currentName to $serviceTag..."
    
    try {
        Rename-Computer -NewName $serviceTag -Force -PassThru
        Write-Output "✅ Rename successful. A reboot is required."
    } catch {
        Write-Error "❌ Rename failed: $_"
    }
} else {
    Write-Output "Computer name is already set to the Service Tag ($serviceTag). No action taken."
}


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

Stop-Transcript
Write-Output "`n✅ Cleanup complete. Reboot recommended."
exit 0