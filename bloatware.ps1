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


Stop-Transcript
Write-Output "`n✅ Cleanup complete. Reboot recommended."
exit 0