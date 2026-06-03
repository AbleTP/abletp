$ErrorActionPreference = "SilentlyContinue"

Write-Host "Closing Office apps..."
$apps = @(
    "winword","excel","powerpnt","outlook","onenote",
    "msaccess","publisher","visio","winproj",
    "officeclicktorun","officec2rclient"
)

foreach ($app in $apps) {
    Get-Process $app -ErrorAction SilentlyContinue | Stop-Process -Force
}

Write-Host "Stopping Office services..."
Stop-Service ClickToRunSvc -Force -ErrorAction SilentlyContinue
Stop-Service OfficeSvc -Force -ErrorAction SilentlyContinue

Start-Sleep -Seconds 3

Write-Host "Detecting Office 2016 / 2019 products..."

$uninstallKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$officeApps = Get-ItemProperty $uninstallKeys |
Where-Object {
    $_.DisplayName -match "Microsoft Office|Office 16|Office 2016|Office 2019|Visio|Project|Access Runtime" -or
    $_.DisplayName -match "Microsoft 365|Office 365"
} |
Sort-Object DisplayName -Unique

foreach ($app in $officeApps) {

    Write-Host "Found: $($app.DisplayName)"

    $uninstall = $app.UninstallString

    if ([string]::IsNullOrWhiteSpace($uninstall)) {
        Write-Host "No uninstall string found, skipping."
        continue
    }

    # MSI-based Office
    if ($uninstall -match "msiexec") {
        $guid = ($uninstall -replace ".*?({[A-Fa-f0-9-]+}).*", '$1')

        if ($guid -match "^{[A-Fa-f0-9-]+}$") {
            Write-Host "Uninstalling MSI product: $guid"
            Start-Process "msiexec.exe" -ArgumentList "/x $guid /qn /norestart" -Wait
        }

        continue
    }

    # Click-to-Run Office
    if ($uninstall -match "OfficeClickToRun.exe") {

        Write-Host "Uninstalling Click-to-Run product..."

        if ($uninstall -match '^"([^"]+)"\s*(.*)$') {
            $exe  = $matches[1]
            $args = $matches[2]
        } else {
            $parts = $uninstall -split "\s+", 2
            $exe  = $parts[0]
            $args = $parts[1]
        }

        if (Test-Path $exe) {
            Start-Process $exe -ArgumentList $args -Wait
        }

        continue
    }
}

Write-Host "Running explicit Office 2016 / 2019 Click-to-Run removals..."

$c2rPaths = @(
    "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe",
    "$env:ProgramFiles(x86)\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
)

$productIds = @(
    "ProPlusRetail",
    "ProfessionalRetail",
    "HomeBusinessRetail",
    "HomeStudentRetail",
    "StandardRetail",
    "O365ProPlusRetail",
    "VisioProRetail",
    "VisioStdRetail",
    "ProjectProRetail",
    "ProjectStdRetail"
)

foreach ($c2r in $c2rPaths) {
    if (Test-Path $c2r) {
        foreach ($product in $productIds) {
            Write-Host "Attempting C2R removal: $product"
            Start-Process $c2r -ArgumentList "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=$product.16_en-us_x-none culture=en-us version.16=16.0 displaylevel=false forceappshutdown=true" -Wait
        }
    }
}

Write-Host "Stopping services again..."
Stop-Service ClickToRunSvc -Force -ErrorAction SilentlyContinue
Stop-Service OfficeSvc -Force -ErrorAction SilentlyContinue

Write-Host "Removing scheduled tasks..."
Get-ScheduledTask | Where-Object {
    $_.TaskName -match "Office|ClickToRun"
} | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue

Write-Host "Removing leftover folders..."

$folders = @(
    "$env:ProgramFiles\Microsoft Office",
    "$env:ProgramFiles(x86)\Microsoft Office",
    "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun",
    "$env:ProgramFiles(x86)\Common Files\Microsoft Shared\ClickToRun",
    "$env:ProgramFiles\Common Files\Microsoft Shared\OFFICE16",
    "$env:ProgramFiles(x86)\Common Files\Microsoft Shared\OFFICE16",
    "$env:ProgramData\Microsoft\Office",
    "$env:ProgramData\Microsoft\ClickToRun",
    "$env:LOCALAPPDATA\Microsoft\Office",
    "$env:APPDATA\Microsoft\Office"
)

foreach ($folder in $folders) {
    if (Test-Path $folder) {
        Write-Host "Deleting: $folder"
        takeown /F $folder /R /D Y | Out-Null
        icacls $folder /grant Administrators:F /T | Out-Null
        Remove-Item $folder -Recurse -Force
    }
}

Write-Host "Removing Office registry leftovers..."

$regKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Office",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office",
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun",
    "HKCU:\Software\Microsoft\Office"
)

foreach ($key in $regKeys) {
    if (Test-Path $key) {
        Write-Host "Deleting registry key: $key"
        Remove-Item $key -Recurse -Force
    }
}

Write-Host ""
Write-Host "Office 2016 / 2019 removal cleanup complete."
Write-Host "REBOOT REQUIRED before reinstalling Office/Microsoft 365."
Write-Host ""

$ErrorActionPreference = 'SilentlyContinue'

Write-Host "Stopping Microsoft processes..."

Get-Process olk,outlook,teams,msedgewebview2,OneDrive -ErrorAction SilentlyContinue |
    Stop-Process -Force

Write-Host "Removing New Outlook for all users..."

Get-AppxPackage -AllUsers Microsoft.OutlookForWindows |
    Remove-AppxPackage -AllUsers

Write-Host "Stopping TokenBroker..."

Stop-Service TokenBroker -Force -ErrorAction SilentlyContinue

Write-Host "Leaving Workplace Join..."

dsregcmd /leave

$Profiles = Get-ChildItem "C:\Users" -Directory |
Where-Object {
    $_.Name -notin @(
        'Public',
        'Default',
        'Default User',
        'All Users',
        'defaultuser0'
    )
}

foreach ($Profile in $Profiles) {

    Write-Host ""
    Write-Host "Processing profile: $($Profile.Name)" -ForegroundColor Cyan

    $ProfilePath = $Profile.FullName

    # -----------------------------------------------------
    # File cleanup
    # -----------------------------------------------------

    $Paths = @(
        "$ProfilePath\AppData\Local\Microsoft\Olk",
        "$ProfilePath\AppData\Local\Microsoft\OneAuth",
        "$ProfilePath\AppData\Local\Microsoft\IdentityCache",
        "$ProfilePath\AppData\Local\Microsoft\TokenBroker",
        "$ProfilePath\AppData\Local\Packages\Microsoft.OutlookForWindows_8wekyb3d8bbwe",
        "$ProfilePath\AppData\Local\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy",
        "$ProfilePath\AppData\Local\Packages\Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy",
        "$ProfilePath\AppData\Roaming\Microsoft\Outlook",
        "$ProfilePath\AppData\Roaming\Microsoft\Office",
        "$ProfilePath\AppData\Local\Microsoft\Office"
    )

    foreach ($Path in $Paths) {

        if (Test-Path $Path) {

            Write-Host "Removing: $Path"

            takeown /F $Path /R /D Y | Out-Null
            icacls $Path /grant Administrators:F /T | Out-Null

            Remove-Item $Path -Recurse -Force
        }
    }

    $NtUser = Join-Path $ProfilePath "NTUSER.DAT"

    if (Test-Path $NtUser) {

        $HiveName = "TempHive_$($Profile.Name)"

        reg load "HKU\$HiveName" $NtUser | Out-Null

        $RegKeys = @(
            "Registry::HKEY_USERS\$HiveName\Software\Microsoft\Office",
            "Registry::HKEY_USERS\$HiveName\Software\Microsoft\OneAuth",
            "Registry::HKEY_USERS\$HiveName\Software\Microsoft\IdentityCRL",
            "Registry::HKEY_USERS\$HiveName\Software\Microsoft\AADBrokerPlugin",
            "Registry::HKEY_USERS\$HiveName\Software\Microsoft\Windows\CurrentVersion\AAD",
            "Registry::HKEY_USERS\$HiveName\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin"
        )

        foreach ($Key in $RegKeys) {

            if (Test-Path $Key) {
                Write-Host "Removing Registry: $Key"
                Remove-Item $Key -Recurse -Force
            }
        }

        reg unload "HKU\$HiveName" | Out-Null
    }
}

Write-Host ""
Write-Host "Removing Microsoft credentials..."

cmdkey /list |
ForEach-Object {

    if ($_ -match "Target: (.+)") {

        $Target = $Matches[1]

        if ($Target -match "Microsoft|Office|ADAL|Azure|MSOID|OneAuth|Outlook|Teams") {

            Write-Host "Deleting credential: $Target"

            cmdkey /delete:$Target | Out-Null
        }
    }
}

# Download and install latest Microsoft 365 Apps for Enterprise

$TempFolder = "C:\Temp\OfficeInstall"
New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null

Invoke-WebRequest `
    -Uri "https://officecdn.microsoft.com/pr/wsus/setup.exe" `
    -OutFile "$TempFolder\setup.exe"

@'
<Configuration>
    <Add OfficeClientEdition="64" Channel="Current">
        <Product ID="O365ProPlusRetail">
            <Language ID="en-us" />
        </Product>
    </Add>
    <Display Level="Full" AcceptEULA="TRUE" />
    <Updates Enabled="TRUE" />
</Configuration>
'@ | Set-Content "$TempFolder\configuration.xml"

Start-Process `
    -FilePath "$TempFolder\setup.exe" `
    -ArgumentList "/configure configuration.xml" `
    -WorkingDirectory $TempFolder `
    -Wait

Write-Host "Microsoft 365 Apps installation completed."
Write-Host ""
Write-Host "COMPLETE"
Write-Host "REBOOTING"

Restart-Computer -Force
