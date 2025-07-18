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


# Function to get the latest version and download URL of Adobe Acrobat Reader DC
function Get-AdobeAcrobatReaderDCUrls {
    [CmdletBinding()]
    param ()

    # URL of the Adobe Acrobat Reader DC release notes page
    $apiUrl = 'https://helpx.adobe.com/acrobat/release-note/release-notes-acrobat-reader.html'
    Write-Debug "Fetching main release notes page: $apiUrl"

    try {
        # Fetch the main release notes page using curl.exe
        $response = curl.exe -s $apiUrl
        if ($response) {
            $htmlContent = $response
            Write-Debug "Main release notes page content fetched."
        } else {
            throw "Failed to fetch main release notes page."
        }
    } catch {
        # Handle errors in fetching the main release notes page
        Write-Debug "Error fetching main release notes page: $_"
        Write-Output "Error fetching main release notes page: $_"
        exit
    }

    # Extract the first <a> link that matches the specified pattern
    $linkPattern = [regex]::new('<a href="(https://www\.adobe\.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/[^"]+)"[^>]*>(DC [^<]+)</a>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $linkMatch = $linkPattern.Match($htmlContent)
    Write-Debug "Searching for the first release notes link..."

    if ($linkMatch.Success) {
        # Extract the release notes URL and version from the matched link
        $releaseNotesUrl = $linkMatch.Groups[1].Value
        $version = $linkMatch.Groups[2].Value
        Write-Debug "Release Notes URL: $releaseNotesUrl"
        Write-Debug "Version: $version"

        # Fetch the release notes page to get the .msp file link
        Write-Debug "Fetching release notes page: $releaseNotesUrl"
        try {
            $releaseNotesResponse = curl.exe -s $releaseNotesUrl
            if ($releaseNotesResponse) {
                $releaseNotesContent = $releaseNotesResponse
                Write-Debug "Release notes page content fetched."
            } else {
                throw "Failed to fetch release notes page."
            }
        } catch {
            # Handle errors in fetching the release notes page
            Write-Debug "Error fetching release notes page: $_"
            Write-Output "Error fetching release notes page: $_"
            exit
        }

        # Find the .msp file link in the release notes page
        $mspLinkPattern = [regex]::new('<a[^>]+href="([^"]+\.msp)"[^>]*>([^<]+)</a>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $mspLinkMatch = $mspLinkPattern.Match($releaseNotesContent)
        Write-Debug "Searching for the .msp file link..."

        if ($mspLinkMatch.Success) {
            # Extract the .msp file URL and version
            $mspUrl = $mspLinkMatch.Groups[1].Value
            Write-Debug "MSP URL: $mspUrl"
            $mspFileName = [System.IO.Path]::GetFileNameWithoutExtension($mspUrl)
            $mspVersion = $mspFileName -replace '.*?(\d{4,}).*', '$1'
            Write-Debug "Extracted MSP Version: $mspVersion"

            # Construct the download URLs for the MUI installer and MSP update files
            $MUIurl = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/$mspVersion/AcroRdrDC${mspVersion}_MUI.exe"
            Write-Debug "MUI URL: $MUIurl"

            $MUIurl64 = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/$mspVersion/AcroRdrDCx64${mspVersion}_MUI.exe"
            Write-Debug "MUI URL 64-bit: $MUIurl64"

            $MUImspURL = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/$mspVersion/AcroRdrDCUpd${mspVersion}_MUI.msp"
            Write-Debug "MUI MSP URL: $MUImspURL"

            $MUImspURL64 = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/$mspVersion/AcroRdrDCx64Upd${mspVersion}_MUI.msp"
            Write-Debug "MUI MSP URL 64-bit: $MUImspURL64"

            # Write-Output [PSCustomObject]@{
            #     Version         = $version
            #     ReleaseNotesUrl = $releaseNotesUrl
            #     MUIurl          = $MUIurl
            #     MUIurl64        = $MUIurl64
            #     MUImspURL       = $MUImspURL
            #     MUImspURL64     = $MUImspURL64
            # }
            # Return the extracted information as a PowerShell custom object
            return $MUIurl64
            
        } else {
            # Handle cases where the .msp file link is not found
            Write-Debug "MSP file link not found."
            Write-Output "MSP file link not found."
            exit
        }
    } else {
        # Handle cases where the version link is not found
        Write-Debug "Version link not found."
        Write-Output "Version link not found."
        exit
    }
}

# Example usage
$latest = Get-AdobeAcrobatReaderDCUrls

# Write the latest version and URLs
$latest

# 1. Adobe Reader DC
if (-not (Is-AppInstalled "Adobe Acrobat Reader")) {
    Write-Output "Installing Adobe Reader..."
    $adobeUrl = $latest
    $adobeExe = "$workDir\AdobeReader.exe"
    Download-File $adobeUrl $adobeExe
    Start-Process $adobeExe -ArgumentList "/sAll /rs /rps /msi EULA_ACCEPT=YES /quiet" -Wait
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


#Pin Outlook to taskbar
$OutlookPath = "$env:ProgramFiles\Microsoft Office\root\Office16\OUTLOOK.EXE"
$Shell = New-Object -ComObject Shell.Application
$Folder = $Shell.Namespace((Split-Path $OutlookPath))
$Item = $Folder.ParseName((Split-Path $OutlookPath -Leaf))
$Item.InvokeVerb("Pin to Tas&kbar")
Write-Host "`nOutlook pinned to taskbar."

Write-Host "`n✅ Finished. Reboot may be required."
