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
    Invoke-WebRequest -Uri $url -OutFile $destination -UseBasicParsing
}

# 1. Install Adobe Reader DC
$adobeUrl = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2300820415/AcroRdrDC2300820415_en_US.exe"
$adobeExe = "$workDir\AdobeReader.exe"
Download-File $adobeUrl $adobeExe
Start-Process $adobeExe -ArgumentList "/sAll /rs /rps /msi EULA_ACCEPT=YES" -Wait

# 2. Install Google Chrome
$chromeUrl = "https://dl.google.com/chrome/install/375.126/chrome_installer.exe"
$chromeExe = "$workDir\Chrome.exe"
Download-File $chromeUrl $chromeExe
Start-Process $chromeExe -ArgumentList "/silent /install" -Wait

# 3. Install Zoom (per-machine)
$zoomUrl = "https://zoom.us/client/latest/ZoomInstallerFull.msi"
$zoomMsi = "$workDir\Zoom.msi"
Download-File $zoomUrl $zoomMsi
Start-Process msiexec.exe -ArgumentList "/i `"$zoomMsi`" /quiet /norestart" -Wait

# # 4. Install Teams (per-machine)
# $teamsUrl = "https://statics.teams.cdn.office.net/production-windows-x64/1.00.6774.0/Teams_windows_x64.msi"
# $teamsMsi = "$workDir\Teams.msi"
# Download-File $teamsUrl $teamsMsi
# Start-Process msiexec.exe -ArgumentList "/i `"$teamsMsi`" ALLUSERS=1 /quiet /norestart" -Wait

# # 5. Install Office 365 Business (Microsoft 365 Apps for Business)
# # Download Office Deployment Tool
# $odtUrl = "https://download.microsoft.com/download/6/D/4/6D48F0C5-401C-43B2-B3C6-BB46B9CE5A4C/OfficeDeploymentTool.exe"
# $odtExe = "$workDir\ODT.exe"
# Download-File $odtUrl $odtExe
# Start-Process $odtExe -ArgumentList "/quiet /extract:$workDir" -Wait

# # Create Office configuration XML
# $officeXml = @"
# <Configuration>
#   <Add OfficeClientEdition="64" Channel="MonthlyEnterprise">
#     <Product ID="O365BusinessRetail">
#       <Language ID="en-us" />
#     </Product>
#   </Add>
#   <Display Level="None" AcceptEULA="TRUE" />
#   <Property Name="AUTOACTIVATE" Value="1" />
# </Configuration>
# "@
# $officeConfig = "$workDir\OfficeConfig.xml"
# $officeXml | Out-File -Encoding UTF8 -FilePath $officeConfig

# # Run Office install
# Start-Process "$workDir\setup.exe" -ArgumentList "/configure $officeConfig" -Wait

Write-Host "✅ All applications installed. Reboot may be required."
