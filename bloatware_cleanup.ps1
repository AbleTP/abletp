# Ensure execution policy allows running scripts
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Download and run VBS script silently using CSCRIPT
$githubVbsUrl = "https://raw.githubusercontent.com/demo7up/abletp/main/remove_c2r.vbs"
$tempVbsPath = "$env:TEMP\remove_c2r.vbs"

Invoke-WebRequest -Uri $githubVbsUrl -OutFile $tempVbsPath -UseBasicParsing

$output = & cscript.exe //Nologo "$tempVbsPath"
$output | Out-File "$env:ProgramData\Debloat\vbs_output.log" -Encoding UTF8 -Force
Write-Output "VBS output saved to: $env:ProgramData\Debloat\vbs_output.log"

Remove-Item $tempVbsPath -Force

# Download bloatware.ps1 and run from GitHub
$githubPs1Url = "https://raw.githubusercontent.com/demo7up/abletp/main/bloatware.ps1"
$tempPs1Path = "$env:TEMP\bloatware.ps1"

Invoke-WebRequest -Uri $githubPs1Url -OutFile $tempPs1Path -UseBasicParsing

Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File", "`"$tempPs1Path`"" -Wait

Remove-Item $tempPs1Path -Force

# Download device_rename.ps1 and run from GitHub
$githubPs1Url = "https://raw.githubusercontent.com/demo7up/abletp/main/app_deployment.ps1"
$tempPs1Path = "$env:TEMP\app_deployment.ps1"

Invoke-WebRequest -Uri $githubPs1Url -OutFile $tempPs1Path -UseBasicParsing

Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File", "`"$tempPs1Path`"" -Wait

Remove-Item $tempPs1Path -Force


