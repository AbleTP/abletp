# Ensure execution policy allows running scripts
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Download and run VBS script silently using CSCRIPT
$githubPs1Url = "https://github.com/AbleTP/abletp/blob/main/office.ps1"
$tempPs1Path = "$env:TEMP\office.ps1"

Invoke-WebRequest -Uri $githubPs1Url -OutFile $tempPs1Path -UseBasicParsing

Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File", "`"$tempPs1Path`"" -Wait

Remove-Item $tempPs1Path -Force

# Download bloatware.ps1 and run from GitHub
$githubPs1Url = "https://raw.githubusercontent.com/AbleTP/abletp/main/bloatware.ps1"
$tempPs1Path = "$env:TEMP\bloatware.ps1"

Invoke-WebRequest -Uri $githubPs1Url -OutFile $tempPs1Path -UseBasicParsing

Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File", "`"$tempPs1Path`"" -Wait

Remove-Item $tempPs1Path -Force

# Download app_deployment and run from GitHub
$githubPs1Url = "https://raw.githubusercontent.com/AbleTP/abletp/main/app_deployment.ps1"
$tempPs1Path = "$env:TEMP\app_deployment.ps1"

Invoke-WebRequest -Uri $githubPs1Url -OutFile $tempPs1Path -UseBasicParsing

Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File", "`"$tempPs1Path`"" -Wait

Remove-Item $tempPs1Path -Force


