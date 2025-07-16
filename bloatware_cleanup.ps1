# Ensure execution policy allows running scripts
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Download bloatware.ps1 and run from GitHub
$githubPs1Url = "https://raw.githubusercontent.com/demo7up/abletp/refs/heads/main/bloatware.ps1"
$tempPs1Path = "$env:TEMP\bloatware.ps1"

Invoke-WebRequest -Uri $githubPs1Url -OutFile $tempPs1Path

# Run the PowerShell script
Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tempPs1Path`"" -Wait

Remove-Item $tempPs1Path

# Download and run C2R VBS from GitHub
$githubVbsUrl = "https://raw.githubusercontent.com/demo7up/abletp/refs/heads/main/remove_c2r.vbs"
$tempVbsPath = "$env:TEMP\remote_script.vbs"

Invoke-WebRequest -Uri $githubVbsUrl -OutFile $tempVbsPath
Start-Process "wscript.exe" -ArgumentList "`"$tempVbsPath`"" -Wait
Remove-Item $tempVbsPath