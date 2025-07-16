# Ensure execution policy allows running scripts
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Download bloatware.ps1 and run from GitHub
$githubPs1Url = "https://raw.githubusercontent.com/demo7up/abletp/main/bloatware.ps1"
$tempPs1Path = "$env:TEMP\bloatware.ps1"

Invoke-WebRequest -Uri $githubPs1Url -OutFile $tempPs1Path -UseBasicParsing

Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File", "`"$tempPs1Path`"" -Wait

Remove-Item $tempPs1Path -Force

# Download and run VBS script silently using CSCRIPT
$githubVbsUrl = "https://raw.githubusercontent.com/demo7up/abletp/main/remove_c2r.vbs"
$tempVbsPath = "$env:TEMP\remove_c2r.vbs"

Invoke-WebRequest -Uri $githubVbsUrl -OutFile $tempVbsPath -UseBasicParsing

$logPath = "$env:ProgramData\Debloat\vbs_output.log"

# Run VBS script and capture live output line by line
$cscript = Start-Process -FilePath "cscript.exe" `
    -ArgumentList "//Nologo", "`"$tempVbsPath`"" `
    -NoNewWindow -PassThru -RedirectStandardOutput "$logPath" -Wait

# Display output in PowerShell console after run
Get-Content $logPath | ForEach-Object { Write-Output $_ }

Write-Output "`nVBS output saved to: $logPath"

Remove-Item $tempVbsPath -Force

# Download device_rename.ps1 and run from GitHub
$githubPs1Url = "https://raw.githubusercontent.com/demo7up/abletp/main/device_rename.ps1"
$tempPs1Path = "$env:TEMP\device_rename.ps1"

Invoke-WebRequest -Uri $githubPs1Url -OutFile $tempPs1Path -UseBasicParsing

Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File", "`"$tempPs1Path`"" -Wait

Remove-Item $tempPs1Path -Force