# Download and run C2R VBS from GitHub
$githubVbsUrl = "https://raw.githubusercontent.com/demo7up/abletp/refs/heads/main/remove_c2r.vbs"
$tempVbsPath = "$env:TEMP\remote_script.vbs"

Invoke-WebRequest -Uri $githubVbsUrl -OutFile $tempVbsPath
Start-Process "wscript.exe" -ArgumentList "`"$tempVbsPath`"" -Wait
Remove-Item $tempVbsPath


# Download and run C2R VBS from GitHub
$githubVbsUrl = "https://raw.githubusercontent.com/demo7up/abletp/refs/heads/main/remove_c2r.vbs"
$tempVbsPath = "$env:TEMP\remote_script.vbs"

Invoke-WebRequest -Uri $githubVbsUrl -OutFile $tempVbsPath
Start-Process "wscript.exe" -ArgumentList "`"$tempVbsPath`"" -Wait
Remove-Item $tempVbsPath