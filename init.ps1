# Run as Administrator

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$Scripts = @(
    "office.ps1",
    "bloatware.ps1",
    "app_deployment.ps1"
)

foreach ($ScriptName in $Scripts) {

    $Url = "https://raw.githubusercontent.com/AbleTP/abletp/main/$ScriptName"
    $TempPath = Join-Path $env:TEMP $ScriptName

    Write-Host ""
    Write-Host "========================================"
    Write-Host "Downloading: $ScriptName"
    Write-Host "URL: $Url"
    Write-Host "========================================"

    Remove-Item $TempPath -Force -ErrorAction SilentlyContinue

    Invoke-WebRequest -Uri $Url -OutFile $TempPath -UseBasicParsing

    if (!(Test-Path $TempPath)) {
        throw "Download failed: $ScriptName"
    }

    $Size = (Get-Item $TempPath).Length
    Write-Host "Downloaded $ScriptName - Size: $Size bytes"

    if ($Size -lt 50) {
        Get-Content $TempPath
        throw "$ScriptName downloaded incorrectly or is empty."
    }

    Write-Host "Running: $ScriptName"

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $TempPath

    $ExitCode = $LASTEXITCODE
    Write-Host "$ScriptName finished with exit code: $ExitCode"

    if ($ExitCode -ne 0) {
        throw "$ScriptName failed with exit code $ExitCode"
    }

    Remove-Item $TempPath -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "All scripts completed."
