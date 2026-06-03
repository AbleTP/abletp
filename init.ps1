# Run as Administrator

$ErrorActionPreference = "Stop"

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$Scripts = @(
    @{
        Name = "office.ps1"
        Url  = "https://raw.githubusercontent.com/AbleTP/abletp/main/office.ps1"
    },
    @{
        Name = "bloatware.ps1"
        Url  = "https://raw.githubusercontent.com/AbleTP/abletp/main/bloatware.ps1"
    },
    @{
        Name = "app_deployment.ps1"
        Url  = "https://raw.githubusercontent.com/AbleTP/abletp/main/app_deployment.ps1"
    }
)

foreach ($Script in $Scripts) {

    $TempPath = Join-Path $env:TEMP $Script.Name

    Write-Host ""
    Write-Host "Downloading $($Script.Name)..." -ForegroundColor Cyan

    Invoke-WebRequest `
        -Uri $Script.Url `
        -OutFile $TempPath `
        -UseBasicParsing

    if (!(Test-Path $TempPath)) {
        throw "Failed to download $($Script.Name)"
    }

    Write-Host "Running $($Script.Name)..." -ForegroundColor Green

    $Process = Start-Process powershell.exe `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$TempPath`"" `
        -Wait `
        -PassThru

    Write-Host "$($Script.Name) exited with code $($Process.ExitCode)"

    if ($Process.ExitCode -ne 0) {
        throw "$($Script.Name) failed with exit code $($Process.ExitCode)"
    }

    Remove-Item $TempPath -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "All scripts completed successfully." -ForegroundColor Green
