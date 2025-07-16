# Must be run as Administrator

# Get Service Tag (Dell) or Serial Number (fallback)
$serviceTag = (Get-WmiObject -Class Win32_BIOS).SerialNumber.Trim()

# Get current computer name
$currentName = $env:COMPUTERNAME

# Check if renaming is necessary
if ($currentName -ne $serviceTag) {
    Write-Output "Renaming computer from $currentName to $serviceTag..."
    
    try {
        Rename-Computer -NewName $serviceTag -Force -PassThru
        Write-Output "✅ Rename successful. A reboot is required."
    } catch {
        Write-Error "❌ Rename failed: $_"
    }
} else {
    Write-Output "Computer name is already set to the Service Tag ($serviceTag). No action taken."
}
