# Update script from Repository
Write-Output "Download latest script."
try {
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/quwaji/pc-inventory-gas/main/getinfopost.ps1" -OutFile "getinfopost.ps1"    
}
catch {
    Write-Output "Download failed. Check internet connection."
}

# End script
Write-Output "Done."
Pause
