# Update script from Repository
Write-Output "Download latest script."
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/quwaji/pc-inventory-gas/main/getinfopost.ps1" -OutFile "getinfopost.ps1"

# End script
Write-Output "Done."
Pause
