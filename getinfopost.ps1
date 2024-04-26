# Check Paramaters

param (
    [string]$room,
    [string]$gidcp,
    [string]$gidmn,
    [string]$msnmn,
    [string]$note,
    [string]$checkby
)

# Room
if(-not $room) {
    $room = Read-Host "Room: (ex. LRC113) "
}

# GID - Goverment Serial No. for Computer
if(-not $gidcp) {
    $gidcp = Read-Host "GID_CP: Goverment Serial No. for Computer"
}

# GID - Goverment Serial No. for Monitor 
if(-not $gidmn) {
    $gidmn = Read-Host "GID_MN: Goverment Serial No. for Monitor"
}

# Manufactures Service Tag(S/N). for Monitor 
if(-not $msnmn) {
    $msnmn = Read-Host "SERIAL_MN: Manufactures Service Tag(S/N). for Monitor"
}

# note
if(-not $note) {
    $note = Read-Host "NOTE: Comment at this moment"
}

# checkby - Checked by 
if(-not $checkby) {
    $checkby = Read-Host "checkby: Name who check this "
}

Write-Host "Room: $room,", "Goverment Serial No. $gidcp (Computer),", "Goverment Serial No. $gidcp (Monitor),", "Manufacture Service Tag(S/N). $msnmn (Monitor),", "Note: $note", "Cheked by $checkby"

# Get system information using Get-ComputerInfo
Write-Output "Getting ComputerInfo..."
$systemInfo = Get-ComputerInfo
# $systemInfo = Get-ComputerInfo | Select-Object CsName, CsManufacturer, CsModel, WindowsProductName, BiosSeralNumber

# Get installed programs from Registory
Write-Output "Getting installed programs using Registory..."

## List 32bit installed programs 
$regPath = 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' 
$app32 = Get-ItemProperty -Path $regPath | 
Select-Object DisplayName, DisplayVersion, Publisher, InstallDate

## List 64bit installed programs
$regPath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
$app64 = Get-ItemProperty -Path $regPath |
Select-Object DisplayName, DisplayVersion, Publisher, InstallDate

## AntiVirus programs
$antiVirusInfo = Get-ItemProperty -Path $regPath | Where-Object {$_.DisplayName -like 'Symantec Endpoint Protection'} | Select DisplayName, DisplayVersion, Publisher, InstallDate
$antiVirus = $antiVirusInfo.DisplayName, $antiVirusInfo.DisplayVersion -join " "
Write-Output "AntiVirus: $antiVirus"

# Get/Set Proxy Setting
Write-Output "Getting ProxySetting..."
$regKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
$proxysetting = Get-ItemProperty -Path $regKey MigrateProxy
$migrateproxy = $proxysetting.MigrateProxy
if ($proxysetting.MigrateProxy) {
    Write-Output "Proxy Settings Auto-Detect is enable."
}
else {
    Write-Output "Proxy Settings Auto-Detect is enable."
    # Set MigrateProxy
    Set-ItemProperty -Path $regKey MigrateProxy -Value 1 -ErrorAction Stop
    Write-Output "Proxy Settings Auto-Detect is now Enabled. "
}

# Analyze Network
$gdn_ip = (Get-NetIPAddress | Where-Object {$_.AddressFamily -eq "IPv4"} | Where-Object {$_.IPAddress -like "10.97.*"}).IPAddress
if ($gdn_ip) {
    Write-Output "Setting Proxy for GDN..."

    $regKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    $proxysetting = Get-ItemProperty -Path $regKey ProxyEnable,ProxyServer
    $proxyserver = $proxysetting.ProxyServer
    if ($proxysetting.ProxyEnable) {
        Write-Output "Proxy is enable. Server: $proxyserver"
    }
    else {
        if ($proxysetting.ProxyServer = "") {
            Write-Output "Proxy is not setting."
        }
        else {
            Write-Output "Proxy is disable. Server: $proxyserver"
        }
        # Set Proxy
        Set-ItemProperty -Path $regKey ProxyEnable -Value 1 -ErrorAction Stop
        Set-ItemProperty -Path $regKey ProxyServer -Value "10.97.114.6:3128" -ErrorAction Stop
        Write-Output "Proxy is now Enabled. "
    }
}

# Combine system info and installed programs into a single object
$combinedData = @{
    ROOM = $room
    PC_NAME = $systemInfo.CsName
    MAKER_MODEL = $systemInfo.CsManufacturer, $systemInfo.CsModel -join " "
    GID_CP = $gidcp
    GID_MN = $gidmn
    SERVICE_TAG = $systemInfo.BiosSeralNumber
    SERIAL_MN = $msnmn
    NOTE = $note
    CHECK_BY = $checkby
    MIGRATE_PROXY = $migrateproxy
    WINDOWS_PRODUCT_NAME = $systemInfo.WindowsProductName
    _SystemInfo = $systemInfo
    # InstalledPrograms = $installedPrograms
    _App32 = $app32
    _App64 = $app64
    ANTI_VIRUS = $antiVirus
}

# Convert to JSON
$jsonData = $combinedData | ConvertTo-Json

# Specify the URL where you want to post the data
$targetUrl = "https://script.google.com/macros/s/AKfycby6f4gdDgIhM7QJfpyyo4Ybys9WFEuBzYjTKLFBcCfnC4e_5_n9JPxjN8TdOXmtgAKa/exec"

# Send the JSON data to the URL (you can use Invoke-RestMethod or any other method)
Write-Output "Sending data to Spreadsheet..."
Invoke-RestMethod -Uri $targetUrl -Method Post -Body $jsonData -ContentType "application/json"

# Write to local file
$outfilename = $systemInfo.CSName, "json" -join "."
$jsonData | Out-File -FilePath $outfilename
Write-Output "Data save to file $outfilename"

# End script
Pause
