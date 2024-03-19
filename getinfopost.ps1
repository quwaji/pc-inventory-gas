# Check Paramaters

param (
    [string]$room,
    [string]$gid,
    [string]$checkby
)

# Room
if(-not $room) {
    $room = Read-Host "Room: (ex. LRC113) "
}

# GID - Goverment Serial No. 
if(-not $gid) {
    $gid = Read-Host "GID: Goverment Serial No. "
}

# checkby - Checked by 
if(-not $checkby) {
    $checkby = Read-Host "checkby: Name who check this "
}

Write-Host "Room: $room,", "Goverment Serial No. $gid,", "Cheked by $checkby"

# Get system information using Get-ComputerInfo
Write-Output "Getting ComputerInfo..."
$systemInfo = Get-ComputerInfo
# $systemInfo = Get-ComputerInfo | Select-Object CsName, CsManufacturer, CsModel, WindowsProductName, BiosSeralNumber

# Get installed programs using WMIC
# Write-Output "Getting installed programs using WMIC..."
# $installedPrograms = Invoke-Expression "wmic product get Name,Vendor,Version /FORMAT:CSV" | Select-Object -Skip 1 | ForEach-Object {
#     $program = $_.Trim()
#     [PSCustomObject]@{
#         Name = $program.Split(",")[1]
#         Vendor = $program.Split(",")[2]
#         Version = $program.Split(",")[3]
#     }
# }

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
$regKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
$proxysetting = Get-ItemProperty -Path $regKey ProxyEnable,ProxyServer
$proxyenable = $proxysetting.ProxyEnable
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

# Combine system info and installed programs into a single object
$combinedData = @{
    ROOM = $room
    PC_NAME = $systemInfo.CsName
    MAKER_MODEL = $systemInfo.CsManufacturer, $systemInfo.CsModel -join " "
    GID = $gid
    SERVICE_TAG = $systemInfo.BiosSeralNumber
    CHECK_BY = $checkby
    PROXY_ENABLE = $proxyenable
    PROXY_SERVER = $proxyserver
    WINDOWS_PRODUCT_NAME = $systemInfo.WindowsProductName
    SystemInfo = $systemInfo
    # InstalledPrograms = $installedPrograms
    App32 = $app32
    App64 = $app64
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
