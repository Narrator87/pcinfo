<#
    .SYNOPSIS
    PCInfo (Client-side)
    Version: 0.14 26.03.2021

     © Anton Kosenko mail:Anton.Kosenko@gmail.com
    Licensed under the Apache License, Version 2.0

    .DESCRIPTION
    This script using wmi requests get information about workstations and write it in file by json format 
#>

# requires -version 2

# Log function
    $Date = Get-Date -Format "yyyyMMdd"
    $Logfile = ".\$Date.log"
    Function LogWrite
    {
        Param ([string]$logstring)
        Add-content $Logfile -value $logstring
    }
<#
      Processing array function
      Array values are given in here-string. From the first to next-to-last the elements of the array, characters are added "," and in the last element only "".
      I got part of code for this function from this page:
      https://stackoverflow.com/questions/1785474/get-index-of-current-item-in-powershell-loop 
#>   
function SetArraytoString {
    param (
        [parameter(Mandatory=$true)]
        [array]
        $ArrayName
        )
    if (($ArrayName).count -gt 1) 
       {
        0..(($ArrayName).count-2) | ForEach-Object { '"{0}",' -f $ArrayName[$_]} -ErrorAction SilentlyContinue
        (($ArrayName).count-1) | ForEach-Object { '"{0}"' -f $ArrayName[$_]} -ErrorAction SilentlyContinue
    }
    else {
        '"{0}"' -f $ArrayName
    }
  }
# Function for creating string by json format from values array
# Here a good article about json-array: https://www.w3schools.com/js/js_json_arrays.asp
    function SetJsonArray {
        param (
            [parameter(Mandatory=$true)]
            [array]
            $ArrayName
        )
        $Global:Info = $null
        foreach ($ArrayElement in $ArrayName) {
            $Global:Info += ($ArrayElement -join [Environment]::NewLine).Replace(";",",").Trim("@")
            $Global:Info += ","
                }
            $Global:Info.TrimEnd(",").Replace("=",":")   
            }
# Declare variables
    $StartScript = Get-Date
    $FileName = ""
    $JsonFileExt = ".json"
    $Timestamp = Get-Date -Format "yyyyMMddHHmm"
# Filter for excluding software printers by PortName
    $Filter_PortName = @("pdf",
                        "fax",
                        "ts0",
                        "wsd",
                        "onenote",
                        "document",
                        "nul:",
                        "snagit",
                        "AD_Port",
                        "BIPORT",
                        "DOP7:",
                        "FILE:",
                        "FOXIT_Reader:",
                        "FPR7:",
                        "Journal Note Writer Port:",
                        "MMSPORT:",
                        "Nuance Image Printer Writer Port",
                        "NVK7:",
                        "NVO5:",
                        "TS219",
                        "XPSPort:",
                        "PORTPROMPT:"
    )    
# Check rdp-session
    $RdpCheck = $null -ne $env:Clientname
    if ($RdpCheck -eq "True")
    {
        $Clientname = $env:Clientname
        $SessionName = $env:SESSIONNAME
    } 
    else {
        $Clientname = "Local"
        $SessionName = "Local"
    }
# Get information about Site PC and domain name
    $DomainInfo = (Get-WmiObject Win32_ComputerSystem).Domain
    $SiteInfo = Get-WmiObject Win32_NTDomain | Where-object {($_.Description -eq "CONTOSO") -or ($_.Description -eq "AnCONTOSO")}
# Get information about motherboard
    $BoardInfo = Get-WmiObject Win32_Baseboard | Select-Object Manufacturer, Product
# Get information about processor
    $CPUInfo =  Get-WmiObject win32_processor | Select-Object Manufacturer, Name, NumberOfCores, Caption, SocketDesignation
# Get information about BIOS
    $BIOSInfo = Get-WmiObject Win32_BIOS | Select-Object Manufacturer, Version, Name
# Get information about Physical Memory
    [array]$PhysicalMemoryInfo = Get-WmiObject win32_PhysicalMemory | Select-Object @{Label = '"Capacity"'
    Expression =  {'"{0}"' -f $_.Capacity}}, @{Label = '"DeviceLocator"'
    Expression =  {'"{0}"' -f $_.DeviceLocator}}
# Get information about VideoController
    [array]$VideoControllerInfo = Get-WmiObject Win32_VideoController | Where-Object {$_.Availability -match "3"} | Select-Object Name, VideoProcessor, AdapterRAM, VideoModeDescription, Status
# Get information about DiskDrives
    [array]$HardDiskInfo = Get-WmiObject Win32_DiskDrive | Where-Object {$_.InterfaceType -notmatch "USB"} | Select-Object @{Label = '"Model"'
    Expression =  {'"{0}"' -f $_.Model}}, @{Label = '"InterfaceType"' 
    Expression =  {'"{0}"' -f $_.InterfaceType}}, @{Label = '"SerialNumber"' 
    Expression = { if ($_.SerialNumber) {'"{0}"' -f $_.SerialNumber } else { '"n/a"' } } }, @{Label = '"Status"'
    Expression =  {'"{0}"' -f $_.Status}}
    $HardDiskCount = $HardDiskInfo.Count
# Get information about Network Adapters
    [array]$NetworkAdapterInfo = Get-WmiObject Win32_NetworkAdapter | Where-Object {$_.NetConnectionStatus -gt "0"}
# Get information about physical printers
    [array]$PrinterInfo = Get-WmiObject win32_printer | Where-Object  PortName -notmatch ($Filter_portname -Join "|") | Select-Object @{Label = '"Name"'  
    Expression = {'"{0}"' -f $_.Name.Replace("\","/")}}, @{Label = '"DriverName"' 
    Expression =  {'"{0}"' -f $_.DriverName}}, @{Label = '"Local"' 
    Expression =  {'"{0}"' -f $_.Local}}, @{Label = '"Default"'
    Expression =  {'"{0}"' -f $_.Default}}, @{Label = '"PortName"'
    Expression = {'"{0}"' -f $_.PortName.Replace("\","/")}}, @{Label = '"Status"' 
    Expression =  {'"{0}"' -f $_.Status}}
# Get information about Operating System
    $OSInfo = Get-WmiObject Win32_OperatingSystem | Select-Object Caption, InstallDate, OSArchitecture, Version, @{Label = "NumberOfProcesses"
    Expression = { if ($_.NumberOfProcesses -gt 75) {"gt75"} else {"lt75"} } }
# Get information about logical disks
    [array]$DiskInfo = get-WmiObject win32_logicaldisk | Where-Object {$_.MediaType -eq "12"} | Select-Object @{Label = '"DeviceId"'  
    Expression = {'"{0}"' -f $_.DeviceId}}, @{Label = '"Size"' 
    Expression =  {'"{0}"' -f $_.Size}}, @{Label = '"VolumeName"' 
    Expression =  {'"{0}"' -f $_.VolumeName}}, @{Label = '"FreeSpace"'
    Expression = { if ($_.FreeSpace -lt 10Gb) { '"lt10"' } else { '"gt10"' } } }
# Get information about installed applications
    $AppInfo = Get-WmiObject Win32_Product | Select-Object @{Label='"Name"'; Expression = { if ($_.Name) { '"{0}"' -f $_.Name.Replace('"','') } else { '"n/a"'}} }, @{Label='"Version"'
    Expression = {'"{0}"' -f $_.Version}}
# Get information about Startup Commands
    $AppStartUpInfo = Get-WmiObject -Namespace ROOT\CIMV2 -Class Win32_StartupCommand | Select-Object @{Label='"Name"' 
    Expression = {'"{0}"' -f $_.Name.Replace('"','')}}, @{Label = '"Command"'
    Expression = {'"{0}"' -f $_.command.Replace( '\','/').Replace('"','')}}, @{Label = '"Location"'
    Expression = {'"{0}"' -f $_.Location.Replace('\','/')}},  @{Label = '"User"'
    Expression = {'"{0}"' -f $_.User.Replace('\',')/')}}
# Get information about installed MS updates
    $MSUpdateInfo = Get-WmiObject win32_quickfixengineering | Sort-Object -Property InstalledOn -Descending | Where-Object {($_.InstalledOn).Year -ge (Get-Date).Year} | Select-Object @{Label='"HotFixID"' 
    Expression = {'"{0}"' -f $_.HotFixID}}, @{Label = '"InstalledOn"'
    Expression = {'"{0}"' -f $_.InstalledOn}}
# Get information about personal user certs
    $CertInfo = Get-ChildItem -path cert:\CurrentUser\My | Where-Object {$_.Subject -match "SN"} | Select-Object @{Label = '"Expiring"'
    Expression = {'"{0}"' -f $_.NotAfter.ToShortDateString()} }, @{Label = '"Issuer"'
    Expression = {'"{0}"' -f ($_.Issuer.Split(",") | Where-Object {$_ -match "CN="}).Replace("CN=","").Replace('"',"")} }, @{Label = '"SubjectSName"'
    Expression = {'"{0}"' -f ($_.Subject.Split(",") | Where-Object {$_ -match "SN="}).Replace("SN=","")} }, @{Label = '"SubjectGName"'
    Expression = {'"{0}"' -f ($_.Subject.Split(",") | Where-Object {$_ -match "G="}).Replace("G=","")} }
<#
    Variable $CertInfo given value null if users don't have cert.
    This check is written as a result of which at a given value an array is created with null values falling into the final json
#>
    if ($null -eq $CertInfo) {
        $CertInfo = '{"Certificates" : "is not installed"}'
        }
# Generating json with the necessary information using here-strings
$Basic = @"
{
    "ComputerName" : "$env:COMPUTERNAME",
    "UserName" : "$env:username",
    "ClientName" : "$Clientname",
    "SessionName" : "$SessionName",
    "DomainName" : "$DomainInfo",
    "SiteName"  :   "$($SiteInfo.ClientSiteName)",
    "GeneralHardwareInfo" :
    {
        "CPUManufacturer": "$(($CPUInfo).Manufacturer)",
        "CPUName": "$(($CPUInfo).Name)",
        "CPUNumberOfCores": "$(($CPUInfo).NumberOfCores)",
        "CPUCaption": "$(($CPUInfo).Caption)",
        "CPUSocketDesignation": "$(($CPUInfo).SocketDesignation)",
        "BoardManufacturer": "$(($BoardInfo).Manufacturer)",
        "BoardProduct": "$(($BoardInfo).Product)",
        "BIOSManufacturer": "$(($BIOSInfo).Manufacturer)",
        "BIOSVersion": "$(($BIOSInfo).Version)",
        "BIOSName": "$(($BIOSInfo).Name)",
        "PhysicalMemoryInfo": [
            $(SetJsonArray -ArrayName $PhysicalMemoryInfo)
            ],
        "VideoControllerName": "$(($VideoControllerInfo | Select-Object @{Label="Name"; Expression = {$_.Name}}).Name)",
        "VideoControllerVideoProcessor": "$(($VideoControllerInfo | Select-Object @{Label="VideoProcessor"; Expression = {$_.VideoProcessor}}).VideoProcessor)",
        "VideoControllerVideoModeDescription": "$(($VideoControllerInfo | Select-Object @{Label="VideoModeDescription"; Expression = { if ($_.VideoModeDescription) {$_.VideoModeDescription} else { "n/a" } }}).VideoModeDescription)",
        "VideoControllerAdapterRAM": "$(($VideoControllerInfo | Select-Object @{Label="AdapterRAM"; Expression = {$_.AdapterRAM}}).AdapterRAM / 1Mb) Mb",
        "HardDiskCount" : "$HardDiskCount",
        "HardDiskInfo": [
            $(SetJsonArray -ArrayName $HardDiskInfo)
            ],
        "NetworkAdapterName": [
            $(SetArraytoString -ArrayName $(($NetworkAdapterInfo | Select-Object @{Label="Name"; Expression = {$_.Name}}).Name))
        ],
        "NetworkAdapterMACAddress": [
            $(SetArraytoString -ArrayName $(($NetworkAdapterInfo | Select-Object @{Label="MACAddress"; Expression = { if ($_.MACAddress) { $_.MACAddress } else { "n/a" } }}).MACAddress.Replace(":","")))
        ],
        "PrinterInfo" : [
            $(SetJsonArray -ArrayName $PrinterInfo)
        ]
    },
    "OSInfo":
    {
        "OSInfoCaption" : "$($OSInfo.Caption)",
        "OSInfoInstallDate" : "$($OSInfo.InstallDate)",
        "OSInfoOSArchitecture" : "$($OSInfo.OSArchitecture)",
        "OSInfoOSVersion" : "$($OSInfo.Version)",
        "DiskInfo": [
            $(SetJsonArray -ArrayName $DiskInfo)
        ],
        "AppInfo": [
            $(SetJsonArray -ArrayName $AppInfo)
        ],
        "AppStartUpInfo": [
            $(SetJsonArray -ArrayName $AppStartUpInfo)
        ],
        "MSUpdateInfo": [
            $(SetJsonArray -ArrayName $MSUpdateInfo)
        ],
        "CertInfo": [
            $(SetJsonArray -ArrayName $CertInfo)
        ]
    },
    "TriggerInfo":{
        "OSInfoNumberOfProcesses" : "$($OSInfo.NumberOfProcesses)"
    },
    "AddInfo":{
        "ChassisType" : "$($ChassisTypeInfo.ChassisTypes)"
        }
}
"@
if ($MacAdr.Count -gt "1")
    {
        [array]$NetworkAdapter2 = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE
        $MacAdr = $null
        $MacAdr = $NetworkAdapter2[0] | Select-Object MACAddress
    }
    $MacAdr = $MacAdr.MACAddress.Replace(":","")
    $FileName = $Timestamp + "_" + $MacAdr + "_" + $env:COMPUTERNAME + $JsonFileExt
    $JSON = "$Filename"
    $JSONAN = "$Filename"
    $StopScript=Get-Date
    $TimeToExec=($StopScript-$StartScript).TotalMinutes
# Проверяем домен для отправки логов
    if ($DomainInfo -notmatch "ancontoso.local") {
        $Basic | Out-File $JSON
        LogWrite "$StopScript;$env:username;$env:COMPUTERNAME;$env:USERDNSDOMAIN;$Clientname;$TimeToExec min"
    }
    else {
        $Logfile = ".\$Date.log"
        LogWrite "$StopScript;$env:username;$env:COMPUTERNAME;$DomainInfo;$Clientname;$TimeToExec min"
        $Basic | Out-File $JSONAN
    }
# Debug
#$b = $Basic | ConvertFrom-Json
#$b.OsInfo.MSUpdateInfo