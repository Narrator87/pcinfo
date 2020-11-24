<#
    .SYNOPSIS
    PCInfo (Server-Side)
    Version: 0.05 17.09.2020

    © Anton Kosenko mail:Anton.Kosenko@gmail.com
    Licensed under the Apache License, Version 2.0

    .DESCRIPTION
    This script parse json files given from pcinfo (client-side), adds info from MS AD and write final json with necessary information about PC.
#>

# requires -version 3

# Log function
    $Logfile = ".\AdvInfo.log"
    Function LogWrite
    {
        Param ([string]$logstring)
        Add-content $Logfile -value $logstring
    }
# Function convert array to string   
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
# Mail sending function
function MailError
{
    $PSEmailServer = "ip or domain name"
    $Mail_HTMLBody = "<head><style>table {border-collapse: collapse; padding: 2px;}table, td, th {border: 1px solid #ffffff;}</style></head>"
    $Mail_HTMLBody += "<body style='background:#ffffff'><font face='Courier New'; size='2' color=#cc0000>"		
    $Mail_Subject = "ArchDoc $Datestamp"
    $Mail_HTMLBody += "<center><h2>ArchDoc Error Report</h2></center>"
    $Mail_HTMLBody += $Mail_TextBody
    $Mail_HTMLBody += "</font></body>"
    Send-MailMessage -From "service account mail" -To "admin mail" -Subject $Mail_Subject -Body $Mail_HTMLBody -BodyAsHtml -Encoding UTF8
    }
# Declare variables
    $StartScript = Get-Date
    $Infolog = ".\info.log"
    $FileName = ""
    $Timestamp = Get-Date -Format "yyyyMM"
    $JSONFolder = "\\path\to$\#Folder\"
    $JSONAnFolder = "\\path\to$\#Folder\"
    $FinalFolder = $JSONFolder + $Timestamp
    $Logfolder = $JSONFolder + "Login\"+ $Timestamp + ".csv"
    $RemovingFiles = $null
    $CntFiles = 0
    $CntFilesAll = 0
    $CntFilesUniq = 0
    Start-Transcript -path  "$Infolog" -append
    LogWrite "######## Start $StartScript ###################" 
    if ($StartScript.Day -eq "01") {
        $Timestamp = $Timestamp - 1
    }
# Create folder for final files
    if (!(test-path -path $FinalFolder)) {new-item $FinalFolder -itemtype directory | Out-Null}
    if (!((test-path -path $JSONFolder) -and (test-path -path $JSONAnFolder))) {
        LogWrite "Folder is unreachable"
        LogWrite "######## Stop $StopScript ###################" 
        $Mail_TextBody += "Folder is unreachable.</center>`n"
        MailError
        Stop-Transcript
        exit}
    $Files = Get-ChildItem $JSONFolder, $JSONAnFolder | Where-Object {($_.Extension -eq ".json") -and (($_.CreationTime).Date -eq (Get-date).Date.AddDays(-1))}
    $CntFilesCONTOSO = Get-ChildItem $JSONFolder | Where-Object {($_.Extension -eq ".log") -and (($_.CreationTime).Date -eq (Get-date).Date.AddDays(-1))} | Get-Content
    $CntFilesAnCONTOSO = Get-ChildItem $JSONAnFolder | Where-Object {($_.Extension -eq ".log") -and (($_.CreationTime).Date -eq (Get-date).Date.AddDays(-1))} | Get-Content
    $CntFiles = $CntFilesCONTOSO + $CntFilesAnCONTOSO
# Compare files with same content
    foreach ($File in $Files) {
        $CntFilesAll = $CntFilesAll + 1
        $FileName = $File.FullName
        $SameFile = $File.BaseName.Split("_")
        $SameFiles = $Files | Where-Object {$_.Basename -match $SameFile[1]}
            If ($SameFiles.Count -gt "1") {
                $CntTargetStart = 1
                $CntTargetEnd = $SameFiles.Count
                Do {
                    $CntTargetStart | Out-Null
# Debug                
#                $SameFiles[0].FullName
#                $SameFiles[$CntTargetStart].FullName
                    $ResultCompare = Compare-Object $(Get-Content $SameFiles[0].FullName) $(Get-Content $SameFiles[$CntTargetStart].FullName)
                    if ($null -eq $ResultCompare) {
                        $RemovingFiles = @()
                        $RemovingFiles += $SameFiles[$CntTargetStart].FullName
                    }
                    $CntTargetStart++
                        }
                while ($CntTargetStart -lt $CntTargetEnd)
                    }       
        }
    if (!($null -eq $RemovingFiles)) {Remove-Item -Path $RemovingFiles -ErrorAction SilentlyContinue}
    if ($Files.Count -eq $CntFilesAll) {
        Write-Host "Processed $CntFilesAll in "$CntFiles.Count". Remove "$RemovingFiles.Count" file(s)"
    }
    else {
        Write-Host "Processed $CntFilesAll in "$CntFiles.Count"."
        LogWrite "Many files not processed."
        $Mail_TextBody += "Many files not processed.</center>`n"
        MailError
    }
# Check count PCs writing files 
    if ($CntFiles.Count -ne $CntFilesAll) {
        $BadWrite = $CntFiles.Count - $CntFilesAll
        Write-Host "$Badwrite PC's can't write output json"
        LogWrite "$Badwrite PC's can't write output json."
    }
    $UniqFiles = Get-ChildItem $JSONFolder, $JSONAnFolder | Where-Object {($_.Extension -eq ".json") -and (($_.CreationTime).Date -eq (Get-date).Date.AddDays(-1))}
    foreach ($File in $UniqFiles) {
        $CntFilesUniq = $CntFilesUniq + 1
        $FileName = $File.FullName
        $InfoJson = $null
        $InfoJsonCheck = Get-Content $FileName
# Check availability escape symbols in files
        if ($InfoJsonCheck -match "\\") {
            $InfoJsonCheck = $InfoJsonCheck.Replace("\","/")
        }
       $InfoJson = $InfoJsonCheck | ConvertFrom-Json -ErrorAction SilentlyContinue
# Get info about PC from Active Directory
        if ($InfoJson.DomainName -eq "ancontoso.local"){
            $SearchBase = "DC=ancontoso,DC=local"
            $DCServer = "dc.ancontoso.local"
            $UserName = $InfoJson.UserName
            $TestADUser = Get-ADUser -LDAPFilter "(sAMAccountName=$UserName)"
        }
        else {
            $SearchBase = "DC=contoso,DC=local"
            $DCServer = "dc.contoso.local"
        }
        if ($null -ne $TestADUser) {
            $ADUserGroup = Get-ADPrincipalGroupMembership $InfoJson.UserName | Select-Object name
        }
        else {
            $ADUserGroup = Get-ADPrincipalGroupMembership $InfoJson.UserName -Server $DCServer | Select-Object name
        }
        $ComputerName = $InfoJson.ComputerName
        if ($null -ne $ComputerName){
            $ADCompGroup = Get-ADPrincipalGroupMembership (Get-ADComputer -Filter {Name -eq $ComputerName} -SearchBase "$SearchBase" -Server "$DCServer") | Select-Object name 
        }
# Get total physical memory 
        $TotalMemory = 0
        $InfoJson.GeneralHardwareInfo.PhysicalMemoryInfo.Capacity | foreach-object {$TotalMemory += [int64]$_}
# Get chassis type
        if ($InfoJson.AddInfo.ChassisType -eq "8" -or $InfoJson.AddInfo.ChassisType -eq "9" -or $InfoJson.AddInfo.ChassisType -eq "10" -or $InfoJson.AddInfo.ChassisType -eq "11" -or $InfoJson.AddInfo.ChassisType -eq "12" -or $InfoJson.AddInfo.ChassisType -eq "14" -or $InfoJson.AddInfo.ChassisType -eq "18" -or $InfoJson.AddInfo.ChassisType -eq "21") {
            $ChassisType = "Notebook"    
        }
        else {
            $ChassisType = "Workstation"
        }
# Get current time
        if ($FileName -notmatch $Timestamp) {
            $CMOS = "CMOS checksum error"
            $FileModName = $File.Name.Split("_")
            $TrueTime = Get-Date -Format "yyyyMMddHHmm"
            if ($null -eq $FileModName[3]) {
                $File = $TrueTime + "_" + $FileModName[1] + "_" + $FileModName[2]
            }
        else {
            $File = $TrueTime + "_" + $FileModName[1] + "_" + $FileModName[2] + "_" + $FileModName[3]
        }
        }
        else {
            $CMOS = "CMOS OK"
        }
$Basic = @"
{
    "ComputerName" : "$($InfoJson.ComputerName)",
    "UserName" : "$($InfoJson.UserName)",
    "ClientName" : "$($InfoJson.ClientName)",
    "SessionName" : "$($InfoJson.SessionName)",
    "DomainName" : "$($InfoJson.DomainName)",
    "SiteName"  :   "$($InfoJson.SiteName)",
    "GeneralHardwareInfo" : 
        $($InfoJson.GeneralHardwareInfo | ConvertTo-Json),
    "OSInfo":
        $($InfoJson.OSInfo | ConvertTo-Json),
    "TriggerInfo":{
        "OSInfoNumberOfProcesses" : "$($InfoJson.TriggerInfo.OSInfoNumberOfProcesses)",
        "CMOSStatus" : "$($CMOS)"
        },
    "ADInfo":
    {
        "ADCompGroup" :  [
            $(SetArraytoString -ArrayName $($ADCompGroup.Name))
            ],
        "ADCompGroup" :  [
            $(SetArraytoString -ArrayName $($ADUserGroup.Name))
            ]
    },
    "AddInfo":{
    "TotalMemory" : "$($TotalMemory / 1Gb)",
    "ChassisType" : "$($ChassisType)"
    }
}
"@
    $AdvJSON = "$FinalFolder\$File"
    $Basic | Out-File $AdvJSON #-WhatIf
        if ((Test-Path $AdvJSON) -eq "True") {
            Remove-Item -Path $FileName #-WhatIf
        }    
        }
    if ($UniqFiles.Count -eq $CntFilesUniq) {
        Write-Host "Processed $CntFilesUniq in "$UniqFiles.Count"."
        }
    else {
        Write-Host "Processed $CntFilesUniq in "$UniqFiles.Count"."
        LogWrite "Many files not processed."
        LogWrite "######## Stop $StopScript ###################" 
        $Mail_TextBody += "Many files not processed.</center>`n"
        MailError
        }
# Process log files
    $LogFilesCONTOSO = Get-ChildItem $JSONFolder | Where-Object {($_.Extension -eq ".log") -and ($_.FullName -match $Timestamp)} | Get-Content
    $LogFilesAnCONTOSO = Get-ChildItem $JSONAnFolder | Where-Object {($_.Extension -eq ".log") -and ($_.FullName -match $Timestamp)} | Get-Content
    $LogFiles = $LogFilesCONTOSO + $LogFilesAnCONTOSO
    $RemovingLogFiles = Get-ChildItem $JSONFolder, $JSONAnFolder | Where-Object {($_.Extension -eq ".log") -and ($_.FullName -match $Timestamp)}
    $RemovingLogFolders = Get-ChildItem $JSONFolder, $JSONAnFolder | Where-Object {($_.Attributes -eq "Directory") -and ($_.Name -notmatch "Log") -and (($_.CreationTime).Date -lt (Get-date).Date.AddMonths(-3))}
    Add-content $Logfolder -value $LogFiles
    if ($null -ne $RemovingLogFiles) {
        Remove-Item -Path $RemovingLogFiles.Fullname
    }
    if ($null -ne $RemovingLogFolders) {
        Remove-Item -Path $RemovingLogFolders.Fullname -Recurse
    }
# Process log files with false time
    $Falsefolder = $JSONFolder + "Login\falsedate" + ".csv"
    $FalseTimeLogFiles = Get-ChildItem $JSONFolder, $JSONAnFolder | Where-Object {($_.Extension -eq ".log") -and ($_.FullName -notmatch $Timestamp) -and (($_.CreationTime).Date -ne (Get-date).Date.AddDays(-1))}
    $ContentFalseTimeLogFiles = $FalseTimeLogFiles | Get-Content
    Add-content $Falsefolder -value $ContentFalseTimeLogFiles
    if ($null -ne $FalseTimeLogFiles) {
        Remove-Item -Path $FalseTimeLogFiles.Fullname
    }
    $StopScript=Get-Date
    Stop-Transcript
    LogWrite "######## Stop $StopScript ###################"