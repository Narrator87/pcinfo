<#
    .SYNOPSIS
    PCInfo (Server-Side)
    Version: 0.04 27.11.2019

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
    $FinalFolder = $JSONFolder + $Timestamp
    $RemovingFiles = $null
    $CntFiles = 0
    $CntFilesAll = 0
    $CntFilesUniq = 0
    Start-Transcript -path  "$Infolog" -append
    LogWrite "######## Start $StartScript ###################" 
# Create folder for final files
    if (!(test-path -path $FinalFolder)) {new-item $FinalFolder -itemtype directory | Out-Null}
    if (!(test-path -path $JSONFolder)) {
        LogWrite "Folder is unreachable"
        LogWrite "######## Stop $StopScript ###################" 
        $Mail_TextBody += "Folder is unreachable.</center>`n"
        MailError
        Stop-Transcript
        exit}
    $Files = Get-ChildItem $JSONFolder | Where-Object {($_.Extension -eq ".json") -and (($_.CreationTime).Date -eq (Get-date).Date.AddDays(-1))}
    $CntFiles = Get-ChildItem $JSONFolder | Where-Object {($_.Extension -eq ".log") -and (($_.CreationTime).Date -eq (Get-date).Date.AddDays(-1))} | Get-Content
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
    $UniqFiles = Get-ChildItem $JSONFolder | Where-Object {($_.Extension -eq ".json") -and (($_.CreationTime).Date -eq (Get-date).Date.AddDays(-1))}
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
    $ADCompGroup = Get-ADPrincipalGroupMembership (Get-ADComputer $InfoJson.ComputerName) | Select-Object name
    $ADUserGroup = Get-ADPrincipalGroupMembership $InfoJson.UserName | Select-Object name
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
    "TriggerInfo":
        $($InfoJson.TriggerInfo | ConvertTo-Json),
    "ADInfo":
    {
        "ADCompGroup" :  [
            $(SetArraytoString -ArrayName $($ADCompGroup.Name))
            ],
        "ADCompGroup" :  [
            $(SetArraytoString -ArrayName $($ADUserGroup.Name))
            ]
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
    $StopScript=Get-Date
    Stop-Transcript
    LogWrite "######## Stop $StopScript ###################"