<#
    .SYNOPSIS
    PCInfo2DB
    Version: 0.01 22.11.2019

    © Anton Kosenko mail:Anton.Kosenko@gmail.com
    Licensed under the Apache License, Version 2.0

    .DESCRIPTION
    This script process final json files, compare them to current value in DB and insert changed json in DB.
#>

# requires -version 3

# Log function
$Logfile = ".\2DBInfo.log"
Function LogWrite
{
    Param ([string]$logstring)
    Add-content $Logfile -value $logstring
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
# Function Select string in DB
function SelectData {
    try {                
        $adDB = New-Object System.Data.Odbc.OdbcDataAdapter
        $adDB.SelectCommand = New-Object System.Data.Odbc.OdbcCommand("Select * from $tablename where macaddres = '$macaddres' ORDER by last_login desc limit 1", $Conn2DB) 
        $adDB.Fill($SelectData) | Out-Null
    }
    catch [System.Data.Odbc.OdbcException] {
        $_.Exception
        $_.Exception.Message
        $_.Exception.ItemName
        LogWrite "Problem with select '$TargetFile' Error: '$_.Exception.ItemName'"
        $Global:ErrorMessage += "Problem with select '$TargetFile' .</center>`n"
    }
    
}
# Function Insert string in DB
function InsertData {
    try {
        $adDB = New-Object System.Data.Odbc.OdbcDataAdapter
        $adDB.InsertCommand = New-Object System.Data.Odbc.OdbcCommand("insert into $tablename (last_login, macaddres, hostname, info) values ('$last_login', '$macaddres', '$hostname', '$info')", $Conn2DB)
        $adDB.InsertCommand.ExecuteNonQuery() | Out-Null
    }
    catch [System.Data.Odbc.OdbcException] {
        $_.Exception
        $_.Exception.Message
        $_.Exception.ItemName
        $Global:ErrorMessage += "Problem with insert '$TargetFile' Error: '$_.Exception.ItemName' .</center>`n"
        LogWrite "Problem with insert '$TargetFile' Error: '$_.Exception.ItemName'"
    }
    
}
# Declare Variables
    $Global:ErrorMessage = ""
    $StartScript = Get-Date
    $Infolog = ".\pginfo.log"
    $JSONFolder = "\\path\to$\#Folder\"
    $Timestamp = Get-Date -Format "yyyyMM"
#    $Timestamp = "201910"
    $dbServer = "ip or domain name"
    $dbName = "db name"
    $dbUser = "db user"
    $dbPass = "passwd"
#    $tablename = "table_test"
    $tablename = "table"
    $CntTarget = 0
    $CntProcessedFiles = 0
    Start-Transcript -path  "$Infolog" -append
    LogWrite "######## Start $StartScript ###################" 
# Get all the files that will be inserted into the DB
if (!(test-path -path $JSONFolder)) {
    LogWrite "Folder is unreachable"
    LogWrite "######## Stop $StopScript ###################" 
    $Mail_TextBody += "Folder is unreachable.</center>`n"
    MailError
    Stop-Transcript
    exit
    }
    $TargetFiles = Get-ChildItem $JSONFolder | Where-Object {$_.BaseName -eq $Timestamp} | Get-ChildItem | Where-Object {((($_.CreationTime).Date -le (Get-date).Date) -and ($_.Extension -eq ".json"))}
    $CntTarget = $TargetFiles.Count
# Connect into DB
    $Conn2DB = New-Object System.Data.Odbc.OdbcConnection
    $Conn2DB.ConnectionString= "Driver={PostgreSQL UNICODE};Server=$dbServer;Port=5432;Database=$dbName;Uid=$dbUser;Pwd=$dbPass;"
    $Conn2DB.open()
    if ($Conn2DB.State -eq "Close") {
        LogWrite "DB is ubreachable"
        LogWrite "######## Stop $StopScript ###################" 
        $Mail_TextBody += "DB is ubreachable.</center>`n"
        MailError
        Stop-Transcript
        exit
    }
# Search for files and insert them into the DB
    foreach ($TargetFile in $TargetFiles) {
        $CntProcessedFiles = $CntProcessedFiles + 1
# Parse the file name
        $TargetArray = @{}
        $TargetArray = $TargetFile.BaseName.Split("_")
        $lastlogin = [datetime]::parseexact($TargetArray[0], 'yyyyMMddHHmm', $null)
        $last_login = $lastlogin.ToString("yyyy-MM-dd HH:mm:ss")
        $macaddres = $TargetArray[1]
        $hostname = $TargetArray[2]
        $info = Get-Content $TargetFile.FullName
# Search existing entries by macaddress PC and latest timestamp
        $SelectData = $null
        $SelectData = New-Object System.Data.DataSet
        SelectData    
# Check for changes
        if ($null -ne $SelectData.Tables.Rows){
            $ReferenceObject = $null
            $ReferenceObject = $info -join " "
            $ResultCompare = Compare-Object -ReferenceObject $ReferenceObject -DifferenceObject $SelectData.Tables.Rows[3]
                if ($null -ne $ResultCompare) {
                   InsertData
                   $TargetFile | Rename-Item  -NewName  {$_.Name -replace '.json','.done'} #-WhatIf
                }
                else {
                    $TargetFile | Rename-Item  -NewName  {$_.Name -replace '.json','.done'} #-WhatIf
                }
            }
            else {
                InsertData
                $TargetFile | Rename-Item  -NewName  {$_.Name -replace '.json','.done'} #-WhatIf
                }
        }
# Close database connection
    $Conn2DB.Close()
# Check the number of processed files
    if ($CntTarget -eq $CntProcessedFiles) {
        LogWrite "Processed $CntProcessedFiles from $CntTarget"
    }    
    else {
        Write-host "Many files not processed."
        $Mail_TextBody += "<center>Many files not processed.</center>`n"
        MailError
        Stop-Transcript
        $StopScript = Get-Date
        LogWrite "######## Stop $StopScript ###################"
        Exit        
        }
# Send an email if an error occurs
    if ($Global:ErrorMessage -match "Error") {
        $Mail_TextBody += $Global:ErrorMessage
        MailError
    }
    $StopScript = Get-Date
    Stop-Transcript
    LogWrite "######## Stop $StopScript ###################"