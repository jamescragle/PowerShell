<#
.SYNOPSIS
Use this script to review the status of or disable the accounts of syncplicity users using an app created in the Box admin console. Fully read the description section.

.DESCRIPTION
This script will make a REST call to the Box api and allow you to set the access level of users based on a csv of users. There are two options for running the script. 
If you just run it with no options, it will allow you to choose a csv file through a GUI from wherever you want.
You will need to also choose the csv header under which your users email addresses are located. It's free form text to it can match anything in your csv.
You can also use the command line to start this script.  The script will create a folder in the root of c for logging. You can update it if you feel adventurous.  
You will need 3 things to make it work located section "retrieveToken", all of which are located in the admin console of box. These are the first 3 items in the body array.
Client ID - This is generated when you create an application in box
Client Secret - This is generated when you create an application in box
Box Subject ID - This is the company ID located on the billing page (aka Enterprise ID). 
See this for more details https://developer.box.com/guides/authentication/client-credentials/ 


.PARAMETER FilePath
    Specifies if you will be loading a file that is not from the a migration folder
    If no file path provided, will prompt a file explorer

.PARAMETER selectedRights
    In box, there are only 4 valid options and they are "active","inactive","cannot_delete_edit","cannot_delete_edit_upload"
    If you are using the command line, you will need to provide the value
    Running the script interactive in ISE and you can pick from a list of options.

.PARAMETER UsersEmailColumn
    This is the column in the csv that contains your email addresses which will map to the users login.
#>
param(
[string]$FilePath = "",
[validateset("active","inactive","cannot_delete_edit","cannot_delete_edit_upload","")]$selectedRights,
[string]$UsersEmailColumn =""
)
#plumbing stuff for file picker and logging
Add-Type -AssemblyName system.windows.forms
if (Test-Path -Path "c:\logfiles") {
    $rootPath = "C:\logfiles"
}else{
    $ErrorActionPreference = "stop"
    try{
        New-Item -Path "c:\" -Name "logfiles" -ItemType "directory"
        $rootPath = "C:\logfiles"
    }catch{
        $null = read-host "Cannot create the folder 'c:\logfiles' manually create it and run this again. Press enter to exit"
        exit;
    }
}
if ($PSVersionTable.PSVersion.Major -ne "5"){
    write-host "This script requires PowerShell version 5"
    exit;
}

#variables
$date = Get-Date -Format "yyyy-MM-dd"
$procLog = "$($rootpath)\PostProcessing-process-$date.log" 
Start-Transcript -Path "$($rootPath)\Transcript-results-$date.log" -Append -Force 
$tokenExpirationTime
$token = ""
$usersFromBox

#functions start
Function Log(){
    Param( [string]$Text ="",
    [string]$Flag = ""
)
if($Flag -eq "Error"){
    Write-Host $Text -ForegroundColor Red
    $log = new-object System.IO.StreamWriter("$procLog", [System.Text.Encoding]::UTF8)
    $log.WriteLine("$(Get-Date),$text")
    $log.Close()
    }
if($Flag -eq "Warning"){
    Write-Host $Text -ForegroundColor Yellow
    $log = new-object System.IO.StreamWriter("$procLog", [System.Text.Encoding]::UTF8)
    $log.WriteLine("$(Get-Date),$text")
    $log.Close()
    }

if($Flag -eq ""){
    Write-Host $Text -ForegroundColor Green
    $log = new-object System.IO.StreamWriter("$procLog", [System.Text.Encoding]::UTF8)
    $log.WriteLine("$(Get-Date),$text")
    $log.Close()
    }
}

Function retrieveToken {
$url = "https://api.box.com/oauth2/token"
$headers = @{"content_type"="application/x-www-form-urlencoded"
}
$body = @{
"client_id"=""
"client_secret"=""
"box_subject_id"=""
"box_subject_type"="enterprise"
"grant_type"="client_credentials"
}
$ErrorActionPreference = "stop"
try{
$results = Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $body
$script:tokenExpirationTime = (Get-Date).AddSeconds(5100) 
Log "Updating token expiration time to $($tokenExpirationTime)"
return $results.access_token
}
catch{
Log "$($Error)" "Error"
return "Error"
}
}

function getBoxUserAccount ($emailaddress) {
    if ($tokenExpirationTime -lt $(get-date)){
        $script:token = retrieveToken
    }
    $turl = "https://api.box.com/2.0/users?filter_term=$emailaddress"
    $theaders = @{
        "Authorization" = "Bearer $($script:token)"
        "Content-Type" = "application/json"
    }
    $ErrorActionPreference = "Stop"
    try{
    $trequest = Invoke-RestMethod -Uri $turl -Headers $theaders -Method Get
    Log "request for $emailaddress retrieved ID successfully"
    return $trequest.entries[0].id
    }
    catch{
        Log $_.exception.response.statusdescription, "Error"
        exit;
    }

    
}

Function updateUsers ($batchfile, $mode, $emailcolumn )  {
    $usersList = New-Object System.Data.DataTable
    $null=$usersList.Columns.Add("Email Address")
    $null=$usersList.Columns.Add("Setting")
    Log "$($mode) Users" 
    foreach($batch in $batchfile){
        if ($script:tokenExpirationTime -lt $(Get-Date)) {
        $script:token = retrieveToken 
        }
        $userid = getBoxUserAccount $batch.($emailcolumn)
        if (($userid -match "error") -or ($userid -eq "")){
            $row = $usersList.NewRow()
            $row.'Email Address' = $batch.email
            $row.Setting = "Error getting user"
            $userslist.rows.add($row)
            log "Count not find $($batch.email)"
        }else{
            $url = "https://api.box.com/2.0/users/$($userid)"
            $headers = @{
            "Content-Type" = "application/json"
            "Authorization"="Bearer $($token)"
            }
            $body = @{"status"= $mode} | ConvertTo-Json -Compress
    
            $count=1
        $ErrorActionPreference = "Stop"
        do{
            try{
                $request = Invoke-RestMethod -Uri $url -Headers $headers -Body $body -Method PUT 
                $row = $usersList.NewRow()
                $row.'Email Address' = $request.login
                $row.Setting = $request.status
                $usersList.Rows.Add($row)
                $count=11
            }
            catch{
                if($_.exception.response.statuscode.value__ -ne 429){
                    Log "Status code: $($_.exception.response.statuscode.value__)" "Error"
                    Log "Status Description:  $($_.exception.response.StatusDescription)" "Error"
                    $row = $usersList.NewRow()
                    $row.'Email Address' = $batch.($emailcolumn)
                    $row.Setting = "Error getting user"
                    $userslist.rows.add($row)
                    $count = 11
                }
                else{
                    Log "Status code: $($_.exception.response.statuscode.value__) try number $($count) of 10" "Error"
                    Log "Status Description:  $($_.exception.response.StatusDescription) retrying $($batch.($emailcolumn)) in 10 seconds" "Error"
                    start-sleep -Seconds 5
                    $count++
                }
            }
        }
        until($count -gt 10)
        }
    }
    $usersList | Format-Table | Out-Host
}


Function importSpreadsheet {
    Param( [string]$spreadsheet
)
    $ErrorActionPreference = "Stop"
    try{
        $results = Import-Csv (Get-ChildItem -Path $spreadsheet) 
        Log "FileName: $($spreadsheet) imported"
        return $results
    }catch{
        Log $($Error), "Error"
        Stop-Transcript
        exit;
    }
 
}

function ColumnToImport ($testcolumn, $batchfile){
    if($testcolumn -eq ""){
        $testcolumn = read-host "What column name is being imported from? (ctrl-c to quit)"
    }
    #get all spreadsheet columns
    $cols = $batchfile | Get-Member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
    if(!$cols.Contains($testcolumn)){
        Log "The spreadsheet doesnt contain the column you are referring to"
        Stop-Transcript
        exit;
    }
    return $testcolumn
}

#Program start
#run from the command line or the ISE window

if ($FilePath -eq ""){ #test if the parameter is empty, if so show dialog
    $filebrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{Filter = 'csv files (*.csv)|*.csv'}
    $null = $filebrowser.ShowDialog()
    if ($filebrowser.FileName -ne ""){
    $FilePath = $filebrowser.FileName
    }else{
        $null = read-host "You did not select a file. Press enter to exit"
        Stop-Transcript
        exit;
    }
    }
    # check to see if the variable selected rigths is empty, if yes, present a switch to get imput
    $invalid = $true
    if ($selectedRights -eq $null){
        while ($invalid -eq $true){
        write-host "1 - active"
        write-host "2 - inactive"
        write-host "3 - cannot delete or edit"
        write-host "4 - cannot delete, edit or upload"
        write-host "Press ctrl-c to quit"
        $getChoice = Read-Host "Choose the number of the action you wish to apply to the users"
        switch ($getChoice){
            "1" {$selectedRights = "active"; $invalid = $false}
            "2" {$selectedRights = "inactive"; $invalid = $false}
            "3" {$selectedRights = "cannot_delete_edit"; $invalid = $false}
            "4" {$selectedRights = "cannot_delete_edit_upload"; $invalid = $false}
            default {$invalid = $true}
        }
        }
    }
    Log "selection criteria met, attempting to import the spreadsheet"
    if (Test-Path $FilePath){
        $batchfile = importSpreadsheet -spreadsheet $FilePath
        $UsersEmailColumn = ColumnToImport $UsersEmailColumn $batchfile
        $token = retrieveToken
            if($token -eq "error"){
            write-host "There was an error retrieving the token from Box. Try again"
            Stop-Transcript
            Exit;}
            else{
                updateUsers $batchfile $selectedRights $UsersEmailColumn
            }
        }else{
           Log "No file found, check folder path and contents" "Error"
        }
    
    Log "Reading from manually selected csv file"

Stop-Transcript
