<#
.SYNOPSIS
Use this script to check if a list of users has a OneDrive account provisioned

.DESCRIPTION

You can use the -FilePath parameter to choose a csv directly. The csv requires the heading of "source_email" and the users will be under that. 
A PNP powershell connection will be made to the tenant and then the users in the CSV will be compared against the tenant to see if they exist


.PARAMETER FilePath
Specifies if you will be loading a file that is not from the a migration folder

#>
param(
[Parameter(mandatory=$true)][string]$FilePath,
[Parameter(mandatory=$true)][string]$SPOTenant
)

function ImportCsv ($FilePath) {
Import-Csv $FilePath
}
function ConnectTenant ($SPOUrl) {
Connect-PnPOnline -Url "https://$($SPOUrl)-admin.sharepoint.com" -Interactive
}

function getPnPPersonalSites ($items, $SPOUrl){
    Write-Host "This may take a long time depending on how may users are requested"
    $usersList = New-Object System.Data.DataTable
    $null=$usersList.Columns.Add("Email Address")
    $null=$usersList.Columns.Add("OneDrive URL")
    $null=$usersList.Columns.Add("Provisioned")
    $currentitem = 0
    $percentcomplete = 0
    foreach($item in $items){
        Write-Progress -Activity "Checking Onedrive" -Status "$percentcomplete% Complete:" -PercentComplete $percentcomplete
        $ErrorActionPreference = "stop"
        $row = $usersList.NewRow()
        $row.'Email Address' = $item.source_email
        
        try{
            $odurl = (Get-PnPUserProfileProperty -Account $item.source_email).PersonalUrl
            $row.Provisioned = "True"
         }
         catch{
            $row.Provisioned = "False or Error"
         }
         finally{
            if ($odurl){
                $row.'OneDrive URL' = $odurl
            }else{
                $row.Provisioned = "False or Error"
            }
            $usersList.rows.Add($row)
            $currentitem++
            $percentcomplete = [int] (($currentitem / $items.count) * 100)
        }
    }
      $usersList | Format-Table | Out-Host
}

function SpOnlineMgmt {
write-host "Microsoft PnP SharePoint Powershell is not installed on this system!"
$answer = read-Host "Do you wish to install this module? (y/n)"
if ($answer -eq "y" -or $answer -eq "yes") {
Install-Module -name PnP.PowerShell -RequiredVersion 1.7.0 }
else{
write-host "Exiting Powershell script. Run this script on a machine with Microsoft PnP PowerShell installed"
EXIT
}

}

#main program
write-host "start"
#check if right software installed
if(get-module -name PnP.PowerShell -ListAvailable){}else{SpOnlineMgmt}



$items = ImportCsv $FilePath
ConnectTenant $SPOTenant
getPnPPersonalSites -items $items -SPOUrl $SPOTenant
