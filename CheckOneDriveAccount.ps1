<#
.SYNOPSIS
Use this script to check if a list of users has a OneDrive account provisioned

.DESCRIPTION

You can use the -FilePath parameter to choose a csv directly. The csv requires the heading of "source_email" and the users will be under that. 
A PNP powershell connection will be made to the tenant and then the users in the CSV will be compared against the tenant to see if they exist


.PARAMETER FilePath
Specifies if you will be loading a file that is not from the a migration folder

.PARAMETER SPOTenant
Specifies the tenant name, no need to add the -admin to the tenant name, we do that for you

#>
param(
[Parameter(mandatory=$true)][string]$FilePath,
[Parameter(mandatory=$true)][string]$SPOTenant
)

function ImportCsv ($FilePath) {
Import-Csv $FilePath
}
function ConnectTenant ($SPOUrl) {
Connect-SPOService -Url "https://$($SPOUrl)-admin.sharepoint.com" -ModernAuth $true
}


function getPersonalSites ($items, $SPOUrl) {
    Write-Host "This may take a long time of no activity depending on how many users are requested"
    $usersList = New-Object System.Data.DataTable
    $null=$usersList.Columns.Add("Email Address")
    $null=$usersList.Columns.Add("OneDrive URL")
    $null=$usersList.Columns.Add("Provisioned")
    $currentitem = 0
    $percentcomplete = 0
    foreach ($item in $items) {
     Write-Progress -Activity "Checking OneDrive" -Status "$percentcomplete% Complete:" -PercentComplete $percentcomplete
     $url = ($item.source_email).replace(".","_")
     $url = ($url).replace("@","_")
     $url = "https://$($SPOUrl)-my.sharepoint.com/personal/$($url)"
     $ErrorActionPreference = "stop"
     $row = $usersList.NewRow()
     $row.'Email Address' = $item.source_email
     $row.'OneDrive URL' = $url
     try{
     Get-SPOSite $url | Out-Null
     $row.Provisioned = "True"
     }
     catch{
     $row.Provisioned = "False or Error"
     }
     finally{
     $usersList.rows.Add($row)
     $currentitem++
     $percentcomplete = [int] (($currentitem / $items.count) * 100)
     }

    }
   $usersList | Format-Table | Out-Host

}

function SpOnlineMgmt {
write-host "Microsoft SharePoint Powershell is not installed on this system!"
$answer = read-Host "Do you wish to install this module? (y/n)"
if ($answer -eq "y" -or $answer -eq "yes") {
Install-Module -name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser}
else{
write-host "Exiting Powershell script. Run this script on a machine with Microsoft SharePoint PowerShell installed"
EXIT
}

}

#main program

#check if right software installed
if(get-module -name Microsoft.Online.SharePoint.PowerShell -ListAvailable){}else{SpOnlineMgmt}

$items = ImportCsv $FilePath
ConnectTenant $SPOTenant
getPersonalSites -items $items -SPOUrl $SPOTenant
