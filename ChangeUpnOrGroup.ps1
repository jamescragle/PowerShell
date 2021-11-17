<#
.synopsis
Use this script when you need to change the email addresses or group domain suffix.
Author: James Cragle

.Description
The fields that must be present in the csv for the script to work are
ObjectID - This is the Group ID or User ID guid (required)
Get the GUID of groups or users by exporting them from Microsoft 365 Admin Center page under "Users -> Active Users" or "Teams & Groups -> Active Teams & Groups"

Recommended add a NewPrimaryEmailAddress to the spreadsheet to know what address you are working with. GUID are unfriendly to humans (optional and only used for tracking)

IMPORTANT - these changes take several minutes to replicate through the system. It's not recommended to run this script back to back without waiting 5 minutes.  
The primary email address will update to the UPN after a few minutes.  
Unified groups will set the current address as an alias
Other group types the suffix is replaced but the original address is lost. 

.parameter FilePath
This will be the location of the CSV file that contains the Object ID's of the users or groups to convert

.parameter UnifiedGroups
This is optional. By default, the script targets Users and not Unified Groups. Change to $true to update Unified Groups suffix

.parameter DomainSuffix
This is optional. By default, the script will read the .onmicrosft.com tenant name and assign. Choose you own domain if you need change UPN suffix

.parameter RemoveAlias
This is optional. By default, m365 unified groups and users when changing the sufix will transfer the old primary as an alias. Select $true to delete all aliases and leave only the primary smtp address.

.example
PS> .\ChangeUpnOrGroup.ps1 -filePath .\fileWithUsers.csv
Will change the users upn in the csv to domain.onmicrosoft.com

.example
PS> .\ChangeUpnOrGroup.ps1 -filePath .\fileWithGroups.csv -UnifiedGroups $true
Will change unified groups suffix in the csv to domain.onmicrosoft.com

.example
PS> .\ChangeUpnOrGroup.ps1 -filePath .\fileWithUsers.csv -DomainSuffix "example.com"
Will change the upn of the users in the csv to a upn of example.com

.example
PS> .\ChangeUpnOrGroup.ps1 -filePath .\fileWithGroups.csv -UnifiedGroups $true -DomainSuffix "example.com
Will change unified groups suffix in the csv to example.com

.example
CSV Format
ObjectID                             | DisplayName  | EmailAddress
12345678-1234-1234-1234-123456789abc | bob smith    | bob@example.com

#>

param (
[Parameter(Mandatory=$true)][string] $FilePath = "",
[Parameter(Mandatory=$false)][boolean] $UnifiedGroups = $false,
[Parameter(Mandatory=$false)][string] $DomainSuffix = "",
[Parameter(Mandatory=$false)][boolean] $RemoveAlias = $false 

)


#automatically set variables
$currentDomain = ""
Add-Type -AssemblyName presentationframework

function removeAlias($currentGroup){
    $aliasgroup = Get-UnifiedGroup $currentGroup
    $aliasgroup.emailaddresses | % {if ($_ -cmatch "smtp"){if($_ -notlike "*onmicrosoft*"){$aliasgroup | Set-UnifiedGroup -EmailAddresses @{remove=$_}}}}
}

function ChangeUnifiedGroups () {

$fileData = Import-Csv -Path $filePath
foreach ($item in $fileData) {
#this command will change the UPN, set the upn as default email address and put the the default as alternate
   
  $ErrorActionPreference = "Stop"
  try {
  $thisgroup = Get-group $item.ObjectID
  if ($thisgroup.recipienttypedetails -eq "MailUniversalDistributionGroup"){
    $firstPartUpn = $thisgroup.WindowsEmailAddress.split("@")[0] 
    $newUpn = $firstPartUpn + "@" + $DomainSuffix
    write-host $newUpn
    #$thisgroup | set-group -WindowsEmailAddress $newUpn
  }
  if ($thisgroup.recipienttypedetails -eq "GroupMailbox"){
    $thisUG = Get-UnifiedGroup $item.ObjectID
    $firstPartUpn = $thisUG.PrimarySmtpAddress.split("@")[0] 
    $newUpn = $firstPartUpn + "@" + $DomainSuffix
    write-host $newUpn
    $thisUG | Set-UnifiedGroup -PrimarySmtpAddress $newupn
    if ($RemoveAlias -eq $true){
        removeAlias $item.ObjectID 
    }
  }
  if ($thisgroup.recipienttypedetails -eq "MailUniversalSecurityGroup"){
    $firstPartUpn = $thisgroup.WindowsEmailAddress.split("@")[0] 
    $newUpn = $firstPartUpn + "@" + $DomainSuffix
    write-host $newUpn
    #$thisgroup | set-group -WindowsEmailAddress $newUpn
  }
  
  }
  catch {write-host "$($item.objectid) is not a valid group"}
  $ErrorActionPreference = "Continue"
}
}

function ChangeUsers () {

$fileData = Import-Csv -Path $filePath
foreach ($item in $fileData) {
#this command will change the UPN, set the upn as default email address and put the the default as alternate
  $thisuser = Get-AzureADUser -ObjectId $item.ObjectID 
  $firstPartUpn = $thisuser.userprincipalname.split("@")[0] 
  $newUpn = $firstPartUpn + "@" + $DomainSuffix
  $thisuser | Set-AzureADUser -UserPrincipalName $newupn
}

} 

function ShowMessages($messageTitle, $MessageBody, $ntype) {
$message = [System.Windows.MessageBox]::Show($MessageBody,$messageTitle,$nType)

return $message
} 

function getOnMSDomain () {
$domains = (Get-AzureADTenantDetail).verifieddomains
foreach ($item in $domains) {
    if ($item.initial -like "true") {
    return $item.name
    }
}
}

function connectToExchange() {
#Connect & Login to ExchangeOnline (MFA)
$getsessions = Get-PSSession | Select-Object -Property State, Name
$isconnected = (@($getsessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
If ($isconnected -ne "True") {
Connect-ExchangeOnline
}

}

# start here connect-azure ad
if(([Microsoft.Open.Azure.AD.CommonLibrary.AzureSession]::AccessTokens).Count -eq 0 ){
$null = Connect-AzureAD
}


if ($DomainSuffix -eq "") {
$DomainSuffix = getOnMSDomain
}

if ($RemoveAlias -eq $true){
$selection = showmessages "WARNING! Remove Alias Selected" "Choosing to remove the alias will delete any non .onmicrosoft.com aliases. This cannot be undone." "YesNo"
switch($selection){
No{
write-host "Request Cancelled" -BackgroundColor Yellow -ForegroundColor Black
exit}
}
}


if ($UnifiedGroups -eq $false) {
$selection = showmessages  "Change the User UPN" "You selected to change the UPN of the users saved in the file $filePath to $DomainSuffix, are you sure you want to continue? This cannot be undone" "YesNo"
switch($selection){
    Yes {changeusers
        $null = showmessages "Processing Complete" "Please wait a few minutes to see results in the Admin Center." "Ok"
        }

    No {write-host "Request Cancelled" -BackgroundColor Yellow -ForegroundColor Black}
    }
} 

if ($UnifiedGroups -eq $true) {
$selection = showmessages  "Change the Groups Domain Name" "You selected to change the group names saved in the file $filePath to $DomainSuffix, are you sure you want to continue? This cannot be undone" "4"
switch($selection){
    Yes {connectToExchange
        ChangeUnifiedGroups
        $null = showmessages "Processing Complete" "Please wait a few minutes to see results in the Admin Center." "Ok"
        }

    No {write-host "Request Cancelled" -BackgroundColor Yellow -ForegroundColor Black}
    }
}
