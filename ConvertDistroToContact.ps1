<#
.synopsis
Use this script when you need to convert a distribution list to a contact.
Author: James Cragle

.Description
If you ever have a need to convert a Distribution Group into a contact card and use that to forward mail.
This script is only useful if you are migrating from one tenant to another and you want to clean up distro's but still keep forwarding mail. 
The fields that must be present in the csv for the script to work are
DisplayName - This is the Display Name of the Distribution Group that will be turned into a contact (required)
ForwardingDetails - this is the email address in a Microsoft 365 Tenant that will act as the destination. (required)

.parameter ImportCsvFile
This will be the location of the CSV file that contains the names of the distros to convert

.example
PS> .\ConvertDistroToContact.ps1 -ImportCsvFile .\testconvertdistro.csv
Will use the csv list to convert the distro into a contact. 

#>

Param(
[string] $ImportCsvFile
)

function deleteDistroList($currentDistro){
    Write-Log "ALERT: removing Distribution list $currentDistro"
    $ErrorActionPreference = "Stop"
    try{ 
        Remove-DistributionGroup -Identity $currentDistro -Confirm:$false
        Write-Log "SUCCESS: removed Distribution list $currentDistro"
     }
    catch{
    Write-Log "ERROR: Could not delete Distribution list $currentDistro `n $($error[0])" 
    }
    $ErrorActionPreference = "Continue"
}

function retrieveDistroData($currentDistro){
    Write-Log "INFO: Testing for existence of $currentDistro"
    if((Get-DistributionGroup $currentDistro -erroraction "silentlycontinue") -eq $null){
        Write-Log "ALERT: The $currentDistro was not found, skipping."
        return "@{DisplayName=;PrimarySmtpAddress=}"
    }
    else {
        Write-Log "INFO: Retrieving data for $currentDistro"
        return (Get-DistributionGroup -Identity $currentDistro | select DisplayName, PrimarySmtpAddress)
        
    }
}

Function Write-Log ($LogString) {
    $LogStatus = $LogString.Split(":")[0]
    If ($LogStatus -eq "SUCCESS") {
        Write-Host $LogString -ForegroundColor Green
        $LogString | Out-File $RunLog -Append  }
    If ($LogStatus -eq "INFO") {
        Write-Host "$LogString" -ForegroundColor Cyan
        $LogString | Out-File $RunLog -Append }
    If ($LogStatus -eq "ALERT") {
        Write-Host $LogString -ForegroundColor Yellow
        $LogString | Out-File $RunLog -Append }
    If ($LogStatus -eq "ERROR") {
        Write-Host $LogString -BackgroundColor Red
        $LogString | Out-File $RunLog -Append
        "`n" | Out-File $ErrorLog -Append
        $LogString | Out-File $ErrorLog -Append }
    If ($LogStatus -eq "AUDIT") {
        Write-Host $LogString -ForegroundColor DarkGray
        $LogString | Out-File $RunLog -Append  }
    If ($LogStatus -eq "") {
        Write-Host ""
        Write-Output "`n" | Out-File $RunLog -Append }
}

function createContactCard($ContactName, $internalAddress, $forwardingDetails){
    if((Get-MailContact $ContactName -erroraction 'silentlycontinue')-eq $null){
    $null = New-MailContact -Name $ContactName -ExternalEmailAddress $forwardingDetails
    Write-Log "SUCCESS: Contact $ContactName was created"
    }else{
    Write-Log "INFO: The contact $ContactName already exists, we will attempt to use that."
    }

    Set-MailContact $ContactName -EmailAddresses @{Add="smtp:$($internalAddress)"}
    Write-Log "SUCCESS: Set the emailaddresses to SMTP: $forwardingDetails, smtp: $internalAddress"

}

<#
csv will need to have a field "DisplayName, ForwardingDetails"

#>


<#----------Start of Script actions -----------#>

# --- Initialize log files


$TimeStamp = Get-Date -Format yyMMddhhmmss
$RunLog = $TimeStamp + "_ConvertDistroToContact_RunLog.txt"
$ErrorLog = $TimeStamp + "_ConvertDistroToContact_ErrorLog.txt"

# --- Initialize Variables
$DistroGroupDetails = @() #used to pass information between functions

# --- start processing
if(Test-Path $ImportCsvFile){
    Connect-ExchangeOnline
    $groups = Import-Csv -Path $ImportCsvFile 
    foreach ($group in $groups){
        if($group.ForwardingDetails -ne ""){
            $DistroGroupDetails = retrieveDistroData $group.DisplayName
            Write-Log "INFO: Data retrieved $DistroGroupDetails"
            if($DistroGroupDetails.PrimarySmtpAddress -ne $null){
                deleteDistroList $group.DisplayName
                createContactCard $DistroGroupDetails.DisplayName $DistroGroupDetails.PrimarySmtpAddress $group.ForwardingDetails
            }else{
            Write-Log "ERROR: The $($group.DisplayName) could not be deleted and the contact card could not be created."
            }
        }else{Write-Log "ERROR: The Distribution list $($group.DisplayName) does not have a ForwardingDetails listed "}
    }

}
else{
Write-Log "ERROR: Input file not found. [$ImportCsvFile]"
}

