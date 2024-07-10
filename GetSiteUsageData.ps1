## SharePoint Server 2013: PowerShell Script to Produce a CSV Report on Web Analytics Data From Web Application And / Or Site Collection Level ##

#------------------------------------------------------------------------------------------- 
# Name:            ExportRollupAnalyticsData
# Description:     This script will export SharePoint 2013 Web Analytics data to a CSV file
# Usage:           Run the function with the required parameters.  
#                  Scope can be all SPSite and / or all SPWeb objects in a Web Application
# Author:          Chris LaQuerre | http://sp2013wade.codeplex.com
#
# Reference:       SearchServiceApplicationProxy.GetRollupAnalyticsItemData method parameters
#                  http://msdn.microsoft.com/en-us/library/office/microsoft.office.server.search.administration.searchserviceapplicationproxy.getrollupanalyticsitemdata.aspx
#                  http://sp2013wade.codeplex.com
#
#                  eventType
#                  Type: System.Int32
#                  the event type, e.g. 1 for Site Usage Reports
#
#                  tenantId
#                  Type: System.Guid
#                  leave blank
#
#                  siteId
#                  Type: System.Guid
#                  the site collection id
#                  
#                  scopeId
#                  Type: System.Guid
#                  the scope id, e.g. the web id for View events, or Guid.Empty for the entire site collection
#
# Inspiration:     http://www.sharepointtalk.net/2014/02/query-sharepoint-search-analytics-using.html
#                  http://gallery.technet.microsoft.com/office/Get-SharePoint-Web-19cd2137 (Ivan Josipovic)
#------------------------------------------------------------------------------------------- 

# program start at bottom

Add-PSSnapin "Microsoft.SharePoint.PowerShell"

function ExportRollupAnalyticsData {
    Param(
    [string]$RootSiteUrl,
    [string]$Scope,
    [switch]$IncludeSites,
    [switch]$IncludeWebs,
    [string]$OutputFilePath
    )


    # Get Web Application for Root Site
    $Site = Get-SPSite $RootSiteUrl
    

    # Get Search Service Application
    $SearchApp = Get-SPEnterpriseSearchServiceApplication

           # Export Site analtyics if -IncludeSites flag is present
        If ($IncludeSites.IsPresent) {
            $Scope = "Site"
            $SiteTitle = $Site.RootWeb.Title.Replace(",", "")
            $SiteUrl = $Site.Url
            $UsageData = $SearchApp.GetRollupAnalyticsItemData(1,[System.Guid]::Empty,$Site.ID,[System.Guid]::Empty)
            $LastProcessingTime = $UsageData.LastProcessingTime
            $CurrentDate = $UsageData.CurrentDate
            $TotalHits = $UsageData.TotalHits
            $TotalUniqueUsers = $UsageData.TotalUniqueUsers
            $LastProcessingHits = $UsageData.LastProcessingHits
            $LastProcessingUniqueUsers = $UsageData.LastProcessingUniqueUsers

            # Write Web details to CSV File
            $OutputString = $Scope + "," + $SiteTitle + "," + $SiteUrl + "," + $LastProcessingTime + "," + $TotalHits + "," + $TotalUniqueUsers + "," + $LastProcessingHits + "," + $LastProcessingUniqueUsers + "," + $CurrentDate
            $OutputString | Out-File $OutputFilePath -Append
        }      

        # Export Web analtyics if -IncludeWebs flag is present
        If ($IncludeWebs.IsPresent) {
            
            # Loop through all Webs in Site Collection
            ForEach($Web in $Site.AllWebs) {
                $Scope = "Web"
                $WebTitle = $Web.Title.Replace(",", "")
                $WebUrl = $Web.Url
                $UsageData = $SearchApp.GetRollupAnalyticsItemData(1,[System.Guid]::Empty,$Site.ID,$Web.ID)
                $LastProcessingTime = $UsageData.LastProcessingTime
                $CurrentDate = $UsageData.CurrentDate
                $TotalHits = $UsageData.TotalHits
                $TotalUniqueUsers = $UsageData.TotalUniqueUsers
                $LastProcessingHits = $UsageData.LastProcessingHits
                $LastProcessingUniqueUsers = $UsageData.LastProcessingUniqueUsers

                # Write Web details to CSV File
                $OutputString = $Scope + "," + $WebTitle + "," + $WebUrl + "," + $LastProcessingTime + "," + $TotalHits + "," + $TotalUniqueUsers + "," + $LastProcessingHits + "," + $LastProcessingUniqueUsers + "," + $CurrentDate
                $OutputString | Out-File $OutputFilePath -Append 
            }
        }
         
        # Dispose Site Collection Object
        $Site.Dispose()
    
}

# ---------------- program start ---------------------

# Load the necessary .NET assembly
Add-Type -AssemblyName System.Windows.Forms
# where do you want to save the file to? appending the current date to csv file save name
$FileSaveLocation = "c:\temp\output" + (Get-Date -format "yyyyMMdd") + ".csv"

# Delete CSV file if existing
  If (Test-Path $FileSaveLocation) {
	 Remove-Item $FileSaveLocation
   }


# Write header row to CSV File
$OutputHeader = "Scope,Name,URL,Most Recent Day with Usage,Hits - All Time,Unique Users - All Time,Hits - Most Recent Day with Usage,Unique Users - Most Recent Day with Usage,Current Date"
$OutputHeader | Out-File $FileSaveLocation -Append 

# Create an OpenFileDialog object
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = 'CSV files (*.csv)|*.csv'  # Only show CSV files
}

# Display the dialog box
$null = $FileBrowser.ShowDialog()

# Access the selected file name (if needed)
$selectedCsvFile = $FileBrowser.FileName

$CsvFile = Import-Csv $selectedCsvFile

foreach ($site in $CsvFile){
    ExportRollupAnalyticsData -RootSiteUrl $site.SiteURL -OutputFilePath $FileSaveLocation -IncludeSites -IncludeWebs
}

# Sample Usage: Export both SPSite and SPWeb analytics 
#ExportRollupAnalyticsData -RootSiteUrl "http://sitecollection" -OutputFilePath "C:\temp\analytics-export-sites-webs.csv" -IncludeSites -IncludeWebs

# Sample Usage: Export only SPSite analytics
#ExportRollupAnalyticsData -RootSiteUrl "http://sitecollection" -OutputFilePath "C:\temp\analytics-export-sites-only.csv" -IncludeSites

# Sample Usage: Export only SPWeb analytics
#ExportRollupAnalyticsData -RootSiteUrl "http://sitecollection" -OutputFilePath "C:\temp\analytics-export-webs-only.csv" -IncludeWebs
