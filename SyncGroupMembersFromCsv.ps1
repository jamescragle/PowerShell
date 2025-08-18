<#

USAGE INSTRUCTIONS

Prerequisites:
- Install the Microsoft Graph PowerShell SDK (Microsoft.Graph).
- Run PowerShell as an administrator.
- Ensure you are signed in with a Global Admin account.

How to Run:

Option 1: With CSV File Picker (GUI)
    .\SyncGroupMembersFromCsv.ps1
    (You will be prompted to select a CSV file with 'sourcevalue' and 'destinationvalue' columns.)

Option 2: With CSV Path Parameter
    .\SyncGroupMembersFromCsv.ps1 -CsvPath "C:\Path\To\YourFile.csv"
    (Skips the file picker and uses the provided CSV file.)

Test Mode:
    Add the -TestMode flag to simulate all actions without committing any changes:
    .\SyncGroupMembersFromCsv.ps1 -CsvPath "C:\Path\To\YourFile.csv" -TestMode
    (All actions will be logged, but no changes will be made to groups.)

What the Script Does:
- Connects to Microsoft Graph.
- Processes all Microsoft 365 groups.
- For each group, checks the CSV for matching members and adds new members as needed.
- Logs all activities and errors to a 'logs' folder in the script’s directory.

Log Files:
- All logs: logs\<scriptname>_all.txt
- Error logs: logs\<scriptname>_errors.txt
#>

param(
    [string]$CsvPath,
    [switch]$TestMode
)


# Function to show a GUI file picker for CSV files
Function Get-CSVFilePath {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "CSV files (*.csv)|*.csv"
    $dialog.Title = "Select CSV File"
    $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    } else {
        throw "No CSV file selected. Script will exit."
    }
}
# Function to write to logs
# example usage Write-Log "Please specify either a local upload path or a folder structure JSON path, not both." -output true -color Red
# exmple using in a catch Write-Log "Failed to load mock employee data. Exits script." -errorMessage "Invalid input data." -output true -color Red 
Function Write-Log { param ([string]$message, [string]$errorMessage = $null, [Exception]$exception = $null, [string]$output = $false, [string]$color = "Green")

    # Define log level - Can be "errors" or "all"
    $logLevel = "all"

    # Create logs directory if it doesn't exist
    if (-not (Test-Path ".\logs")) {
        New-Item -Path . -Name "logs" -ItemType 'directory' > $null
    }

    $dateTime = Get-Date

    # Set log filename to the name of the script
    $logFilename = Get-Script-Name
    $debugErrorFile = ".\logs\" + $logFilename + "_errors.txt"
    $debugAllFile = ".\logs\" + $logFilename + "_all.txt"

    if ($exception -or $errorMessage) {
        $severity = "ERROR"
    } else {
        $severity = "INFO"
    }

    if ($exception.Response) {
        $result = $exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($result)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
    }

    $logMessage = ($severity + "`t")
    $logMessage += ($dateTime)
    $logMessage += ("`t" + $message + "`t")

    if ($exception) {
        $logMessage += ($exception.Message + "`t")
    }

    if ($errorMessage) {
        $logMessage += ($errorMessage + "`t")
    }

    if ($responseBody) {
        $logMessage += ("Box responded with: " + $responseBody + "`t")
    }

    if ($output -eq "true") {
        Write-Host $message -ForegroundColor $color
    }

    if ($logLevel -eq "all") {
        $logMessage | Add-Content $debugAllFile

        if ($severity -eq "ERROR") {
            $logMessage | Add-Content $debugErrorFile
        }
    } else {
        if ($severity -eq "ERROR") {
            $logMessage | Add-Content $debugErrorFile
        }
    }
}

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read.All"

# Get all Microsoft 365 groups (Unified groups)
$groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All

write-host "checking login status"

# Determine CSV path and load rows once
if ($CsvPath -and (Test-Path $CsvPath)) {
    Write-Log "Using provided CSV path: $CsvPath" -output true
    $csvRows = Import-Csv -Path $CsvPath
} else {
    $selectedCsvPath = Get-CSVFilePath
    Write-Log "Using selected CSV path: $selectedCsvPath" -output true
    $csvRows = Import-Csv -Path $selectedCsvPath
}

foreach ($group in $groups) {

    Write-Host "Processing group: $($group.DisplayName)"
    Write-Log "Processing group: $($group.DisplayName)" -output true

    # Get members of the group
    $members = Get-MgGroupMember -GroupId $group.Id -All | Where-Object { $_.ODataType -eq "#microsoft.graph.user" }


        # Get current group member UPNs
        $memberUPNs = $members | ForEach-Object { $_.UserPrincipalName }

        # Get current group member IDs for skip logic
        $memberIds = $members | ForEach-Object { $_.Id }

        foreach ($row in $csvRows) {
            $sourceUPN = $row.sourcevalue
            $destinationEmail = $row.destinationvalue

            # Check if sourceUPN is a member of the group
            if ($memberUPNs -contains $sourceUPN) {
                # Get destination user object by email
                $destinationUser = Get-MgUser -Filter "mail eq '$destinationEmail'"
                if ($destinationUser) {
                    # Skip if already a member
                    if ($memberIds -contains $destinationUser.Id) {
                            Write-Host "$destinationUPN is already a member of $($group.DisplayName), skipping."
                            Write-Log "$destinationUPN is already a member of $($group.DisplayName), skipping." -output true
                    } else {
                            Write-Host "Adding $destinationEmail to group $($group.DisplayName)"
                            Write-Log "Adding $destinationEmail to group $($group.DisplayName)" -output true
                            if ($TestMode) {
                                Write-Host "[TEST MODE] No changes committed."
                                Write-Log "[TEST MODE] Skipped actual add for $destinationEmail to $($group.DisplayName)" -output true -color Yellow
                            } else {
                                try {
                                    Add-MgGroupMember -GroupId $group.Id -DirectoryObjectId $destinationUser.Id
                                    Write-Log "$destinationEmail successfully added to $($group.DisplayName)" -output true
                                } catch {
                                    Write-Warning "Failed to add $destinationEmail : $_"
                                    Write-Log "Failed to add $destinationEmail to $($group.DisplayName)" -errorMessage $_.Exception.Message -exception $_.Exception -output true -color Red
                                }
                            }
                    }
                } else {
                        Write-Warning "Destination user $destinationUPN not found."
                        Write-Log "Destination user $destinationUPN not found." -output true -color Yellow
                }
            }
        }
}
