<#

USAGE INSTRUCTIONS

Prerequisites:
- Install the Microsoft Graph PowerShell SDK (Microsoft.Graph).
- Run PowerShell as an administrator.
- Ensure you are signed in with a Global Admin account.
- The log folder will be created in the same directory as the script. 

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

# Function to check for the latest version of the Microsoft.Graph module
Function Test-LatestGraphModule {
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Checking for Microsoft.Graph module..." -ForegroundColor Cyan
    $moduleName = "Microsoft.Graph"
    
    # Check if the module is installed
    $installedModule = Get-InstalledModule -Name $moduleName -ErrorAction SilentlyContinue
    
    if (-not $installedModule) {
        throw "The '$moduleName' module is not installed. Please run 'Install-Module -Name $moduleName' in an elevated PowerShell session and try again."
    }

    Write-Host "Found installed version: $($installedModule.Version)"
    
    # Find the latest version from the PowerShell Gallery
    try {
        Write-Host "Checking PowerShell Gallery for the latest version..."
        # Use -ErrorAction Stop to ensure the catch block is triggered on any error (e.g., no internet)
        $galleryModule = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop
        
        Write-Host "Latest version available: $($galleryModule.Version)"
        # Compare versions
        if ([version]$installedModule.Version -lt [version]$galleryModule.Version) {
            Write-Warning "A newer version of the '$moduleName' module is available ($($galleryModule.Version)). Consider running 'Update-Module -Name $moduleName' for the latest features and fixes."
        } else {
            Write-Host "You have the latest version of the '$moduleName' module." -ForegroundColor Green
        }
    } catch {
        Write-Warning "Could not connect to the PowerShell Gallery to check for the latest version. Will proceed with the installed version: $($installedModule.Version)."
    }
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
}
# Function to write to logs
# example usage Write-Log "Please specify either a local upload path or a folder structure JSON path, not both." -output true -color Red
# exmple using in a catch Write-Log "Failed to load mock employee data. Exits script." -errorMessage "Invalid input data." -output -color Red 
Function Write-Log {
    param (
        [string]$message,
        [string]$errorMessage = $null,
        [System.Exception]$exception = $null,
        [switch]$output,
        [string]$color = "Green"
    )

    # Create logs directory if it doesn't exist
    $logDir = Join-Path -Path $PSScriptRoot -ChildPath "logs"
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType 'directory' -ErrorAction SilentlyContinue > $null
    }

    $dateTime = Get-Date

    # Set log filename to the name of the script
    $scriptName = $MyInvocation.MyCommand.Name -replace '\..*$'
    $debugErrorFile = Join-Path -Path $logDir -ChildPath "${scriptName}_errors.txt"
    $debugAllFile = Join-Path -Path $logDir -ChildPath "${scriptName}_all.txt"

    if ($exception -or $errorMessage) {
        $severity = "ERROR"
    } else {
        $severity = "INFO"
    }
    
    $logMessage = ($severity + "`t")
    $logMessage += ($dateTime)
    $logMessage += ("`t" + $message + "`t")

    if ($exception) {
        $logMessage += ($exception.Message + "`t")
        # Check for Graph-specific error details
        if ($exception.ErrorDetails) {
            $logMessage += ("Graph Error: " + ($exception.ErrorDetails | ConvertTo-Json -Depth 3) + "`t")
        }
    }

    if ($errorMessage) {
        $logMessage += ($errorMessage + "`t")
    }

    if ($output.IsPresent) {
        Write-Host $message -ForegroundColor $color
    }

    $logMessage | Add-Content $debugAllFile
    if ($severity -eq "ERROR") {
        $logMessage | Add-Content $debugErrorFile
    }
}

# --- SCRIPT EXECUTION STARTS HERE ---
Test-LatestGraphModule

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read.All"
#Select-MgProfile -Name "v1.0"

# Get all Microsoft 365 groups (Unified groups)
$groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All

# Determine CSV path and load rows once
if (-not ($CsvPath)) {
    $CsvPath = Get-CSVFilePath
} 
elseif (-not (Test-Path $CsvPath)) {
    throw "CSV file not found at path: $CsvPath"
}

Write-Log "Using CSV path: $CsvPath" -output
$csvRows = Import-Csv -Path $CsvPath

# Cache for destination users to avoid repeated API calls
$destinationUserCache = @{}

foreach ($group in $groups) {

   # Write-Host "Processing group: $($group.DisplayName)"
    Write-Log "Processing group: $($group.DisplayName)" -output

    # Get members of the group
    $members = Get-MgGroupMember -GroupId $group.Id -All | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.user" }

    # Get current group member UPNs
    $memberUPNs = $members | ForEach-Object { $_.AdditionalProperties["userPrincipalName"] }

    # Get current group member IDs for skip logic
    $memberIds = $members | ForEach-Object { $_.Id }

    foreach ($row in $csvRows) {
        $sourceUPN = $row.sourcevalue
        $destinationEmail = $row.destinationvalue

        # Check if sourceUPN is a member of the group
        if ($memberUPNs -contains $sourceUPN) {
            
            # Get destination user object by email, using a cache to improve performance
            $destinationUser = $null
            if ($destinationUserCache.ContainsKey($destinationEmail)) {
                $destinationUser = $destinationUserCache[$destinationEmail]
            }
            else {
                try {
                    $destinationUser = Get-MgUser -Filter "mail eq '$destinationEmail'"
                    # Add user to cache, even if not found (cache $null) to prevent re-querying
                    $destinationUserCache[$destinationEmail] = $destinationUser
                }
                catch {
                    Write-Log "Error fetching user '$destinationEmail'." -errorMessage $_.Exception.Message -exception $_.Exception -output -color Red
                }
            }

            if ($destinationUser) {
                # Skip if already a member
                if ($memberIds -contains $destinationUser.Id) {
                        Write-Log "$($destinationUser.UserPrincipalName) is already a member of $($group.DisplayName), skipping." -output
                } else {
                        Write-Log "Adding $destinationEmail to group $($group.DisplayName)" -output
                        if ($TestMode) {
                            Write-Host "[TEST MODE] No changes committed."
                            Write-Log "[TEST MODE] Skipped actual add for $destinationEmail to $($group.DisplayName)" -output -color Yellow
                        } else {
                            try {
                                New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $destinationUser.Id
                                Write-Log "$destinationEmail successfully added to $($group.DisplayName)" -output
                            } catch {
                                Write-Log "Failed to add $destinationEmail to $($group.DisplayName)" -errorMessage $_.Exception.Message -exception $_.Exception -output -color Red
                            }
                        }
                }
            } else {
                    Write-Log "Destination user $destinationEmail not found." -output -color Yellow
            }
        }
    }
}

Write-Log "Script finished." -output
Disconnect-MgGraph
