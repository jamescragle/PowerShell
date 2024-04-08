# Load the necessary .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# Create an OpenFileDialog object
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = 'CSV files (*.csv)|*.csv'  # Only show CSV files
}

# Display the dialog box
$null = $FileBrowser.ShowDialog()

# Access the selected file name (if needed)
$selectedCsvFile = $FileBrowser.FileName
