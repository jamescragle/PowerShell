# Init PowerShell Gui
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing



$form = new-object system.windows.forms.form
#x is first <--> y is second ^-v
$form.ClientSize = '480,300'
$form.StartPosition = "CenterScreen"
$form.Text = "Testing capture data from form"
$form.add_keydown({
    if($_.KeyCode -eq "Escape"){$form.Close()}
})


$title = New-Object system.windows.forms.label
$title.Location = New-Object System.Drawing.Point(20,20)
$title.Text = "this box is at 20,20"
$title.Width = 100

$choice = New-Object System.Windows.Forms.ComboBox
$choice.Width = 100
$choice.Location = New-Object System.Drawing.Point(20,40)
@('One','Two','Three') | ForEach-Object {[void]$choice.Items.Add($_)}

$returnbutton = New-Object System.Windows.Forms.Button
$returnbutton.Text = "Return Choice"
$returnbutton.Location = New-Object System.Drawing.Point(20,80)
$returnbutton.Width = 90
$returnbutton.Add_click({$form.close()})


$form.Controls.AddRange(@($title,$choice,$returnbutton))

[void]$form.ShowDialog()

#the objects (ie the text fields and drop downs) remain present after the form closes and are accessed like variables
Write-Host "You chose $($choice.SelectedItem)"

