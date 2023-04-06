#create the data table

$dataTable = New-Object system.data.datatable


$col1 = new-object system.data.datacolumn("ItemID")
$col2 = new-object system.data.datacolumn("BeforeText")
$col3 = new-object system.data.datacolumn("AfterText")


$dataTable.Columns.Add($col1)
$dataTable.Columns.Add($col2)
$dataTable.Columns.Add($col3)


#use the data table

$row = $dataTable.NewRow()
$row["ItemID"] = "your value"
$row["BeforeText"] = "your value"
$row["AfterText"] = "your value"
$dataTable.Rows.Add($row)


#https://blog.russmax.com/powershell-using-datatables/