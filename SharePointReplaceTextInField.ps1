<#
Borrowed pieces of this script from someone who already did the hard work of filtering html.
In my case, someone converted a HTML multi line of text field to plain text
which brought over all the html encoding. 
you can update the "convertfrom-html" function to whatever text replacement you want. 
This could also work if you need to update text, say a business name in every item in a list.
The data table here captures the before and after text of the item.
You will need PnP Powershell to run this script. 

#>
Function ConvertFrom-Html
{
    <#
        .SYNOPSIS
            Converts a HTML-String to plaintext.

        .DESCRIPTION
            Creates a HtmlObject Com object und uses innerText to get plaintext. 
            If that makes an error it replaces several HTML-SpecialChar-Placeholders and removes all <>-Tags via RegEx.

        .INPUTS
            String. HTML als String

        .OUTPUTS
            String. HTML-Text als Plaintext

        .EXAMPLE
        $html = "<p><strong>Nutzen:</strong></p><p>Der&nbsp;Nutzen ist &uuml;beraus gro&szlig;.<br />Test ob 3 &lt; als 5 &amp; &quot;4&quot; &gt; &apos;2&apos; it?"
        ConvertFrom-Html -Html $html
        $html | ConvertFrom-Html

        Result:
        "Nutzen:
        Der Nutzen ist überaus groß.
        Test ob 3 < als 5 ist & "4" > '2'?"


        .Notes
            Author: Ludwig Fichtinger FILU
            Inital Creation Date: 01.06.2021
            ChangeLog: v2 20.08.2021 try catch with replace for systems without Internet Explorer

    #>

     [CmdletBinding(SupportsShouldProcess = $True)]
    Param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, HelpMessage = "HTML als String")]
        [AllowEmptyString()]
        [string]$Html
    )

        $nl = [System.Environment]::NewLine
        $PlainText = $Html -replace '<br>',$nl
        $PlainText = $PlainText -replace '<br/>',$nl
        $PlainText = $PlainText -replace '<br />',$nl
        $PlainText = $PlainText -replace 'br',$nl
        $PlainText = $PlainText -replace '</p>',$nl
        $PlainText = $PlainText -replace '&nbsp;',' '
        $PlainText = $PlainText -replace '&Auml;','Ä'
        $PlainText = $PlainText -replace '&auml;','ä'
        $PlainText = $PlainText -replace '&Ouml;','Ö'
        $PlainText = $PlainText -replace '&ouml;','ö'
        $PlainText = $PlainText -replace '&Uuml;','Ü'
        $PlainText = $PlainText -replace '&uuml;','ü'
        $PlainText = $PlainText -replace '&szlig;','ß'
        $PlainText = $PlainText -replace '&amp;','&'
        $PlainText = $PlainText -replace '&quot;','"'
        $PlainText = $PlainText -replace '&apos;',"'"
        $PlainText = $PlainText -replace '<.*?>',''
        $PlainText = $PlainText -replace '&gt;','>'
        $PlainText = $PlainText -replace '&lt;','<'
        $PlainText = $PlainText -replace '<div>',' '
        $PlainText = $PlainText -replace '/div',' '
        $PlainText = $PlainText -replace 'div',' '


    return $PlainText
}

#create a datatable for tracking
$dataTable = New-Object system.data.datatable

$col1 = new-object system.data.datacolumn("ItemID")
$col2 = new-object system.data.datacolumn("BeforeText")
$col3 = new-object system.data.datacolumn("AfterText")


$dataTable.Columns.Add($col1)
$dataTable.Columns.Add($col2)
$dataTable.Columns.Add($col3)

#configure variables

$SiteURL = "sharepoint online site url"
$ListName ="sharepoint online list name"
$listField = "Description_x0020_of_x0020_Reque" #this is the internal field name (example shown)

Connect-PnPOnline -Url $SiteURL -Interactive

#get all list items

$items = (Get-PnPListItem -List $ListName -Fields $listField )

#loop through items and set them

for($i = 0; $i -le $items.count; $i++){
   $newtext = ConvertFrom-Html -Html $items[$i].FieldValues[$listField] 
   $row = $dataTable.NewRow()
    $row["ItemID"] = $items[$i].Id
    $row["BeforeText"] = $items[$i].FieldValues[$listField]
    $row["AfterText"] = $newtext
    $dataTable.Rows.Add($row)
    Set-PnPListItem -List $ListName -Identity $items[$i].ID -Values @{$listField = $newtext} -UpdateType SystemUpdate
}




