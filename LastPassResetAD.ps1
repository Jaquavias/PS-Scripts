<#
Synopsis - Collectes all enabled accounts from a certain OU and spits back out the name, last pass reset and manager
    Will also sort the last pass reset in descending order.
#>

$OU = "OU="

$Report = foreach ($DisName in Get-ADUser -Filter * -SearchBase $OU -SearchScope Subtree | Select-object DistinguishedName){
    
    $DisName = $DisName -replace [regex]::escape('@{DistinguishedName=')
    $DisName = $DisName -replace [regex]::escape('}')
    Get-ADUser -Filter "(DistinguishedName -eq '$DisName') -and (Enabled -eq 'true')" -Properties passwordlastset,manager 

}
$Report | Export-Csv -Path .\Temp1.csv -NoTypeInformation

$Report = Import-Csv -Path .\Temp1.csv


$newReport = foreach ($row in $Report) {
    $row | Select-Object *, @{Name = "ManagerName"; Expression = {$row.Manager -replace '^CN=|,\S.*$'}}
}

$newReport | Export-Csv -Path .\Temp2.csv -NoTypeInformation



Import-Csv -Path .\Temp2.csv | Select-Object GivenName,Name, PasswordLastSet, ManagerName | Export-Csv -Path .\FinalReport$(Get-Date -Format "yyyy-MMM-dd-HH-mm-ss").csv -NoTypeInformation

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$workBook = $excel.Workbooks.open("C:\workbook location")
#$newSheet = $workBook.Worksheets.Add()
#$newSheet.name = "Test"

$sheet = $workbook.ActiveSheet    
$rangeToSort = $sheet.Range("C1")
$order = [Microsoft.Office.Interop.Excel.xlSortOrder]::xlDescending
$sortOn = [Microsoft.Office.Interop.Excel.XlSortOn]::SortOnValues
$sortData = [Microsoft.Office.Interop.Excel.XlSortDataOption]::xlSortNormal
$header = [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes

$sheet.Sort.SortFields.Clear()
$sheet.Sort.SortFields.Add($rangeToSort, $sortOn, $order, $sortData)
$sheet.sort.setRange($sheet.UsedRange)  
$sheet.sort.header = $header
$sheet.sort.apply()


$workBook.Save()

$workBook.Close()

Import-Csv -Path .\Temp3.csv | Select-Object GivenName,Name, PasswordLastSet, ManagerName | Export-Csv -Path .\FinalReport$(Get-Date -Format "yyyy-MMM-dd-HH-mm-ss").csv -NoTypeInformation

