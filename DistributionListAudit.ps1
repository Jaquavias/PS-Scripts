Connect-ExchangeOnline

$Distro = Get-DistributionGroup -ResultSize Unlimited | Format-Table DisplayName, PrimarySMTPAddress, ManagedBy 
$Distro | Export-Csv .\Temp3.csv -NoTypeInformation

$allGroups = foreach ($grp in Get-DistributionGroup -ResultSize Unlimited){
        [PsCustomObject]@{
            DisplayName        = $grp.DisplayName
            PrimarySMTPAddress = $grp.PrimarySMTPAddress
            DistinguishedName  = $grp.DistinguishedName
            ManagedBy          = $grp.ManagedBy
        }    
}
$allGroups | Export-Csv -Path .\Temp2.csv -NoTypeInformation

$emptyGroups = foreach ($grp in Get-DistributionGroup -ResultSize Unlimited){
    if (@(Get-DistributionGroupMember -Identity $grp.DistinguishedName -ResultSize Unlimited).Count -eq 0){
        [PsCustomObject]@{
            DisplayName        = $grp.DisplayName
            PrimarySMTPAddress = $grp.PrimarySMTPAddress
            DistinguishedName  = $grp.DistinguishedName
        }    
    }
}

$emptyGroups | Export-Csv -Path .\Temp1.csv -NoTypeInformation

#Get All Distribution Groups
$DistributionGroups = Get-Distributiongroup -resultsize unlimited 
$DLS =@()
 
#Collect members of each distribution list
$DistributionGroups | ForEach-Object {
        $Group = $_
        Get-DistributionGroupMember -Identity $Group.Name -ResultSize Unlimited | ForEach-Object {
            $member = $_
            $DLS += [PSCustomObject]@{
            GroupName = $Group.DisplayName
            GroupEmail = $Group.PrimarySmtpAddress
            Member = $Member.Name
            EmailAddress = $Member.PrimarySMTPAddress
            RecipientType= $Member.RecipientType
            }
        }
    }
#Get Distribution List Members
$DLS | Export-Csv -Path .\Temp2.csv -NoTypeInformation


$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$workBook = $excel.Workbooks.open("CSV Path")
$sheet = $workbook.Worksheets.Item(1)  
$rowMax = $sheet.UsedRange.Rows.Count

for ($row = $rowMax; $row -ge 2; $row--){
    $domain = ($sheet.Cells.Item($row, 4).Value2) -replace ".*@"
    if ($domain -eq "domain"){
        $null = $sheet.Rows($row).EntireRow.Delete()
    }
}


$workBook.Save()

$workBook.Close()
