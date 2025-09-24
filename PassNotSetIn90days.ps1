
$excludedOUs = @(
   
)

$allUsers = Get-ADUser -Filter "(ObjectClass -eq 'user') -and (Enabled -eq 'True')" -SearchBase 'OU=BHSI,DC=bhsi,DC=berksi,DC=com' -SearchScope SubTree -Properties PasswordLastSet

$filteredList = $allUsers | Where-Object {
    $userOU = $_.DistinguishedName -replace [regex]::escape('\,')
    $userOu = $userOU -replace '^CN=.*?,'  
    Write-Host $userOU
    -not ($excludedOUs -contains $userOU)
} 

$filteredList = $filteredList | Where-Object {$_.passwordlastset -lt (Get-Date).AddDays(-90)} | Select-Object UserPrincipalName, PasswordLastSet,Enabled, DistinguishedName

$filteredList | Export-Csv .\PassOver90Days.csv -NoTypeInformation
