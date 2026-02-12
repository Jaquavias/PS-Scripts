Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force


if (-not (Get-Module -Name Microsoft.Graph.Users -ListAvailable)) {
    Install-Module Microsoft.Graph.Users -Force -AllowClobber
}

Import-Module Microsoft.Graph.Users

Connect-MgGraph -Scopes "AuditLog.Read.All"

$Properties = @('DisplayName', 'UserPrincipalName', 'SignInActivity', "Id")

$AllUsers = Get-MgUser -Filter "userType eq 'Guest'" -All -Property $Properties

Write-Host $AllUsers

$AllUsers | ForEach-Object {
    $ID = $_.Id
    [pscustomobject]@{
        LastLoginDate = $_.SignInActivity.LastSignInDateTime
        DisplayName = $_.DisplayName
        UPN = $_.UserPrincipalName
        Manager = (Get-MgUserManager -UserId $ID).Id
    }

}

$AllUsers | Export-Csv .\Temp88.csv -NoTypeInformation