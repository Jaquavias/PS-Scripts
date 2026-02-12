Connect-MgGraph -Scopes "AuditLog.Read.All", "User.Read.All"

$Properties = @('DisplayName', 'UserPrincipalName', 'SignInActivity', "Id")

$AllUsers = Get-MgUser -Filter 'accountEnabled eq true' -All -Property $Properties


# Shape the output
$report = $AllUsers | ForEach-Object {
    [pscustomobject]@{
        DisplayName                         = $_.DisplayName
        UserPrincipalName                   = $_.UserPrincipalName
        UserType                            = $_.UserType            # Member/Guest
        AccountEnabled                      = $_.AccountEnabled
        LastInteractiveSignIn               = $_.SignInActivity.LastSignInDateTime
        LastNonInteractiveSignIn            = $_.SignInActivity.LastNonInteractiveSignInDateTime
        LastSuccessfulSignIn                = $_.SignInActivity.LastSuccessfulSignInDateTime
        LastSignInRequestId                 = $_.SignInActivity.LastSignInRequestId
        LastNonInteractiveSignInRequestId   = $_.SignInActivity.LastNonInteractiveSignInRequestId
    }
}


# Export
$report | Export-Csv .\Temp1001.csv -NoTypeInformation