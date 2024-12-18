<#
 .SYNOPSIS
 Azure Automation script that adds and updates Service Accounts from Active Directory into Freshservice CMDB.

 .DESCRIPTION
 Collects the following information from Freshservice and AD to add or update Service Accounts into Freshservice CMDB.

 From Freshservice:
    - All Assets
    - Requesters
    - Agents
    - Departments
    - Service Accounts

 From AD:
    - User Manager
    - User PasswordLastSet
    - USer Description
    - User WhenCreated
    - User UPN
    - User SAM
    - User Department
    - User First Name
    - User Last Name

 .EXAMPLE
 This script is non-interactive with no supported parameters
 
 .INPUTS
 This script contains no predefined inputs

 .OUTPUTS
 This script contains no additional outputs
#>

<#
 Updates:
	Version     Date		 Updated by			 Description
	1.0			12.04.24     Jake Darcy         Initial Script 
#>

$StartTime = Get-Date

Write-Output "*** Starting Execution ***"
##############################
##    FS Asset Variables    ##
##############################
$Fs_ServiceAccountAssetID = ''
$Fs_AssetState = ''

##############################
##    Create Empty Arrays   ##
##############################
$Errors = @()
$Added = @()
$Updated = @()

#######################################
##    Set Connections & Variables    ##
#######################################
# Load Freshservice credentials from Automation Account
try {
    Write-Output "- Getting Freshservice credential..."
    $FS_Cred = Get-AutomationPSCredential -Name "FS_API_Automation"
    $FS_URL = Get-AutomationVariable -Name "FS_URL"
}
catch {
    Write-Warning ("WARNING: Failed to get Freshservice credentials from Automation account:" + $_)
    Write-Output "- Loading Freshservice credentials from server locally..."
    # Load encrypted credentials for Freshservice
    $FS_KeyFile = "C:\scripts\creds\"
    $FS_Key = Get-Content $FS_KeyFile
    $FS_Encrypted_Key = Get-Content "C:\Scripts\Creds\" | ConvertTo-SecureString -Key $FS_Key
    $FS_Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $FS_Encrypted_Key, $FS_Encrypted_Key
    $FS_URL = "https://domain.freshservice.com"
}

# Set FreshService Connection variables
$FS_APIKey = $FS_Cred.GetNetworkCredential().password
$FS_EncodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $FS_APIKey,$null)))
$FS_HTTPHeaders = @{}
$FS_HTTPHeaders.Add('Authorization', ("Basic {0}" -f $FS_EncodedCredentials))
$FS_HTTPHeaders.Add('Content-Type', 'application/json')

# Load Email credentials from Automation Account
try {
    Write-Output "- Getting Email credential..."
    $EMAIL_Cred = Get-AutomationPSCredential -Name "Automation Email"
}
catch {
    Write-Warning "WARNING: Failed to get Email credentials from Automation account:" 
    Write-Output $_
    Write-Output "- Loading Email credentials from server locally..."
    # Load encrypted credentials for svc_automation to send email
    $EMAIL_UN = ""
    $SA_KeyFile = "C:\scripts\creds\"
    $SA_Key = Get-Content $SA_KeyFile
    $SA_Encrypted_Key = Get-Content "C:\Scripts\Creds\" | ConvertTo-SecureString -Key $SA_Key
    $EMAIL_Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $EMAIL_UN, $SA_Encrypted_Key
}

####################################
##   Get Data from Freshservice   ##
####################################
#============#
# Requesters #
#============#
Write-Output "- Getting requesters from Freshservice..."
$URL_Requesters = "$FS_URL/api/v2/requesters?per_page=100"
# array for the retuned requesters to go into 
$Requesters = @()
do {
    try {
        $FS_Requesters = Invoke-WebRequest -Method GET -Headers $FS_HTTPHeaders -Uri $URL_Requesters -UseBasicParsing -ErrorAction SilentlyContinue
    }
    catch {
        Write-Error "StatusCode:" $_.Exception.Response.StatusCode.value 
        Write-Error "StatusDescription:" $_.Exception
    } 
    $RequesterObjects = $FS_Requesters.Content | ConvertFrom-Json
    $Requesters += $RequesterObjects.requesters
 
    # get next page of objects
    $URL_Requesters = [regex]::Match($FS_Requesters.Headers.link, '\<(.*)\>').Groups[1].Value
}
while ($URL_Requesters)

if ($Requesters) {
    Write-Output "- Received $($Requesters.count) reqeusters from Freshservice"
}
else {
    Write-Error "ERROR: Failed to retreive requesters from Freshservice"
}


#========#
# Agents #
#========#
Write-Output "- Getting agents from Freshservice..."
$URL_Agents = "$FS_URL/api/v2/agents?per_page=100"
# array for the retuned agents to go into 
$Agents = @()
do {
    try {
        $FS_Agents = Invoke-WebRequest -Method GET -Headers $FS_HTTPHeaders -Uri $URL_Agents -UseBasicParsing -ErrorAction SilentlyContinue
    }
    catch {
        Write-Error "StatusCode:" $_.Exception.Response.StatusCode.value 
        Write-Error "StatusDescription:" $_.Exception
    } 
    $AgentObjects = $FS_Agents.Content | ConvertFrom-Json
    $Agents += $AgentObjects.agents
 
    # get next page of objects
    $URL_Agents = [regex]::Match($FS_Agents.Headers.link, '\<(.*)\>').Groups[1].Value
}
while ($URL_Agents)

if ($Agents) {
    Write-Output "- Received $($Agents.count) agents from Freshservice"
}
else {
    Write-Error "ERROR: Failed to retreive agents from Freshservice"
}

#=============#
# Departments #
#=============#
Write-Output "- Getting departments from Freshservice..."
$URL_Departments = "$FS_URL/api/v2/departments"
# array for the retuned departments to go into 
$Departments = @()
do {
    try {
        $FS_Departments = Invoke-WebRequest -Method GET -Headers $FS_HTTPHeaders -Uri $URL_Departments -UseBasicParsing -ErrorAction SilentlyContinue
    }
    catch {
        Write-Output "- StatusCode:" $_.Exception.Response.StatusCode.value 
        Write-Output "- StatusDescription:" $_.Exception
    } 
    $DepartmentObjects = $FS_Departments.Content | ConvertFrom-Json
    $Departments += $DepartmentObjects.departments
 
    # get next page of objects
    $URL_Departments = [regex]::Match($FS_Departments.Headers.link, '\<(.*)\>').Groups[1].Value
}
while ($URL_Departments)

if ($Departments) {
    Write-Output "- Received $($Departments.count) departments from Freshservice"
}
else {
    Write-Error "ERROR: Failed to retreive departments from Freshservice"
}

##########################
##  Service Accounts    ##
##########################

Write-Output "- Getting service accounts from Freshservice..."
$URL_Assets = "$FS_URL/api/v2/assets?&include=type_fields&query="
# array for the retuned service accounts to go into 
$FS_SVCAccounts = @()
do {
    try {
        $FS_ServiceAccounts = Invoke-WebRequest -Method GET -Headers $FS_HTTPHeaders -Uri $URL_Assets -UseBasicParsing -ErrorAction SilentlyContinue
    }
    catch {
        Write-Error "StatusCode:" $_.Exception.Response.StatusCode.value 
        Write-Error "StatusDescription:" $_.Exception
    } 
    $ServiceAccountObjects = $FS_ServiceAccounts.Content | ConvertFrom-Json
    $FS_SVCAccounts += $ServiceAccountObjects.assets
 
    # get next page of objects
    $URL_Assets = [regex]::Match($FS_ServiceAccounts.Headers.link, '\<(.*)\>').Groups[1].Value
}
while ($URL_Assets)

if ($FS_SVCAccounts) {
    Write-Output "- Received $($FS_SVCAccounts.count) Service Accounts from Freshservice"
}
else {
    Write-Error "ERROR: Failed to retreive Servicee Accounts from Freshservice"
}

#####################################################################################################################################################

##########################################
##  Get Service Account Info From AD    ##
##########################################

$ServiceAccountsAD = @()
$ServiceAccountsAD += Get-ADUser -Filter 'enabled -eq $true' -SearchBase "OU=" -SearchScope Subtree -Properties manager, passwordlastset, Description, whenCreated | Select-Object manager, distinguishedname, GivenName, passwordlastset, UserPrincipalName, SamAccountName, description, whenCreated

#$ServiceAccountsAD = $null
#Uncomment for testing a single account and comment out the line above
#$ServiceAccountsAD += Get-ADUser -Filter "(UserPrincipalName -eq '')" -Properties manager, passwordlastset, Description, whenCreated | Select-Object manager, distinguishedname, GivenName, passwordlastset, UserPrincipalName, SamAccountName, description, whenCreated

$ServiceAccountADInfo = @()
foreach ($Account in $ServiceAccountsAD){
    $SAN = $Account.SamAccountName
    $LastPassReset = $Account.passwordlastset
    $Description = $Account.description
    $CreatedDate = $Account.whenCreated
    $Manager = (Get-ADUser (Get-ADUser $SAN -properties manager).manager).UserPrincipalName
    $ManagerFirstName = (Get-ADUser (Get-ADUser $SAN -Properties manager).manager).GivenName  
    $ManagerLastName = (Get-ADUser (Get-ADUser $SAN -Properties manager).manager).Surname
    $Department = (Get-ADUser -Properties department(Get-ADUser $SAN -Properties manager).manager).department


    # set attributes from FS
    $FS_Manager = $FS_SVCAccounts | ?{$Manager -eq $_.primary_email}
    $FS_Manager_ID = $FS_Manager.id
        if (!$Manager) {$FS_Manager_ID = 0000000}
        elseif ($FS_Manager.primary_email -notcontains $Manager){
                $FS_Manager = $Agents | ?{$Manager -eq $_.email}
                $FS_Manager_ID = $FS_Manager.id}
    $FS_Department = $Departments | ?{$Department -eq $_.name}
    $FS_Department_ID = $FS_Department.id
        if (!$FS_Department_ID) {$FS_Department_ID = 0000000}    

    $Properties = [Ordered]@{
        'AssetType'         ="Service Account"
        'SamAccountName'    =$SAN
        'ManagerFirstName'  =$ManagerFirstName
        'ManagerLastName'   =$ManagerLastName
        'LastPasswordReset' =$LastPassReset
        'Description'       =$Description
        'CreatedDate'       =$CreatedDate
        'Department'        =$Department
        'Manager'           =$Manager
        'ManagerID'         =$FS_Manager_ID
        'DepartmentID'      =$FS_Department_ID
        
    }
    $Object = New-Object -TypeName PSCustomObject -Property $Properties
    $ServiceAccountADInfo += $Object


    $CreatedDate = $null
    $Department = $null
    $ManagerFirstName = $null
    $ManagerLastName = $null
    $LastPassReset = $null
    $Description = $null
    $SAN = $null
    $Manager = $null
    $FS_Manager_ID = $null
    $FS_Department_ID = $null
}


##############################################
##  Add Service Accounts to Freshservice    ##
##############################################

$Url_Add = "$FS_Url/api/v2/assets"
$ServiceAccountADInfo | ForEach-Object{
    $SVC_Name = $_.SamAccountName
    $SVC_LastPassReset = $_.LastPasswordReset
    $SVC_Description = $_.Description
    $SVC_CreatedDate = $_.CreatedDate
    $FS_ManagerID = $_.ManagerID
    $FS_DeptID = @($_.DepartmentID)


    #make array of all service accounts in freshservice before this point
    if(($FS_SVCAccounts.name -NotContains $SVC_Name)){
        Write-Output "- $SVC_Name will be added to Freshservice"
        $Attributes = @{}
        $SubAttributes = @{}
        
        $Attributes.add('type_fields', $SubAttributes)
        $Attributes.add('name', $SVC_Name)
        $Attributes.add('asset_type_id', [int64]$Fs_ServiceAccountAssetID)
        $Attributes.add('description', $SVC_Description)
        $Attributes.add('agent_id', [int64]$FS_ManagerID)
        $Attributes.add('department_id', [int64]$FS_DeptID)

        $SubAttributes.add($Fs_AssetState, 'Enabled')
        $SubAttributes.add('original_created_at_', $SVC_CreatedDate.ToString("yyyy-MM-ddTHH:mm:ssZ"))
        $SubAttributes.add('last_password_reset_', $SVC_LastPassReset.ToString("yyyy-MM-ddTHH:mm:ssZ"))

        Write-Output "Debug values:"
        Write-Output "SVC_Name: $SVC_Name"
        Write-Output "Fs_ServiceAccountAssetID: $Fs_ServiceAccountAssetID"
        Write-Output "SVC_Description: $SVC_Description"
        Write-Output "FS_ManagerID: $FS_ManagerID"
        Write-Output "FS_DeptID: $FS_DeptID"
        Write-Output "FS_AssetState: $Fs_AssetState"
        Write-Output "SVC_CreatedDate: $SVC_CreatedDate"
        Write-Output "SVC_LastPassReset: $SVC_LastPassReset"

        $JSON = $Attributes | ConvertTo-Json -Depth 10
        Write-Output "- JSON used to add the Service Account:" 
        Write-Output $JSON
      
     
 

        # Attempt to make the API call and add the Service account from AD
        try {
            Write-Output "`nAttempting to create asset for $SVC_Name..."
    
            $response = Invoke-WebRequest -Method Post `
                                        -Uri $URL_Add `
                                        -Headers $FS_HTTPHeaders `
                                        -Body $JSON `
                                        -ContentType 'application/json' `
                                        -UseBasicParsing
    
            Write-Output "Success! Response Status: $($response.StatusCode)"
            Write-Output "Response Content:"
            Write-Output $response.Content
        } catch {
            Write-Output "`nError occurred during API call:"
            Write-Output "Status Code: $($_.Exception.Response.StatusCode.value__)"
    
            # Get the error response body
            $errorResponse = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $errorBody = $reader.ReadToEnd()
            Write-Output "`nError details from server:"
            Write-Output $errorBody
        }

       Write-Output "- Successfully created $SVC_Name"

        # clear variables
        $JSON = $Attributes = $SubAttributes = $null

     }
}