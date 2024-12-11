


function Get-AccessToken{

    [CmdletBinding()]
    param (
        # The Microsoft Azure AD Tenant Name
        [Parameter(ParameterSetName = 'ClientAuth', Mandatory = $true)]  [string]$TenantName,
        # The Microsoft Azure AD TenantId (GUID or domain)
        [Parameter(ParameterSetName = 'ClientAuth', Mandatory = $true)]  [string]$ClientId,
         # An authentication secret of the Microsoft Azure AD Application Registration
         [Parameter(ParameterSetName = 'ClientAuth', Mandatory = $true)] [string]$ClientSecret
    )
    

    $resource = "https://graph.microsoft.com/"  
    
    $tokenBody = @{  
        Grant_Type    = 'client_credentials'  
        Scope         = 'https://graph.microsoft.com/.default'  
        Client_Id     = $ClientId  
        Client_Secret = $clientSecret  
    }  
    
    try{$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $tokenBody -ErrorAction Stop}catch{return "Error generating access token $($_)"}
    Write-Debug "Successfully generated authentication token"
    return $($tokenResponse.access_token)
    }


    

function Send-GraphRequest{
    
        param(
            [Parameter(Mandatory = $true)] [string]$AccessToken,
            [Parameter(Mandatory = $true)][ValidateSet("Get","Post","Patch","Delete")][string]$Method,
            [Parameter(Mandatory=$true)]
                [ValidateScript({
                    try {
                        [System.Uri]::new($_)
                        return $true
                    }
                    catch {
                        $false
                    }
                })]
                [string]$Uri,
            [Parameter(Mandatory = $false)] [string]$ContentType,
            [Parameter(Mandatory = $false)] [string]$ConsistencyLevel = $false,
            [Parameter(Mandatory = $false)] [string]$Body
        )
        
        BEGIN{
        
            function Get-AzureResourcePaging {
                param (
                    $URL,
                    $AuthHeader
                )
            
            
                $Response = Invoke-RestMethod -Method GET -Uri $URL -Headers $AuthHeader
                
                $Resources = $Response.value
                
            
                $ResponseNextLink = $Response."@odata.nextLink"
            
                while ($ResponseNextLink -ne $null) {
            
                    $Response = (Invoke-RestMethod -Uri $ResponseNextLink -Headers $AuthHeader -Method Get)
                    $ResponseNextLink = $Response."@odata.nextLink"
                    $Resources += $Response.value
                }
        
                if($Resources -eq $null)
                {
                    $Resources = $Response
                }
        
                return $Resources
            }
        
        
            if($ConsistencyLevel -eq $false)
            {
    
                $headers = @{
                    "Authorization" = "Bearer $($Accesstoken)"
                    "Content-type"  = "$ContentType"
                    "Accept"        = "$ContentType"
                }
            }
            else{
                $headers = @{
                    "Authorization" = "Bearer $($Accesstoken)"
                    "Content-type"  = "$ContentType"
                    "Accept"        = "$ContentType"
                    "ConsistencyLevel" = "eventual"
                }
            }
    
        
        
        $ErrorActionPreference = "Stop"
        
        }
        
        
        
        PROCESS
        {
        
        switch($Method){
            
                "Get"{
            
                    $output = Get-AzureResourcePaging -URL $Uri -AuthHeader $headers
        
                }
            
                "Post"{
                    
                    if($body -eq $null)
                    {
                        return "Body is required for POST method"
                    }
                    else{
                    $output = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $Body
                    }
                }
            
                "Patch"{
                    if($body -eq $null)
                    {
                        return "Body is required for PATCH method"
                    }
                    else{
                    $output = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $Body
                    }
                }
            
                "Delete"{
                    
                    $output = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
                }
            
                default{
            
                    $output = "Invalid Method $Method"
                }
        }
        
        
        
        }
        
        
        END
        {
        return $output 
        }
        
        
    }


#UserManagement Read 
$tenantName = ""
$clientId = ""
$clientSecret = ""

$token = Get-AccessToken -TenantName $tenantName -ClientId $clientId -ClientSecret $clientSecret

$userid = "00000000-0000-0000-0000-000000000000"


$URIALLUSERS = "https://graph.microsoft.com/v1.0/users"
$AllUsers = Send-GraphRequest -AccessToken $token -Method Get -Uri $URIALLUSERS -ContentType "application/json"


$URISINGLEUSER = "https://graph.microsoft.com/v1.0/users/$($userid)"
$SingleUser = Send-GraphRequest -AccessToken $token -Method Get -Uri $URISINGLEUSER -ContentType "application/json"


$URIRoleassigments = "https://graph.microsoft.com/v1.0/users/$($userid)/appRoleAssignments?$count=true"
$Roleassigments = Send-GraphRequest -AccessToken $token -Method Get -Uri $URIRoleassigments -ContentType "application/json"


$Roleassigmentid = "00000000-0000-0000-0000-000000000000"
$URIsingleapproleassigment = "https://graph.microsoft.com/v1.0/users/$($userid)/appRoleAssignments/$($Roleassigmentid)"
$singleapproleassigment = Send-GraphRequest -AccessToken $token -Method Get -Uri $URIsingleapproleassigment -ContentType "application/json"


#GET /users/{id}/licenseDetails
$URILicensedetails = "https://graph.microsoft.com/v1.0/users/$($userid)/licenseDetails"
$Licensedetails = Send-GraphRequest -AccessToken $token -Method Get -Uri $URILicensedetails -ContentType "application/json"


#GET /users/{id | userPrincipalName}/memberOf
$URImemberOf = "https://graph.microsoft.com/v1.0/users/$($userid)/memberOf"
$memberOf = Send-GraphRequest -AccessToken $token -Method Get -Uri $URImemberOf -ContentType "application/json"

#Patch /users/{id | userPrincipalName}
$URIPatch = "https://graph.microsoft.com/v1.0/users/$($userid)"
$Patch = Send-GraphRequest -AccessToken $writetoken -Method Patch -Uri $URIPatch -ContentType "application/json" -Body '{"givenName": "Adele", "surname": "Vance", "displayName": "Adele Vance"}'

#POST /users/{id | userPrincipalName}/assignLicense
$URIassignLicense = "https://graph.microsoft.com/v1.0/users/$($userid)/assignLicense"
$assignLicense = Send-GraphRequest -AccessToken $writetoken -Method Post -Uri $URIassignLicense -ContentType "application/json" -Body '{"addLicenses": [{"disabledPlans": [],"skuId": "c42b9cae-ea4f-4ab7-9717-81576235ccac"}],"removeLicenses": []}'


#DELETE /users/{user-id}/appRoleAssignments/{appRoleAssignment-id}
$URIdeleteappRoleAssignments = "https://graph.microsoft.com/v1.0/users/$($userid)/appRoleAssignments/$($Roleassigmentid)"
$deleteappRoleAssignments = Send-GraphRequest -AccessToken $writetoken -Method Delete -Uri $URIdeleteappRoleAssignments -ContentType "application/json"



<#
Graph API working Filters & selects :


        $filterurl = "https://graph.microsoft.com/v1.0/users?`$filter=onPremisesImmutableId eq 'xxxxxx'"
        $filterurl = "https://graph.microsoft.com/v1.0/users?`$filter=onPremisesImmutableId eq 'xxxxxx'&`$select=id,displayName,givenName,onPremisesImmutableId"
        $filterurl = "https://graph.microsoft.com/v1.0/users?`$filter=id eq '5211125cfe-f7336-4e92-87a1-3618a44343434343434349d882'&`$select=id,displayName,givenName,onPremisesImmutableId"
        $filterurl = "https://graph.microsoft.com/v1.0/users?`$filter=id eq '$objectid'"
        
#>




