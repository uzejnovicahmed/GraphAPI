
$tenantname = "" 
$clientid = ""
$secret = ""

function Request-AccessToken {

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
    
    try { $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $tokenBody -ErrorAction Stop }catch { return "Error generating access token $($_)" }
    Write-Debug "Successfully generated authentication token"
    return $($tokenResponse.access_token)
}


$token = Request-AccessToken -TenantName $tenantname -ClientId $clientid -ClientSecret $secret


try{

$headers = @{
    "Authorization" = "Bearer $($token)"
    "Content-type"  = "application/json;charset=utf-8"
}
