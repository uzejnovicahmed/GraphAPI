# Authentication

$clientID = ""
$Clientsecret = ""
$tenantID = ""


function Remove-DirectlyAssignedLicensesGraph {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )
   
    # Get the access token
    $tokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $clientId
        Client_Secret = $clientSecret
    }

    write-output("first invoke")
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody
    $headers = @{
        "Authorization" = "Bearer $($tokenResponse.access_token)"
        "Content-type"  = "application/json"
    }


    # Get the user's current licenses
    write-output("second invoke")
    write-output("user UPN is: " + $UserPrincipalName)
    #get the licenses
    $Licenses = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users?`$filter=UserPrincipalName eq '$UserPrincipalName'&`$select=licenseAssignmentStates,UserPrincipalName" -Headers $headers -Method GET
    

    # Loop through each license
    foreach ($license in  $Licenses.value.licenseAssignmentStates) {
        #check if license is not directly assigned, assignedByGroup would be a Group GUI if License is assigned by Group, so we check the length of the string
        if ($license.assignedByGroup.length -lt 2) {
            try {
                write-output("licenses removal invoke")
                #build the body
                $BodyJsontoassignLicense = @"
            {
                "addLicenses":[ ],
                "removeLicenses": ["$($license.skuID)"]
            }
"@
                #remove the license
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/assignLicense" -Method POST -Headers $headers -Body $BodyJsontoassignLicense | Out-Null
                Add-Tracelog -TraceLog $TraceLog -Message "Removed $($license.skuID)"
            }
            catch {
                Add-Tracelog -TraceLog $TraceLog -Message "some error occured"
            }
        }
        else {
            Add-Tracelog -TraceLog $TraceLog -Message "Skipped one license since it is inherited or from another removal issue"
        }
    }
}

# Usage:
Remove-DirectlyAssignedLicensesGraph -UserPrincipalName $($UPN)
