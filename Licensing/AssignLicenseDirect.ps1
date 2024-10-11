
    $tenantid = ""
    $clientid = ""
    $secret = ""

    $UPN = ""
    $c_O365License = 'ENTERPRISE MOBILITY + SECURITY E3'


    Write-Output 'INFO: Try to add Office 365 Licenses'
    Write-Output 'INFO: Connect to Graph API and get Token'

    $tokenBody = @{
        Grant_Type    = 'client_credentials'
        Scope         = 'https://graph.microsoft.com/.default'
        Client_Id     = $clientid
        Client_Secret = $secret
    } 

    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody
    $headers = @{
        'Authorization' = "Bearer $($tokenResponse.access_token)"
        'Content-type'  = 'application/json'
    }

    Write-Output 'INFO: Check if User exists in Azure AD before assigning the License'

    $Found = $false
    $l = 1

    do {
        $URLtoGetUser = "https://graph.microsoft.com/v1.0/users?`$filter=startswith(userPrincipalName,'$UPN')"
        $Result = Invoke-RestMethod -Headers $headers -Uri $URLtoGetUser -Method GET

        if ($result.value.Length -eq 1) {
            $Found = $true
        }

        Start-Sleep -Seconds 10
        $l++

        if ($l -gt 18) {
            # Timeout reached, log and throw error
            Write-Output 'ERROR: Failed to get User from AAD after 180 Seconds' 
            throw 'Failed to get User from AAD after 180 Seconds'
        }
    } until ($Found)

    Write-Output 'INFO: Define Function Variables and License List'
    


    #Get SKUIDS
    $URLtoGetSKUs = "https://graph.microsoft.com/v1.0/subscribedSkus"
    $skusResponse = Invoke-RestMethod -Headers $headers -Uri $URLtoGetSKUs -Method GET

    # Create an empty hashtable to store the license names and their corresponding SKU IDs
    $LicenseTable = @{}

    # Loop through the SKUs and add to hashtable
    foreach ($sku in $skusResponse.value) {
        $licenseName = $sku.skuPartNumber
        $skuId = $sku.skuId
        $LicenseTable[$licenseName] = $skuId
    }

# Output of the hashtable looks like
#
#--------------------------------------------------------------------
#Name                           Value
#--------------------------------------------------------------------
#ENTERPRISEPACK                 12345678-9abc-def0-1234-56789abcdef0
#EMS                            abcdef12-3456-789a-bcde-f1234567890a
#FLOW_FREE                      09876543-21ab-cdef-0987-6543210fedcb
#--------------------------------------------------------------------



    $License = @()
    $License += $c_O365License

    Write-Output 'INFO: Iterate through each License and assign'
    foreach ($Lic in $License) {
        Start-Sleep -Seconds 10
        Write-Output "INFO: Working with License: $($Lic)"

        $SkuId = $LicenseTable[$($Lic)]        

        # Request the available SKUs (licenses) in the tenant
        # Assign License
        $URLtoassignLicense = "https://graph.microsoft.com/v1.0/users/$UPN/assignLicense"
        Write-Output "INFO: URL to assign License: $($URLtoassignLicense)"
        
        $BodyJsontoassignLicense = @"
        {
            "addLicenses": [
                {
                    "disabledPlans": [],
                    "skuId": "$($SKUID)"
                }
            ],
            "removeLicenses": []
        }
"@
        Write-Output 'INFO: Running Invoke-RestMethod to assign License'
        $Result = Invoke-RestMethod -Headers $headers -Body $BodyJsontoassignLicense -Uri $URLtoassignLicense -Method POST
    }

