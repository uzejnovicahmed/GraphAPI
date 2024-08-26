
Param
(
    [Parameter (Mandatory = $false)]
    [String]$Userprincipalname = "AdeleV@contoso.onmicrosoft.com"
)

$ErrorActionPreference = "stop"


#dont forget to put the app registration details into the Send-GraphMail function ()
function Send-GraphMail {

    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Recipients,

        [Parameter(Mandatory = $true)]
        [string]$MailSender,

        [Parameter(Mandatory = $false)]
        [string[]]$CCRecipients = $null,

        [Parameter(Mandatory = $false)]
        [string[]]$BCCRecipients = $null,

        [Parameter(Mandatory = $true)]
        [string]$Subject = "This is a automated email from Graph API",

        [Parameter(Mandatory = $true)]
        [string]$HtmlBody,

        [Parameter(Mandatory = $false)]
        [string[]]$AttachmentPath = @(),

        [Parameter(Mandatory = $false)]
        [ValidateSet("low", "normal", "high")]
        [string]$Importance = "normal",

        [Parameter(Mandatory = $false)]
        [ValidateSet("notFlagged", "flagged", "complete")]
        [string]$FlagStatus = "notFlagged"
    )



    $TENANTID = "<TenantID>"
    $CLIENTID = "<ClientID>"
    $CLIENTSECRET = "<ClientSecret>"
    $GraphAPIUrl = "https://graph.microsoft.com/v1.0/users/$($MailSender)/sendMail"

    try {
        $tokenBody = @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            Client_Id     = $CLIENTID
            Client_Secret = $CLIENTSECRET
        }

        $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TENANTID/oauth2/v2.0/token" -Method POST -Body $tokenBody -ErrorAction Stop -TimeoutSec 20
    }
    catch {
        Write-Error "Failed to obtain access token. Error details: $_"
        return
    }

    $headers = @{
        "Authorization" = "Bearer $($tokenResponse.access_token)"
        "Content-type"  = "application/json"
    }

    $messageBody = @{
        message = @{
            subject    = $Subject
            body       = @{
                contentType = "HTML"
                content     = $HtmlBody
            }
            flag       = @{ flagStatus = $FlagStatus }
            importance = $Importance
        }
    }

    foreach ($Recipient in $Recipients) {
        $messageBody.message.toRecipients += @(@{ emailAddress = @{ address = $Recipient } })
    }

    if ($CCRecipients) {
        foreach ($CCRecipient in $CCRecipients) {
            $messageBody.message.ccRecipients += @(@{ emailAddress = @{ address = $CCRecipient } })
        }
    }

    if ($BCCRecipients) {
        foreach ($BCCRecipient in $BCCRecipients) {
            $messageBody.message.bccRecipients += @(@{ emailAddress = @{ address = $BCCRecipient } })
        }
    }


    $attachmentList = @()
    if ($AttachmentPath) {
        foreach ($path in $AttachmentPath) {
            $FileName = (Get-Item -Path $path).Name
            $base64string = [Convert]::ToBase64String([IO.File]::ReadAllBytes($path))
            $attachment = @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                Name          = $FileName
                ContentType   = "text/plain"
                ContentBytes  = $base64string
            }
            $attachmentList += $attachment
        }
        $messageBody.message.attachments = $attachmentList
    }



    try {
        Invoke-RestMethod -Method POST -Uri $GraphAPIUrl -Headers $headers -Body ($messageBody | ConvertTo-Json -Depth 4) -ErrorAction Stop -TimeoutSec 30
    

        $output = @{
            "Status"         = "Success"
            "StatusMessage"  = "Mail was successfully sent to $($Recipients) from $($MailSender)"
            "Recipient"      = $Recipients
            "MailSender"     = $MailSender
            "CCRecipient"    = $CCRecipients
            "BCCRecipient"   = $BCCRecipients
            "Subject"        = $Subject
            "Importance"     = $Importance
            "FlagStatus"     = $FlagStatus
            "AttachmentPath" = $AttachmentPath
        }

        Write-Output $output.StatusMessage

        return $output

    }
    catch {
    

        $output = @{
            "Status"           = "Error"
            "ScriptLineNumber" = $_.InvocationInfo.ScriptLineNumber
            "ExceptionMessage" = $_.Exception.Message
            "Exception"        = $_
            "Recipient"        = $Recipients
            "MailSender"       = $MailSender
            "CCRecipient"      = $CCRecipients
            "BCCRecipient"     = $BCCRecipients
            "Subject"          = $Subject
            "Importance"       = $Importance
            "FlagStatus"       = $FlagStatus
            "AttachmentPath"   = $AttachmentPath
        }

        Write-Error "Failed to send the email. Error at Line: $($output.ScriptLineNumber) | Error Details : $_"

        return $output
    }


}


function Get-AccessToken {

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



#endregion

#//----------------------------------------------------------------------------
#//  Main routines
#//----------------------------------------------------------------------------

Try {

    $MailSender = "sender@contoso.onmicrosoft.com"

    $tenantName = "<TenantName>"
    $clientId = "<ClientID>"
    $clientSecret = "<ClientSecret>"
    $token = Get-AccessToken -TenantName $tenantName -ClientId $clientId -ClientSecret $clientSecret
    



    $headers = @{
        "Authorization" = "Bearer $($token)"
        "Content-type"  = "application/json"
        "Accept"        = "application/json"
    }

    Write-Output "Verifying user against Azure AD with Graph API";
    
    $filterurl = "https://graph.microsoft.com/v1.0/users?`$filter=userprincipalname eq '$($Userprincipalname)'&`$select=onPremisesDistinguishedName,userprincipalname,id,Mail,givenname,surname,displayname"
    $AzureADUser = (Invoke-RestMethod -Method Get -Uri $filterurl -Headers $headers).Value
    



    if ($AzureADUser.count -gt 1) {
        Throw "Found more than 1 User"
    }


    #Get Signin Preferences
    $signinpreferencesurl = "https://graph.microsoft.com/beta/users/$($AzureADUser.userPrincipalName)/authentication/signInPreferences"
    $signinpreferences = Invoke-RestMethod -Method Get -Uri $signinpreferencesurl -Headers $headers
    $userPreferredMethodForSecondaryAuthentication = $null
    $userPreferredMethodForSecondaryAuthentication = $signinpreferences.userPreferredMethodForSecondaryAuthentication
    
    #userprefferedtypes = 'push','oath','voiceMobile','voiceAlternateMobile','voiceOffice','sms'

    #Get MFA Methods
    $URIMFAMETHODS = "https://graph.microsoft.com/v1.0/users/$($AzureADUser.userPrincipalName)/authentication/methods"
    $Method2Delete = $null
    $Method2Delete = (Invoke-RestMethod -Method Get -Uri $URIMFAMETHODS -Headers $headers).Value
    $Method2Delete | Add-member -MemberType NoteProperty -Name "defaultsigninmethod" -Value $false
            
    if (!([string]::IsNullOrEmpty($userPreferredMethodForSecondaryAuthentication))) {
        #If user has a preferred method for secondary authentication

        Write-Output "User Preferred Method For Secondary Authentication is : $($userPreferredMethodForSecondaryAuthentication)"

        foreach ($method in $Method2Delete) {

            switch ($userPreferredMethodForSecondaryAuthentication) {
                "push" { 
                    if ($method.'@odata.type'.replace("#", "") -eq "microsoft.graph.microsoftAuthenticatorAuthenticationMethod") {
                        $method.defaultsigninmethod = $true
                        Write-Output "Default sign in method is Push"
                    }
                }
                "oath" { 
                    if ($method.'@odata.type'.replace("#", "") -eq "microsoft.graph.softwareOathAuthenticationMethod") {
                        $method.defaultsigninmethod = $true
                        Write-Output "Default sign in method is Oath"
                    }
                }
                "voiceMobile" { 
                    if ($method.'@odata.type'.replace("#", "") -eq "microsoft.graph.phoneAuthenticationMethod") {
                        $method.defaultsigninmethod = $true
                        Write-Output "Default sign in method is VoiceMobile"
                    }  
                }
                "voiceAlternateMobile" { 
                    if ($method.'@odata.type'.replace("#", "") -eq "microsoft.graph.phoneAuthenticationMethod") {
                        $method.defaultsigninmethod = $true
                        Write-Output "Default sign in method is VoiceAlternateMobile"
                    }
                }
                "voiceOffice" { 
                    if ($method.'@odata.type'.replace("#", "") -eq "microsoft.graph.phoneAuthenticationMethod") {
                        $method.defaultsigninmethod = $true
                        Write-Output "Default sign in method is VoiceOffice"
                    }
                }
                "sms" {
                    if ($method.'@odata.type'.replace("#", "") -eq "microsoft.graph.phoneAuthenticationMethod") {
                        $method.defaultsigninmethod = $true
                        Write-Output "Default sign in method is SMS"
                    }
                            
                }
            }
        }   
                
    }
    else {
        Write-Output "User has no preferred method for secondary authentication"
    }


    $Method2Delete = $Method2Delete | Sort-Object -Property defaultsigninmethod


    $defaulturl = $null
    foreach ($method in $Method2Delete) {

        Write-Output "Try to delete Method $($method)";

        $methodType = $null
        $methodTypes = $method.'@odata.type'.replace("#", "")

        switch ($methodTypes) {
            "microsoft.graph.phoneAuthenticationMethod" { 
                #"Phone"
                $methodType = "phoneMethods"
            }
            "microsoft.graph.microsoftAuthenticatorAuthenticationMethod" { 
                #"AUTH"
                $methodtype = "microsoftAuthenticatorMethods"
            }
            "microsoft.graph.temporaryAccessPassAuthenticationMethod" { 
                #"AUTH"
                $methodtype = "temporaryAccessPassMethods"
            }
            "microsoft.graph.windowsHelloForBusinessAuthenticationMethod" { 
                #"AUTH"
                $methodtype = "windowsHelloForBusinessMethods"
            }
            "microsoft.graph.emailAuthenticationMethod" { 
                #"AUTH"
                $methodtype = "emailMethods"
            }
            "microsoft.graph.passwordAuthenticationMethod" { 
                #"AUTH"
                $methodtype = "passwordMethods"
            }
            "microsoft.graph.softwareOathAuthenticationMethod" { 
                #"AUTH"
                $methodtype = "softwareOathMethods"
            }
            "microsoft.graph.fido2AuthenticationMethod" { 
                #"AUTH"
                $methodtype = "fido2Methods"
            }
        }


        if ([string]$methodType -ne "passwordMethods") {

            $deleteurl = "https://graph.microsoft.com/v1.0/users/$($AzureADUser.userprincipalname)/authentication/$($methodType)/$($method.id)"
    
            try {
                Write-Output "URI for deletion is : $($deleteurl)"; 
                Invoke-RestMethod -Method Delete -Uri $deleteurl -Headers $headers
                Write-Output "Mehtod was deleted : $($methodType)";
            }
            catch {
             
                [array]$errormessage += $_
                Write-Error $errormessage
            }

        }

    }


    [array]$AuthMethods = (Invoke-RestMethod -Method Get -Uri $URIMFAMETHODS -Headers $headers).Value

    
    if ($AuthMethods.count -eq 1) {
    
        Write-Output "All MFA methods have been deleted";
        Write-Output "MFA settings has been removed for account $($AzureADUser.UserPrincipalName)";

        


        $HTMLSUCCESS = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MFA Reset Status Notification</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            color: #333333;
            line-height: 1.6;
        }
        .container {
            max-width: 600px;
            margin: auto;
            padding: 20px;
            border: 1px solid #cccccc;
            border-radius: 10px;
            background-color: #f9f9f9;
        }
        h2 {
            color: #4CAF50;
            text-align: center;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th, td {
            padding: 10px;
            border: 1px solid #dddddd;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
            color: #333333;
        }
        td {
            background-color: #ffffff;
        }
        .footer {
            margin-top: 20px;
            text-align: center;
            font-size: 0.9em;
            color: #777777;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>MFA Reset Successful</h2>

        <p>Dear $($azureaduser.displayname),</p>

        <p>We are pleased to inform you that the Multi-Factor Authentication (MFA) reset for <strong>$($Userprincipalname)</strong> has been successfully completed.</p>

        <table>
            <tr>
                <th>Details</th>
                <th>Information</th>
            </tr>
            <tr>
                <td><strong>User</strong></td>
                <td>$($Azureaduser.Displayname)</td>
            </tr>
            <tr>
                <td><strong>User Email</strong></td>
                <td>$($Azureaduser.Mail)</td>
            </tr>
            <tr>
                <td><strong>Date/Time of Reset</strong></td>
                <td>$(Get-Date)</td>
            </tr>
        </table>

        <p>No further action is required. If you have any questions or need additional assistance, please do not hesitate to contact us.</p>

        <p>Thank you for your attention to this matter.</p>

        <p>Best regards,<br>
        Your Automation Team<br>
       

        <div class="footer">
            <p>This is an automated message. Please do not reply to this email.</p>
        </div>
    </div>
</body>
</html>


"@


        if ( ([string]::IsNullOrEmpty($AzureADUser.Mail)) -eq $true ) {
            Write-Output "Sending notification to $($EMAILTO)";
        }
        else {
            Write-Output "Sending notification to enduser $($EMAILTO)";
        }


        if ($AzureADUser.Mail) {
            Send-GraphMail -Recipients "$($AzureADUser.Mail)" -MailSender "$($MailSender)" -Subject "MFA Reset Successful" -HtmlBody $HTMLSUCCESS
        }

    

    }
    else {

        Write-Output "Following Methods are not deleted : $($AuthMethods.'@odata.type')";
        
        $HTMLFAILURE = @"

        <!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MFA Reset Status Notification</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            color: #333333;
            line-height: 1.6;
        }
        .container {
            max-width: 600px;
            margin: auto;
            padding: 20px;
            border: 1px solid #cccccc;
            border-radius: 10px;
            background-color: #f9f9f9;
        }
        h2 {
            color: #F44336;
            text-align: center;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th, td {
            padding: 10px;
            border: 1px solid #dddddd;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
            color: #333333;
        }
        td {
            background-color: #ffffff;
        }
        .footer {
            margin-top: 20px;
            text-align: center;
            font-size: 0.9em;
            color: #777777;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>MFA Reset Failed</h2>

        <p>Dear $($Azureaduser.Displayname),</p>

        <p>Unfortunately, the Multi-Factor Authentication (MFA) reset attempt for <strong>$($Userprincipalname)</strong> was not successful. Below are the details of the encountered error:</p>

        <table>
            <tr>
                <th>Details</th>
                <th>Information</th>
            </tr>
            <tr>
                <td><strong>User</strong></td>
                <td>$($Azureaduser.Displayname)</td>
            </tr>
            <tr>
                <td><strong>User Email</strong></td>
                <td>$($Azureaduser.Mail)</td>
            </tr>
            <tr>
                <td><strong>Date/Time of Attempt</strong></td>
                <td>$(Get-Date)</td>
            </tr>
            <tr>
                <td><strong>Error Message</strong></td>
                <td>Following Methods are not deleted : $($AuthMethods.'@odata.type')</td>
            </tr>
            <tr>
                <td><strong>Error Details</strong></td>
                <td>$($errormessage)</td>
            </tr>
        </table>

        <p>Please review the error details and take the necessary actions to resolve the issue. If you require assistance or have any questions, feel free to contact us.</p>

        <p>We apologize for any inconvenience this may cause and appreciate your prompt attention to this matter.</p>

        <p>Best regards,<br>
        [Your Automation Team]<br>
        <div class="footer">
            <p>This is an automated message. Please do not reply to this email.</p>
        </div>
    </div>
</body>
</html>


"@


        if ($AzureADUser.Mail) {
            Send-GraphMail -Recipients "$($AzureADUser.Mail)" -MailSender "$($MailSender)" -Subject "MFA Reset Failed" -HtmlBody $HTMLFAILURE
        }
        
    }       
    


}
Catch {
    $ErrorMessage = $_.Exception.Message + "`nAt Line number: $($_.InvocationInfo.ScriptLineNumber)"
    Write-Output ($OutputData | ConvertTo-Json)
    Write-Error -Message $ErrorMessage
}
