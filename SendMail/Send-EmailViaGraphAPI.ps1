<#
.SYNOPSIS
Sends an email using the Microsoft Graph API with customizable parameters, including multiple attachments.

.DESCRIPTION
This script sends an email using the Microsoft Graph API. It allows you to specify various email parameters such as recipient, sender, CC and BCC recipients, subject, HTML body, importance, and flag status. Additionally, you can provide multiple attachment paths for sending attachments. The script obtains an access token to use the Graph API and handles error conditions, providing detailed status and exception messages.

.PARAMETER Recipient
The email addresses of the recipients or the Recipient. Mandatory.

.PARAMETER MailSender
The email address of the sender.

.PARAMETER CCRecipient
The email addresses of the CC recipients. Optional.

.PARAMETER BCCRecipient
The email addresses of BCC recipients. Optional.

.PARAMETER Subject
The subject of the email. Mandatory.

.PARAMETER HtmlBody
The HTML body of the email. Mandatory.

.PARAMETER AttachmentPath
An array of attachment file paths. Optional.

.PARAMETER Importance
The importance of the email (e.g., "normal", "high"). Default is "normal".

.PARAMETER FlagStatus
The flag status of the email (e.g., "flagged", "complete", "notFlagged"). Default is "notFlagged".

.NOTES
File Name      : Send-EmailViaGraphAPI.ps1
Prerequisite   : You need to have valid Microsoft Graph API Permissions.
Permissions    : Mail.Send
Author         : Uzejnovic Ahmed


.EXAMPLE
Send-EmailViaGraphAPI.ps1 -Recipients @("recipient@example.com","recipient2@example.com") -MailSender "mailsender@example.com" -Subject "Test Email" -HtmlBody "<p>This is a test email.</p>"

This example sends a simple email with the specified recipients, sender, subject, and HTML body.

.EXAMPLE
Send-EmailViaGraphAPI.ps1 -Recipients "recipient@example.com" -MailSender "mailsender@example.com" -Subject "Email with Attachments" -HtmlBody "<p>Check out these attachments!</p>" -AttachmentPath @("C:\Attachment1.pdf", "C:\Attachment2.docx)"

This example sends an email with multiple attachments specified by their file paths, with the specified sender.

.EXAMPLE
$params = @{
    MailSender      = "mailsender@example.com"
    Recipients      = @("recipient@example.com", "recipient2@example.com")
    CCRecipients    = @("ccrecipient@example.com")
    BCCRecipients   = @("bccrecipient@example.com")
    Subject         = "Test Email Sent by Graph API"
    HTMLBody        = $emailBody
    Importance      = "normal"
    FlagStatus      = "notflagged"
    AttachmentPath  = @("Path to the attachment", "Path to the second attachment")
}

# Usage of the script with the defined parameters
Send-EmailViaGraphAPI.ps1 @params

This example sends an email with the specified parameters in a hash table, including the specified sender.

#>


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
    [string]$Subject = "Test Email from Graph Mail",

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



$TENANTID = "<Tenant ID>"
$CLIENTID = "<Application Client ID>"
$CLIENTSECRET = "<Application Secret>"
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
