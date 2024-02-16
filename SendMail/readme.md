# Send-EmailViaGraphAPI.ps1

## SYNOPSIS
Sends an email using the Microsoft Graph API with customizable parameters, including multiple attachments.

## DESCRIPTION
This script sends an email using the Microsoft Graph API. It allows you to specify various email parameters such as recipient, sender, CC and BCC recipients, subject, HTML body, importance, and flag status. Additionally, you can provide multiple attachment paths for sending attachments. The script obtains an access token to use the Graph API and handles error conditions, providing detailed status and exception messages.

## PARAMETERS
- **Recipient**: The email addresses of the recipients or the Recipient. Mandatory.
- **MailSender**: The email address of the sender.
- **CCRecipient**: The email addresses of the CC recipients. Optional.
- **BCCRecipient**: The email addresses of BCC recipients. Optional.
- **Subject**: The subject of the email. Mandatory.
- **HtmlBody**: The HTML body of the email. Mandatory.
- **AttachmentPath**: An array of attachment file paths. Optional.
- **Importance**: The importance of the email (e.g., "low", "normal", "high"). Default is "normal".
- **FlagStatus**: The flag status of the email (e.g., "flagged", "complete", "notFlagged"). Default is "notFlagged".

## NOTES
- **File Name**: Send-EmailViaGraphAPI.ps1
- **Prerequisite**: You need to have valid Microsoft Graph API Permissions.
- **Permissions**: Mail.Send
- **Author**: Uzejnovic Ahmed

## EXAMPLES

```powershell
# Example 1: Send a simple email
Send-EmailViaGraphAPI.ps1 -Recipients @("recipient@example.com","recipient2@example.com") -MailSender "mailsender@example.com" -Subject "Test Email" -HtmlBody "<p>This is a test email.</p>"

# Example 2: Send an email with attachments
Send-EmailViaGraphAPI.ps1 -Recipients "recipient@example.com" -MailSender "mailsender@example.com" -Subject "Email with Attachments" -HtmlBody "<p>Check out these attachments!</p>" -AttachmentPath @("C:\Attachment1.pdf", "C:\Attachment2.docx")

# Example 3: Sending an email with defined parameters
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
