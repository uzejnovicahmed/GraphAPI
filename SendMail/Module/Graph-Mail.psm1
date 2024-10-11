
function Request-GraphMail {

<#
 .Synopsis
  Get Access Token for the Graph Send Mail Endpoint.
  Set variables in the Script Scope

 .Description
  Get the Access Token fpr the Graph Send Mail Endpoint and put it in a header object.
  Set the most importend Variables into the Script Scope.


 .Parameter TENANTID
  The Tenant ID of the Azure AD Tenant.

 .Parameter CLIENTID
  The Client or Application ID of the Registred App registration.

 .Parameter CLIENTSECRET
  The Secret of the Registred App registration.

 .Parameter MAILSENDER
  The M365 User Mail Address from where the Mail will be sent.

 .Parameter HighlightDate
  The Recipients Mail Address

 .Example
   # Register-GraphMail.

   Initializing-GraphMail `
        -TENANTID $TID `
        -CLIENTID $CID `
        -CLIENTSECRET $CSECRET ` 
        -MAILSENDER "AdeleV@test.onmicrosoft.com" ` 
        -MAILRECIPIENT "dev.mic@onmicrosoft.com"

  return : {

              headers   : {Content-type, Authorization}
              TENANTID  : xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
              CLIENTID  : xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
              SENDER    : AdeleV@test.onmicrosoft.com
              RECIPIENT : dev.mic@onmicrosoft.com
              ExitCode  : 0

  	   }

#>

  [CmdletBinding()]
  param (
    [Parameter(mandatory = $true)]
    [string]
    $TENANTID,
    [Parameter(mandatory = $true)]
    [string]
    $CLIENTID,
    [Parameter(mandatory = $true)]
    [string]
    $CLIENTSECRET,
    [Parameter(mandatory = $true)]
    [System.Net.Mail.MailAddress]
    $MAILSENDER,
    [Parameter(mandatory = $true)]
    [System.Net.Mail.MailAddress]
    $MAILRECIPIENT
  )

  $error.Clear()
  [int]$exitCode = 0
       
  try {

    $tokenBody = @{
      Grant_Type    = "client_credentials"
      Scope         = "https://graph.microsoft.com/.default"
      Client_Id     = $clientId
      Client_Secret = $clientSecret
    }

    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TENANTID/oauth2/v2.0/token" -Method POST -Body $tokenBody -erroraction Stop -TimeoutSec 20
        
  }
  catch {
    $Exception = $_.Exception
    Write-Error $_.Exception
    $exitCode = 1
  }
      
  if ($exitCode -eq 0)
  {
    $headers = @{
      "Authorization" = "Bearer $($tokenResponse.access_token)"
      "Content-type"  = "application/json"
    }

    $script:TENANTID = $TENANTID
    $script:CLIENTID = $CLIENTID
    $script:INIT_SENDER = $MAILSENDER
    $script:INIT_RECIPIENT = $MAILRECIPIENT
    $script:headers = $headers

    

    [PSCustomObject]@{
          
      'headers'   = $headers
      'TENANTID'  = $TENANTID
      'CLIENTID'  = $CLIENTID
      'SENDER'    = $MAILSENDER
      'RECIPIENT' = $MAILRECIPIENT
      'ExitCode'  = $exitCode
    }

    return "Variables TENANTID,CLIENTID,INIT_SENDER,INIT_RECIPIENT,headers are set in Script Scope..."

  }
  elseif ($exitCode -eq 1) {
    [PSCustomObject]@{

      'ExitCode'  = $exitCode
      'Exception' = $Exception
    }
      
      
  }

}



function Send-GraphMail {
<#
 .Synopsis
  Send Mail through the Microsoft Graph API Mail Endpoint

 .Description
  Send Mails with Attachment
  Send Mails without Attachment
  Acces token generated and parsed into Send-GraphMail function with function Request-GraphMail
  

 .Parameter MAILSENDER
  This Parameter is Optional.
  If not used the Sender from Request-GraphMail will be used

 .Parameter MAILRECIPIENT
  This Parameter is Optional.
  If not used the Recipient from Request-GraphMail will be used

 .Parameter Attachmentpath
 Path for the Attachmentfile. Only A Full Path will be accepted. --> "C:\Temp\dummy_attachment.pdf"

 .Parameter HEADERS
  This Parameter is Optional
  This Parameter will be automaticaly set in the Script. $script:HEADERS : Execution of Request-GraphMail is mandatory
  You can also use a custom header with an access token. But it is not necessary

 .Parameter HTML_CONTENT
  The HTML Content for the Mail you will send.

 .Parameter Subject
  A String as an Subject for the Mail.

 .Example
   # Send-GraphMail with Attachment.
   #Before running Send-Graph mail run Request-GraphMail to authenticate. If not Send-GraphMail won't work !

    Send-GraphMail ` 
        -MAILSENDER "AdeleV@dev.onmicrosoft.com" ` 
        -MAILRECIPIENT "dev.mic@onmicrosoft.com" `
        -Attachmentpath "C:\Temp\dummy_attachment.pdf" `
        -Subject "Dummy Subject sent from Graph API ENDPOINT"


  return : {

              status         : Success
              sender         : AdeleV@dev.onmicrosoft.com
              recipient      : dev.mic@onmicrosoft.com
              Attachmentname : dummy_attachment.pdf
              MailBody       : System.Collections.Hashtable
              ExitCode       : 0

  }


.Example 
  #Before running Send-Graph mail run Request-GraphMail to authenticate. If not Send-GraphMail won't work !
  #Send GraphMail without Attachment

###################################################################################

  Send-GraphMail `
        -MAILSENDER "AdeleV@dev.onmicrosoft.com" `
        -MAILRECIPIENT "dev.mic@onmicrosoft.com" `
        -Subject "Dummy Subject sent from Graph API ENDPOINT"


###################################################################################


  result : 

  {

              status         : Success
              sender         : AdeleV@dev.onmicrosoft.com
              recipient      : dev.mic@onmicrosoft.com
              MailBody       : System.Collections.Hashtable
              ExitCode       : 0
              
  }
#>


  [CmdletBinding()]
  param (
    [Parameter(mandatory = $false)]
    [System.Net.Mail.MailAddress]
    $MAILSENDER = $Script:INIT_SENDER,
    [Parameter(mandatory = $false)]
    [System.Net.Mail.MailAddress]
    $MAILRECIPIENT = $Script:INIT_RECIPIENT,
    [Parameter(mandatory = $false)][ValidateScript({
        if (-Not ($_ | Test-Path) ) {
          throw "File or folder does not exist" 
        }
        if (-Not ($_ | Test-Path -PathType Leaf) ) {
          throw "The Path argument must be a file. Folder paths are not allowed."
        }
        return $true
      })]
    [System.IO.FileInfo]
    $Attachmentpath = $null,
    [Parameter(mandatory = $false)]
    [hashtable]
    $HEADERS = $script:HEADERS,
    [Parameter(mandatory = $false)]
    [string]
    $HTML_CONTENT,
    [Parameter(Mandatory = $false)]
    [string]
    $Subject = "Defaultsubject: Please Modify"
  )

  $error.Clear()
  [int]$exitCode = 0

  #region Get-Attachment
  if ($attachmentpath -ne $null) {

    $FileName = (Get-Item -Path $Attachmentpath).name
    $base64string = [Convert]::ToBase64String([IO.File]::ReadAllBytes($Attachmentpath))

    $Body_attachment = @{
      message         = @{
        subject      = "$($Subject)"
        body         = @{
          contentType = "HTML"
          content     = "$($HTML_CONTENT)"
        }
                          
        toRecipients = @(
          @{
            emailAddress = @{
              address = "$($MAILRECIPIENT)"
            }
          }
        )
                          
        attachments  = @(
          @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            Name          = "$FileName"
            ContentType   = "text/plain"
            ContentBytes  = "$base64string"
          }
        )
      }
      saveToSentItems = "false"
    }
  }
  elseif ($attachmentpath -eq $null) {

    $Body = @{
      Message = @{
        Subject      = "$($Subject)"
        Body         = @{
          ContentType = "HTML"
          Content     = "$HTML_CONTENT"
        }
        ToRecipients = @(
          @{
            EmailAddress = @{
              Address = "$MAILRECIPIENT"
            }
          }
        )
		
      }
    }

  }

  Write-Output "Mailsender : $($MAILSENDER)"
  $SEND_URL = "https://graph.microsoft.com/v1.0/users/$($MAILSENDER)/sendMail"



  if ($attachmentpath -ne $null) {

    Write-Output "Send Mail with Attachment"
    try {   
      Invoke-RestMethod -Method POST -Uri $SEND_URL -Headers $script:headers -Body $($Body_attachment | convertto-json -depth 4) -erroraction stop -TimeoutSec 30
    
      [PSCustomObject]@{
          
        'status'         = "Success"
        'sender'         = "$($SEND_URL)"
        'recipient'      = "$($MAILRECIPIENT)"
        'Attachmentname' = "$($FileName)"
        'MailBody'       = "$($Body_attachment)"
        'ExitCode'       = $exitCode
      }

      return "Mail was successfully sent to $($MAILRECIPIENT) from $($MAILSENDER) with attachment $($FileName)"


    }
    catch {
      $Exception = $_.Exception
      $exitCode = 1
      Write-Error $_.Exception
    }
  }
  elseif ($attachmentpath -eq $null) {
    Write-Output "Send Mail without Attachment"
    Write-Output "$HTML_CONTENT"
    try {
      Invoke-RestMethod -Method POST -Uri $SEND_URL -Headers $script:headers -Body $($Body | convertto-json -depth 4) -erroraction stop -TimeoutSec 10

      [PSCustomObject]@{
          
        'status'    = "Success"
        'sender'    = "$($MAILSENDER)"
        'HTTPURL'   = "$($SEND_URL)"
        'recipient' = "$($MAILRECIPIENT)"
        'MailBody'  = "$($Body)"
        'ExitCode'  = $exitCode
      }

      return "Mail was successfully sent to $($MAILRECIPIENT) from, $($MAILSENDER)"

    }
    catch {
      $Exception = $_.Exception
      $exitCode = 1
      Write-Error $_.Exception
    }
  }

  if ($exitCode -eq 1) {
    [PSCustomObject]@{

      'ExitCode'  = $exitCode
      'Exception' = $Exception
    }
  }

}

Export-ModuleMember -Function Request-GraphMail;
Export-ModuleMember -Function Send-GraphMail;
