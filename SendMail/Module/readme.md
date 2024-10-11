# Send Mail via Graph API

Module to Send Mails via Graph API.

## Getting Started

### Prerequisites

* PS Version 5.1 and above
* APP Registration with Send Mail Permissons

### Installing

Download this repo and import the `Graph-Mail.psm1` module file into your PowerShell session.
Testing with Send-Mail.ps1

```
Import-Module .\Graph-Mail.psm1;

```

## Functions in GRAPH API MAIL Modul

```

Request-GraphMail
Send-GraphMail

```

## Examples

### Request Graph Mail :

```
#$TID = "943a3020-dc82-4n25-aa35-837d3032asdf99ed" # dummy
#$CID = "97e6600c-1b3c-4277-9k44-83e0b78sdaf93b38" # dummy
#$CSECRET = ".GT8Q~tclExvaWdfadghdfahadf1gr~~.YZe-68j7dtq" #dummy


Request-GraphMail ` 
    -TENANTID $TID ` 
    -CLIENTID $CID `
    -CLIENTSECRET $CSECRET `
    -MAILSENDER "AdeleV@dev.onmicrosoft.com" `
    -MAILRECIPIENT "mail.recipient@dev.onmicrosoft.com"

```
### Send Mail without attachment:
```
    Send-GraphMail -Subject "Testsubject" -HTML_CONTENT "This Mail was sent from Graph API"

```
### Send Mail with attachemnt:

```
    Send-GraphMail -Attachmentpath "C:\temp\test.txt" -Subject "Testsubject" -HTML_CONTENT "This Mail was sent from Graph API"

```

## Resources

* https://docs.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-exceptions?view=powershell-7.2
* https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app
* https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http
## Missing features

```
* Errorhandling
* Attachmenttypes
* Check if Request-GraphMail was executed before Send-GraphMail 

```
