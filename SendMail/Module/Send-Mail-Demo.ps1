Import-Module .\Graph-Mail.psm1

$TID = "913a44420-dc82-4675-aa35-837d303299ed"  #not real id 
$CID = "99e6600c-1a3c-4600-9b44-8fe0b74333b38"  #not real id
$CSECRET = "adfgasdgaerzhhadfgadffgdsagdfhdfhgfdh" #dumysecret

Request-GraphMail -TENANTID $TID -CLIENTID $CID -CLIENTSECRET $CSECRET -MAILSENDER "AdeleV@dev.onmicrosoft.com" -MAILRECIPIENT "dev.test@ms.com"

$data = "test"

Send-GraphMail -HTML_CONTENT $data -Subject "Testsubject"

