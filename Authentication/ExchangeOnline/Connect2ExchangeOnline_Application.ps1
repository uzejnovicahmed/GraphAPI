$tenantname = ''
$appID =  ''
$CertificateFilePath = ''
$pfxPassword = ''

Connect-ExchangeOnline -CertificateFilePath $CertificateFilePath `
-CertificatePassword (ConvertTo-SecureString -String $pfxPassword -AsPlainText -Force) `
-AppID $appID `
-Organization $tenantname


$ExchangeConnectiontmpmodulefilepath = (Get-Connectioninformation).ModuleName

Start-Sleep 5
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

if((Test-Path $ExchangeConnectiontmpmodulefilepath) -eq $true)
{
try{
Remove-Item -Path $ExchangeConnectiontmpmodulefilepath -Recurse -Force -ErrorAction Stop
}
catch{}
}
