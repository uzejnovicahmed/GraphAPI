




#region create SelfsignedCertificate

$Subject= "Certificatename or Subject"
$CertLocation="C:\hereisthefolderwithmycertificate\inthisfoldercouldbemycert"
$CertName="Certificatename"
$fullcertfilepath = "$CertLocation\$CertName.pfx" 

$thumb = (New-SelfSignedCertificate -Subject $Subject -CertStoreLocation "cert:\LocalMachine\My"  -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter (Get-Date).AddMonths(24)).Thumbprint
 
$pwd = ConvertTo-SecureString -AsPlainText "myversecurepwdishere" -Force
 
Export-PfxCertificate -cert "cert:\localmachine\my\$thumb" -FilePath $fullcertfilepath -Password $pwd
 
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate($fullcertfilepath, $pwd)
$keyValue = [System.Convert]::ToBase64String($cert.GetRawCertData()) | Out-File "$CertLocation\$CertName-base64.crt"


#endregion --------------------------------------------------


$TenantName = ""
$AppId = "" #client id of the application

$Certificate=get-item Cert:\LocalMachine\My\$thumb


# Create base64 hash of certificate
$CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())
 
# Create JWT timestamp for expiration
$StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()
$JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds
$JWTExpiration = [math]::Round($JWTExpirationTimeSpan, 0)
 
# Create JWT validity start timestamp
$NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds
$NotBefore = [math]::Round($NotBeforeExpirationTimeSpan, 0)
 
# Create JWT header
$JWTHeader = @{
    alg = "RS256"
    typ = "JWT"
    # Use the CertificateBase64Hash and replace/strip to match web encoding of base64
    x5t = $CertificateBase64Hash -replace '\+', '-' -replace '/', '_' -replace '='
}
 
# Create JWT payload
$JWTPayLoad = @{
    # What endpoint is allowed to use this JWT
    aud = "https://login.microsoftonline.com/$TenantName/oauth2/token"
 
    # Expiration timestamp
    exp = $JWTExpiration
 
    # Issuer = your application
    iss = $AppId
 
    # JWT ID: random guid
    jti = [guid]::NewGuid()
 
    # Not to be used before
    nbf = $NotBefore
 
    # JWT Subject
    sub = $AppId
}
 
# Convert header and payload to base64
$JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
$EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)
 
$JWTPayLoadToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
$EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)
 
# Join header and Payload with "." to create a valid (unsigned) JWT
$JWT = $EncodedHeader + "." + $EncodedPayload
 
# Get the private key object of your certificate
$PrivateKey = $Certificate.PrivateKey
 
# Define RSA signature and hashing algorithm
$RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
$HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256
 
# Create a signature of the JWT
$Signature = [Convert]::ToBase64String(
    $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT), $HashAlgorithm, $RSAPadding)
) -replace '\+', '-' -replace '/', '_' -replace '='
 
# Join the signature to the JWT with "."
$JWT = $JWT + "." + $Signature
 
# Create a hash with body parameters
$Body = @{
    client_id             = $AppId
    client_assertion      = $JWT
    client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
    scope                 = $Scope
    grant_type            = "client_credentials"
 
}
 
$Url = "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token"
 
# Use the self-generated JWT as Authorization
$Header = @{
    Authorization = "Bearer $JWT"
}
 
# Splat the parameters for Invoke-Restmethod for cleaner code
$PostSplat = @{
    ContentType = 'application/x-www-form-urlencoded'
    Method      = 'POST'
    Body        = $Body
    Uri         = $Url
    Headers     = $Header
}
 
$Request = Invoke-RestMethod @PostSplat
 
$authHeader = @{
    "Authorization"    = "Bearer $($Request.access_token)"
    "Content-type"     = "application/json"
    "consistencyLevel" = "eventual"
}
