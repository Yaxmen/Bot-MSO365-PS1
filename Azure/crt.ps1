$certname = "N3_CertTest"
$cert = New-SelfSignedCertificate -Subject "CN=$certname" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256

 

Export-Certificate -Cert $cert -FilePath "C:\Users\ASDESA\Desktop\$certname.cer"