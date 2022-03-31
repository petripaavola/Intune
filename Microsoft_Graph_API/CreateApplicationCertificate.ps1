# Create self-signed certificate for App Registration
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Devices for IT
# https://www.github.com/petripaavola


# Create certificate for App Registration Microsoft documentation
# Read this first
# https://docs.microsoft.com/en-us/azure/active-directory/develop/howto-create-self-signed-certificate
# https://docs.microsoft.com/en-us/graph/powershell/app-only


$cert = New-SelfSignedCertificate -Subject "CN=NVS_2022_Application" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256    ## Replace {certificateName}

Export-Certificate -Cert $cert -FilePath ".\NVS_2022_Application.cer"   ## Specify your preferred location and replace {certificateName}

$mypwd = ConvertTo-SecureString -String "FIXME" -Force -AsPlainText  ## Replace {myPassword}

Export-PfxCertificate -Cert $cert -FilePath ".\NVS_2022_Application.pfx" -Password $mypwd   ## Specify your preferred location and replace {privateKeyName}




