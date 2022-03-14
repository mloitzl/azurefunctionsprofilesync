Register-PnPAzureADApp `
    -ApplicationName AzureFunction.UserProfileSync `
    -Tenant loitzl.onmicrosoft.com `
    -CertificatePassword (ConvertTo-SecureString -String "YourPassword" -AsPlainText -Force) `
    -CertificatePath certificate.pfx `
    -GraphApplicationPermissions User.Read.All `
    -SharePointApplicationPermissions User.ReadWrite.All `
    -DeviceLogin