# SharePoint User Profile Sync with Azure Function

Code for the Blog Post:

https://blog.loitzl.com/posts/graph-userprofilesync-with-azure-function

## App Registration in Azure AD

```powershell
Register-PnPAzureADApp `
     -ApplicationName AzureFunction.UserProfileSync `
     -Tenant loitzl.onmicrosoft.com `
     -CertificatePassword (ConvertTo-SecureString -String "YourPassword" -AsPlainText -Force) `
     -CertificatePath certificate.pfx `    
     -GraphApplicationPermissions User.Read.All `
     -SharePointApplicationPermissions User.ReadWrite.All `
     -DeviceLogin
```

##

Create new SharePoint AddIn using `https://<tenant>.sharepoint.com/_layouts/15/appregnew.aspx`:

- Generate Client Id - note the id for later use
- Generate Client Secret - note the secret for later use
- Title: 'Graph SyncJob Delegate App'
- App Domain: "www.localhost.com"
- Redirect URL: "https://www.localhost.com"

Then the permissions need to be assigned using `https://<tenant>-admin.sharepoint.com/_layouts/15/appinv.aspx`

- Enter the App Id from the previous step and click on 'Lookup'.
- In the Add-Ins's permissions request, enter the followings:

```xml
<AppPermissionRequests AllowAppOnlyPolicy="true">
    <AppPermissionRequest Scope="http://sharepoint/social/tenant" Right="FullControl" />
</AppPermissionRequests>
 ```
- Click on "Create"
- Click on "Trust"


# 
```sh
$ npm install -g azure-functions-core-tools@3 --unsafe-perm true
```