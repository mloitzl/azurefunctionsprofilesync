using System;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using com.loitzl.UserProfileSync;
using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Services.Builder.Configuration;

[assembly: FunctionsStartup(typeof(Startup))]

namespace com.loitzl.UserProfileSync;

public class Startup : FunctionsStartup
{
    public override void Configure(IFunctionsHostBuilder builder)
    {
        var config = builder.GetContext().Configuration;
        var settings = new FunctionSettings();
        config.Bind(settings);
        builder.Services.AddPnPCore(options =>
        {
            // Add the base site url
            options.Sites.Add("Default", new PnPCoreSiteOptions
            {
                SiteUrl = settings.SiteUrl
            });
        });
        builder.Services.AddPnPCoreAuthentication(options =>
        {
            // Load the certificate to use
            var cert = LoadCertificate(settings);
        
            // Configure certificate based auth
            options.Credentials.Configurations.Add("CertAuth",
                new PnPCoreAuthenticationCredentialConfigurationOptions
            {
                ClientId = settings.ClientId,
                TenantId = settings.TenantId,
                X509Certificate = new PnPCoreAuthenticationX509CertificateOptions
                {
                    Certificate = LoadCertificate(settings)
                }
            });
        
            // Connect this auth method to the configured site
            options.Sites.Add("Default", new PnPCoreAuthenticationSiteOptions
            {
                AuthenticationProviderName = "CertAuth"
            });
        });
    }

    // from:
    //    https://github.com/pnp/pnpcore/blob/c79872ace50a0afc860d2e0a7b195d5333eb0b23/samples/Demo.AzureFunction.OutOfProcess.AppOnly/Program.cs#L73
    private static X509Certificate2 LoadCertificate(FunctionSettings settings)
    {
        var certBase64Encoded = Environment.GetEnvironmentVariable("CertificateFromKeyVault");

        if (!string.IsNullOrEmpty(certBase64Encoded))
        {
            // Azure Function flow
            return new X509Certificate2(Convert.FromBase64String(certBase64Encoded),
                "",
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.EphemeralKeySet);
        }

        // Local flow
        var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
        var certificateCollection =
            store.Certificates.Find(X509FindType.FindByThumbprint, settings.CertificateThumbprint, false);
        store.Close();

        return certificateCollection.First();
    }
}