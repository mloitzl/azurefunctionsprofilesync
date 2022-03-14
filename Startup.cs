using System.Security.Cryptography.X509Certificates;
using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PnP.Core.Auth;
using PnP.Core.Services.Builder.Configuration;

[assembly: FunctionsStartup(typeof(ProfileSync.Startup))]

namespace ProfileSync
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            var config = builder.GetContext().Configuration;
            var settings = new FunctionSettings();
            config.Bind(settings);

            builder.Services.AddSingleton(_ => settings);

            builder.Services.AddPnPCore(options =>
            {
            // Disable telemetry because of mixed versions on AppInsights dependencies
            options.DisableTelemetry = true;

            // Configure an authentication provider with certificate (Required for app only)
            var authProvider = new X509CertificateAuthenticationProvider(settings.ClientId,
                    settings.TenantId,
                    StoreName.My,
                    StoreLocation.CurrentUser,
                    settings.CertificateThumbprint);
            // And set it as default
            options.DefaultAuthenticationProvider = authProvider;

            // Add a configuration with the tenant admin-site based on the SiteUrl in app settings
            options.Sites.Add("Default",
                    new PnPCoreSiteOptions
                    {
                        SiteUrl = settings.SiteUrl,
                        AuthenticationProvider = authProvider
                    });
            });
        }
    }
}