using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using Microsoft.SharePoint.Client.Utilities;
using PnP.Core.Services;
using PnP.Framework;
using User = Microsoft.Graph.User;

namespace ProfileSync
{
    public class SyncCountry
    {
        private readonly IPnPContextFactory _contextFactory;
        private readonly ILogger<SyncCountry> _logger;
        private readonly FunctionSettings _settings;

        public SyncCountry(
            IPnPContextFactory contextFactory,
            FunctionSettings settings,
            ILogger<SyncCountry> logger)
        {
            _contextFactory = contextFactory;
            _settings = settings;
            _logger = logger;
        }

        [FunctionName("SyncCountry")]
        public async Task Run([TimerTrigger("0 */5 * * * *", RunOnStartup = true)] TimerInfo timer)
        {
            _logger.LogDebug("Running timer triggered 'SyncCountry' job...");

            using var context = await _contextFactory.CreateAsync("Default");

            var graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(requestMessage =>
                context
                    .AuthenticationProvider
                    .AuthenticateRequestAsync(
                        new Uri("https://graph.microsoft.com"),
                        requestMessage)));

            var usersCollection =
                await graphClient
                    .Users
                    .Request()
                    .Select("id,mail,country")
                    .GetAsync();

            do
            {
                await Task.WhenAll(
                    usersCollection
                        .Select(async user =>
                            await SyncProperties(
                                user)));

                if (usersCollection.NextPageRequest != null)
                    usersCollection = await usersCollection.NextPageRequest.GetAsync();
            } while (usersCollection.NextPageRequest != null);
        }

        private async Task<string> SyncProperties(User user)
        {
            using var scope = _logger.BeginScope("{SyncPropertiesId}", Guid.NewGuid());

            _logger.LogDebug("Syncing user '{Mail}'", user.Mail);

            try
            {
                using var clientContext = new AuthenticationManager()
                    .GetACSAppOnlyContext(
                        _settings.SiteUrl,
                        _settings.DelegateAppId,
                        _settings.DelegateAppSecret);

                var resolvedPrincipal =
                    Utility.ResolvePrincipal(
                        clientContext,
                        clientContext.Web,
                        user.Mail,
                        PrincipalType.User,
                        PrincipalSource.All,
                        null,
                        true);

                await clientContext.ExecuteQueryRetryAsync();
                var person = resolvedPrincipal.Value;

                var personLoginName = person.LoginName;

                _logger.LogTrace("✔ Successfully resolved '{LoginName}'", personLoginName);
                _logger.LogTrace("Setting UserProfile CustomCountry to '{Country}'", user.Country);

                var peopleManager = new PeopleManager(clientContext);
                peopleManager.SetSingleValueProfileProperty(
                    personLoginName,
                    "CustomCountry",
                    user.Country);

                clientContext.Load(peopleManager);
                await clientContext.ExecuteQueryRetryAsync();

                _logger.LogInformation("✔ Successfully set UserProfile '{CustomCountry}' for '{LoginName}'", user.Country, personLoginName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Setting UserProfile CustomCountry failed for '{LoginName}', {Message}", user.Mail,
                    ex.Message);
            }

            return user.Mail;
        }
    }
}