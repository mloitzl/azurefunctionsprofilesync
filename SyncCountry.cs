using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using PnP.Core.Services;

namespace com.loitzl.UserProfileSync;

public class SyncCountry
{
    private readonly IPnPContextFactory _contextFactory;
     private readonly ILogger<SyncCountry> _logger;

    public SyncCountry(
         IPnPContextFactory contextFactory
          ,
        ILogger<SyncCountry> logger
    ){
        // _contextFactory = contextFactory;
         _logger = logger;
    }
    [FunctionName("SyncCountry")]
    public async Task Run([TimerTrigger("0 */5 * * * *", RunOnStartup = true)] TimerInfo myTimer)
    {
         _logger.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
        try
        {
            // using var pnpContext = await _contextFactory.CreateAsync("Default");
            // _logger.LogInformation(pnpContext.Uri.ToString());

        }
        catch (Exception ex)
        {
            // _logger.LogError(ex, ex.Message);
        }
    }
}