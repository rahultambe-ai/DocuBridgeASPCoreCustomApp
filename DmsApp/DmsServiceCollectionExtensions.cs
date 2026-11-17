using System.Net;

namespace SSL_Core_Web_App_Identify_Sample.DmsApp;

public static class DmsServiceCollectionExtensions
{
    /// <summary>Registers everything the hosted DMS app needs. Requires AddSharePointIntegration (TokenCredential, ISharePointService).</summary>
    public static IServiceCollection AddDmsPortal(this IServiceCollection services, IConfiguration configuration)
    {
        services.Configure<DmsOptions>(configuration.GetSection(DmsOptions.SectionName));

        services.AddSingleton<IAppTokenProvider, AppTokenProvider>();
        services.AddSingleton<DmsSiteProvider>();

        services
            .AddHttpClient<DmsProxy>(c => c.Timeout = TimeSpan.FromMinutes(15)) // large uploads
            .ConfigurePrimaryHttpMessageHandler(() => new SocketsHttpHandler
            {
                AllowAutoRedirect = false,
                AutomaticDecompression = DecompressionMethods.All
            });

        services.AddHttpClient("license", c => c.Timeout = TimeSpan.FromSeconds(20));

        return services;
    }
}
