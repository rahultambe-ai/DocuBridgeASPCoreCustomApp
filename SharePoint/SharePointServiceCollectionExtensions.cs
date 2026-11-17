using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using System.Security.Cryptography.X509Certificates;

namespace SSL_Core_Web_App_Identify_Sample.SharePoint;

public static class SharePointServiceCollectionExtensions
{
    public static IServiceCollection AddSharePointIntegration(
        this IServiceCollection services, IConfiguration configuration)
    {
        services.AddOptions<SharePointOptions>()
            .Bind(configuration.GetSection(SharePointOptions.SectionName))
            .Validate(o => !string.IsNullOrWhiteSpace(o.TenantId) && !string.IsNullOrWhiteSpace(o.ClientId),
                "SharePoint:TenantId and SharePoint:ClientId are required.")
            .Validate(o => !string.IsNullOrWhiteSpace(o.ClientSecret) || !string.IsNullOrWhiteSpace(o.CertificatePath),
                "Set SharePoint:ClientSecret (user-secrets / Key Vault) or SharePoint:CertificatePath.")
            .ValidateOnStart();

        // ONE credential for the whole app, shared by the Graph client and by the DMS proxy (SharePoint REST).
        // Creating a credential per request (as the earlier sample did) throws away the token cache and
        // re-authenticates against Entra on every call.
        services.AddSingleton<TokenCredential>(sp =>
        {
            var o = sp.GetRequiredService<IOptions<SharePointOptions>>().Value;

            if (!string.IsNullOrWhiteSpace(o.CertificatePath))
            {
                return new ClientCertificateCredential(
                    o.TenantId, o.ClientId,
                    X509CertificateLoader.LoadPkcs12FromFile(o.CertificatePath, o.CertificatePassword));
            }

            return new ClientSecretCredential(o.TenantId, o.ClientId, o.ClientSecret);
        });

        services.AddSingleton(sp => new GraphServiceClient(sp.GetRequiredService<TokenCredential>()));

        services.AddMemoryCache();
        services.AddSingleton<ISharePointService, SharePointService>();

        return services;
    }
}
