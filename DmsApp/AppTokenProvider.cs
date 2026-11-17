using Azure.Core;

namespace SSL_Core_Web_App_Identify_Sample.DmsApp;

/// <summary>Gets access tokens for the APPLICATION's own identity (no user involved).</summary>
public interface IAppTokenProvider
{
    /// <summary>resource = "https://contoso.sharepoint.com" or "https://graph.microsoft.com".</summary>
    Task<string> GetTokenAsync(string resource, CancellationToken ct = default);
}

internal sealed class AppTokenProvider : IAppTokenProvider
{
    private readonly TokenCredential _credential;

    public AppTokenProvider(TokenCredential credential) => _credential = credential;

    public async Task<string> GetTokenAsync(string resource, CancellationToken ct = default)
    {
        var scope = resource.TrimEnd('/') + "/.default";

        // Azure.Identity keeps its own token cache, so calling this on every request is cheap.
        AccessToken token = await _credential.GetTokenAsync(new TokenRequestContext(new[] { scope }), ct);
        return token.Token;
    }
}
