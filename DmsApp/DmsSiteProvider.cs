using Microsoft.Extensions.Options;
using SSL_Core_Web_App_Identify_Sample.SharePoint;

namespace SSL_Core_Web_App_Identify_Sample.DmsApp;

/// <summary>Resolves (once) the URL of the SharePoint site that holds the DMS lists.</summary>
public sealed class DmsSiteProvider
{
    private readonly ISharePointService _sharePoint;
    private readonly DmsOptions _options;
    private volatile Uri? _siteUri;

    public DmsSiteProvider(ISharePointService sharePoint, IOptions<DmsOptions> options)
    {
        _sharePoint = sharePoint;
        _options = options.Value;
    }

    /// <summary>Always ends with '/', e.g. https://contoso.sharepoint.com/sites/DMS/</summary>
    public async Task<Uri> GetSiteUriAsync(CancellationToken ct = default)
    {
        var cached = _siteUri;
        if (cached is not null) return cached;

        var url = _options.SiteUrl;
        if (string.IsNullOrWhiteSpace(url))
        {
            var site = await _sharePoint.GetSiteAsync(ct);
            url = site.WebUrl;
        }

        url = url?.Trim().TrimEnd('/');

        if (string.IsNullOrEmpty(url)
            || !Uri.TryCreate(url + "/", UriKind.Absolute, out var uri)
            || uri.Scheme != Uri.UriSchemeHttps)
        {
            throw new InvalidOperationException(
                "The SharePoint site URL could not be determined. Set Dms:SiteUrl to the full https URL of the DMS site.");
        }

        _siteUri = uri;
        return uri;
    }
}
