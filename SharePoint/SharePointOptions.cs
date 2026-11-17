namespace SSL_Core_Web_App_Identify_Sample.SharePoint;

/// <summary>The APPLICATION's identity towards SharePoint (Entra app registration, application permissions).
/// This is independent of how end users sign in to this web app.</summary>
public sealed class SharePointOptions
{
    public const string SectionName = "SharePoint";

    public string TenantId { get; set; } = "";
    public string ClientId { get; set; } = "";

    /// <summary>Use a secret OR a certificate. A certificate is preferred for production.</summary>
    public string? ClientSecret { get; set; }
    public string? CertificatePath { get; set; }
    public string? CertificatePassword { get; set; }

    /// <summary>Full Graph site id ("host,siteCollectionGuid,webGuid"). Most explicit; wins if set.</summary>
    public string? SiteId { get; set; }

    /// <summary>With SitePath: e.g. "mkx8.sharepoint.com" + "/sites/dms". Leave SitePath empty to use the root site.</summary>
    public string? SiteHostname { get; set; }
    public string? SitePath { get; set; }
}
