namespace SSL_Core_Web_App_Identify_Sample.DmsApp;

/// <summary>Settings for the DMS single-page app that is hosted inside this portal (section "Dms").</summary>
public sealed class DmsOptions
{
    public const string SectionName = "Dms";

    /// <summary>Optional. Full https URL of the SharePoint site that holds the DMS lists and libraries,
    /// e.g. https://contoso.sharepoint.com/sites/DMS. When empty the site configured under "SharePoint"
    /// (SiteId, or SiteHostname + SitePath) is used and its URL is read from Microsoft Graph.</summary>
    public string? SiteUrl { get; set; }

    /// <summary>Document libraries a signed-in user may open files from through /api/files/open
    /// (appsettings.json lists the two the DMS uses). Empty list = every library of the site.</summary>
    public List<string> AllowedLibraries { get; set; } = new();

    public LicenseOptions License { get; set; } = new();

    public sealed class LicenseOptions
    {
        public string ApiUrl { get; set; } = string.Empty;
    }
}
