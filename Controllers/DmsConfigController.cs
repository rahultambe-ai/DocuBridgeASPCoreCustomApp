using System.Text.Json;
using Azure.Identity;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using SSL_Core_Web_App_Identify_Sample.Auth;
using SSL_Core_Web_App_Identify_Sample.DmsApp;
using SSL_Core_Web_App_Identify_Sample.SharePoint;

namespace SSL_Core_Web_App_Identify_Sample.Controllers;

/// <summary>
/// Replaces the SPFx page context for the React app: site / web ids and the signed-in user.
/// The user comes from the portal login (ICurrentUser), NOT from SharePoint, because SharePoint only sees the app.
/// </summary>
[ApiController]
[Authorize(Policy = AppPolicies.DmsUser, AuthenticationSchemes = AppAuthDefaults.CookieScheme)]
[IgnoreAntiforgeryToken]
[RequireAjaxHeader]
public sealed class DmsConfigController : ControllerBase
{
    private readonly DmsProxy _proxy;
    private readonly ICurrentUser _user;
    private readonly SharePointOptions _sharePointOptions;
    private readonly ILogger<DmsConfigController> _logger;

    public DmsConfigController(
        DmsProxy proxy,
        ICurrentUser user,
        IOptions<SharePointOptions> sharePointOptions,
        ILogger<DmsConfigController> logger)
    {
        _proxy = proxy;
        _user = user;
        _sharePointOptions = sharePointOptions.Value;
        _logger = logger;
    }

    [HttpGet("api/config")]
    public async Task<IActionResult> Config(CancellationToken ct)
    {
        try
        {
            var siteUri = await _proxy.GetSiteUriAsync(ct);

            using var web = await _proxy.GetSharePointJsonAsync("_api/web?$select=Id,Title,ServerRelativeUrl", ct);
            using var site = await _proxy.GetSharePointJsonAsync("_api/site?$select=Id", ct);

            var email = _user.Email ?? string.Empty;
            var displayName = _user.DisplayName ?? _user.UserName ?? string.Empty;
            var loginName = !string.IsNullOrEmpty(email)
                ? "i:0#.f|membership|" + email
                : (_user.UserName ?? string.Empty);

            return Ok(new
            {
                siteUrl = siteUri.ToString().TrimEnd('/'),
                tenantId = _sharePointOptions.TenantId,
                siteId = ReadString(site, "Id"),
                webId = ReadString(web, "Id"),
                webTitle = ReadString(web, "Title"),
                webServerRelativeUrl = ReadString(web, "ServerRelativeUrl"),
                user = new
                {
                    displayName,
                    email,
                    loginName
                }
            });
        }
        catch (DmsUpstreamException ex)
        {
            _logger.LogWarning("SharePoint rejected the configuration call: {Status} {Body}", ex.StatusCode, Shorten(ex.Body));
            return StatusCode(502, new { error = new { code = "sharepoint_rejected", message = Explain(ex) } });
        }
        catch (AuthenticationFailedException ex)
        {
            _logger.LogError(ex, "The app could not obtain an access token for SharePoint.");
            return StatusCode(502, new
            {
                error = new
                {
                    code = "app_token_failed",
                    message = "The application could not sign in to Microsoft Entra ID. Check SharePoint:TenantId / ClientId / ClientSecret or certificate, and that admin consent was granted. Details: " + Shorten(ex.Message)
                }
            });
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            _logger.LogError(ex, "Loading the DMS configuration failed.");
            return StatusCode(502, new
            {
                error = new { code = "config_failed", message = "Could not load the SharePoint site: " + Shorten(ex.Message) }
            });
        }
    }

    private static string ReadString(JsonDocument doc, string name) =>
        doc.RootElement.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;

    private static string Shorten(string? text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var oneLine = text.Replace('\r', ' ').Replace('\n', ' ');
        return oneLine.Length <= 300 ? oneLine : oneLine.Substring(0, 300);
    }

    private static string Explain(DmsUpstreamException ex)
    {
        if (ex.Body.Contains("Unsupported app only token", StringComparison.OrdinalIgnoreCase) || ex.StatusCode == 401)
        {
            return "SharePoint refused the application's access token (HTTP " + ex.StatusCode + "). "
                 + "SharePoint REST accepts app-only tokens only when the app signs in with a CERTIFICATE, not a client secret. "
                 + "Set SharePoint:CertificatePath (+ CertificatePassword) and upload the .cer to the app registration. "
                 + "Also make sure the SharePoint application permission (Sites.Selected / Sites.FullControl.All) has admin consent. See README section 2.";
        }

        if (ex.StatusCode == 403)
        {
            return "The application has no access to this SharePoint site (HTTP 403). Grant the app a role on the site (README section 2).";
        }

        if (ex.StatusCode == 404)
        {
            return "The SharePoint site was not found (HTTP 404). Check Dms:SiteUrl or SharePoint:SiteHostname / SitePath.";
        }

        return "SharePoint answered HTTP " + ex.StatusCode + ". " + Shorten(ex.Body);
    }
}
