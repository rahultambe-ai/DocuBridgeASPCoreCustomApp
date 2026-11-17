using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;
using Microsoft.Extensions.Options;
using Microsoft.Net.Http.Headers;
using SSL_Core_Web_App_Identify_Sample.Auth;
using SSL_Core_Web_App_Identify_Sample.DmsApp;

namespace SSL_Core_Web_App_Identify_Sample.Controllers;

/// <summary>
/// GET /api/files/open?path=/sites/x/Shared Documents/Folder/file.pdf
/// Streams a SharePoint file THROUGH this app, so links in the DMS never send users to a Microsoft sign-in page.
/// </summary>
[ApiController]
[Authorize(Policy = AppPolicies.DmsUser, AuthenticationSchemes = AppAuthDefaults.CookieScheme)]
[IgnoreAntiforgeryToken]
[Route("api/files/open")]
public sealed class DmsFileController : ControllerBase
{
    private static readonly FileExtensionContentTypeProvider ContentTypes = new();

    // Only these types may be shown inside the browser; everything else is downloaded (prevents script execution on this origin).
    private static readonly HashSet<string> InlineTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "application/pdf", "image/png", "image/jpeg", "image/gif", "image/webp", "image/bmp"
    };

    private readonly DmsProxy _proxy;
    private readonly DmsOptions _options;
    private readonly ICurrentUser _user;
    private readonly ILogger<DmsFileController> _logger;

    public DmsFileController(
        DmsProxy proxy,
        IOptions<DmsOptions> options,
        ICurrentUser user,
        ILogger<DmsFileController> logger)
    {
        _proxy = proxy;
        _options = options.Value;
        _user = user;
        _logger = logger;
    }

    [HttpGet]
    public async Task<IActionResult> Open([FromQuery] string path, CancellationToken ct)
    {
        var siteUri = await _proxy.GetSiteUriAsync(ct);
        var webRelative = siteUri.AbsolutePath.TrimEnd('/');   // "" for the root site, "/sites/dms" otherwise

        if (!IsAllowedPath(path, webRelative))
            return NotFound();

        var fileName = path.Substring(path.LastIndexOf('/') + 1);
        var escaped = Uri.EscapeDataString(path.Replace("'", "''"));
        var target = new Uri(siteUri, $"_api/web/GetFileByServerRelativePath(decodedurl='{escaped}')/$value");

        HttpResponseMessage resp;
        try
        {
            resp = await _proxy.GetSharePointAsync(target, ct);
        }
        catch (Exception ex) when (ex is HttpRequestException or Azure.Identity.AuthenticationFailedException)
        {
            _logger.LogError(ex, "Opening '{Path}' failed.", path);
            return StatusCode(StatusCodes.Status502BadGateway, "SharePoint could not be reached.");
        }

        if (!resp.IsSuccessStatusCode)
        {
            var status = (int)resp.StatusCode;
            resp.Dispose();
            return status == StatusCodes.Status404NotFound
                ? NotFound()
                : StatusCode(StatusCodes.Status502BadGateway, "SharePoint answered HTTP " + status + ".");
        }

        // The response (and the connection behind it) is released when ASP.NET Core finishes sending the file.
        Response.RegisterForDispose(resp);

        if (!ContentTypes.TryGetContentType(fileName, out var contentType))
            contentType = "application/octet-stream";

        var inline = InlineTypes.Contains(contentType);

        var disposition = new ContentDispositionHeaderValue(inline ? "inline" : "attachment");
        disposition.SetHttpFileName(fileName);
        Response.Headers["Content-Disposition"] = disposition.ToString();
        Response.Headers["X-Content-Type-Options"] = "nosniff";
        Response.Headers["Cache-Control"] = "private, no-store";

        _logger.LogInformation("User {UserId} ({UserName}) opened '{Path}'.", _user.Id, _user.UserName, path);

        var stream = await resp.Content.ReadAsStreamAsync(ct);
        return new FileStreamResult(stream, contentType);
    }

    private bool IsAllowedPath(string? path, string webRelative)
    {
        if (string.IsNullOrWhiteSpace(path) || path.Length > 1024 || path[0] != '/')
            return false;

        if (path.Contains('\\') || path.Contains("..") || path.Any(char.IsControl) || path.EndsWith('/'))
            return false;

        // Must be inside the configured site...
        if (webRelative.Length > 0
            && !path.StartsWith(webRelative + "/", StringComparison.OrdinalIgnoreCase))
            return false;

        // ...and inside one of the document libraries the DMS uses.
        if (_options.AllowedLibraries.Count == 0)
            return true;

        foreach (var library in _options.AllowedLibraries)
        {
            if (!string.IsNullOrWhiteSpace(library)
                && path.StartsWith(webRelative + "/" + library.Trim('/') + "/", StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }
}
