using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.AspNetCore.Http.Features;

namespace SSL_Core_Web_App_Identify_Sample.DmsApp;

/// <summary>
/// Calls SharePoint REST and Microsoft Graph with the APPLICATION's identity (client credentials).
/// The browser never holds a SharePoint or Graph token: it only holds this portal's session cookie.
/// </summary>
public sealed class DmsProxy
{
    private const string GraphResource = "https://graph.microsoft.com";
    private const string GraphBase = GraphResource + "/v1.0/";

    // Only first path segments the app actually uses are reachable through the Graph proxy.
    private static readonly HashSet<string> AllowedGraphRoots =
        new(StringComparer.OrdinalIgnoreCase) { "users", "chats", "organization", "groups" };

    private static readonly HashSet<string> ForwardRequestHeaders = new(StringComparer.OrdinalIgnoreCase)
    {
        "Accept", "If-Match", "If-None-Match", "X-HTTP-Method", "OData-Version", "OData-MaxVersion",
        "Prefer", "ConsistencyLevel", "Slug"
    };

    private static readonly HashSet<string> ForwardResponseHeaders = new(StringComparer.OrdinalIgnoreCase)
    {
        "Content-Type", "ETag", "Content-Disposition", "Retry-After"
    };

    private readonly HttpClient _http;
    private readonly IAppTokenProvider _tokens;
    private readonly DmsSiteProvider _site;

    public DmsProxy(HttpClient http, IAppTokenProvider tokens, DmsSiteProvider site)
    {
        _http = http;
        _tokens = tokens;
        _site = site;
    }

    // ---------------------------------------------------------------- site / tokens

    public Task<Uri> GetSiteUriAsync(CancellationToken ct) => _site.GetSiteUriAsync(ct);

    public async Task<string> GetSharePointTokenAsync(CancellationToken ct)
    {
        var siteUri = await _site.GetSiteUriAsync(ct);
        return await _tokens.GetTokenAsync(siteUri.GetLeftPart(UriPartial.Authority), ct);
    }

    public Task<string> GetGraphTokenAsync(CancellationToken ct) => _tokens.GetTokenAsync(GraphResource, ct);

    // ------------------------------------------------------------- URL building

    /// <summary>
    /// Takes the raw (still percent-encoded) text after /api/sharepoint/ and turns it into a SharePoint URL.
    /// Raw text is used on purpose: SharePoint REST puts file paths with spaces, quotes, '#', '%' etc.
    /// inside OData function arguments and decoding + re-encoding them would corrupt the call.
    /// </summary>
    public async Task<Uri> BuildSharePointUriAsync(string rawRelative, CancellationToken ct)
    {
        var siteUri = await _site.GetSiteUriAsync(ct);

        rawRelative = rawRelative.TrimStart('/');
        if (rawRelative.Contains('\\'))
            throw new BadHttpRequestException("Invalid path.");

        Uri uri;
        try { uri = new Uri(siteUri, rawRelative); }
        catch (UriFormatException) { throw new BadHttpRequestException("Invalid path."); }

        var apiPrefix = siteUri.AbsolutePath.TrimEnd('/') + "/_api/";
        if (!uri.Host.Equals(siteUri.Host, StringComparison.OrdinalIgnoreCase)
            || uri.Scheme != Uri.UriSchemeHttps
            || !uri.AbsolutePath.StartsWith(apiPrefix, StringComparison.OrdinalIgnoreCase))
        {
            throw new BadHttpRequestException("Only the configured SharePoint site's _api endpoints can be called.");
        }

        return uri;
    }

    /// <summary>The app's identity is powerful, so the few calls that could remove the whole site are refused.</summary>
    public static bool IsBlockedSharePointCall(Uri target, string method, string? overrideMethod)
    {
        var path = target.AbsolutePath.TrimEnd('/');
        var apiIndex = path.IndexOf("/_api/", StringComparison.OrdinalIgnoreCase);
        if (apiIndex < 0) return true;

        var rel = path.Substring(apiIndex + "/_api/".Length).ToLowerInvariant();
        var effective = (string.IsNullOrWhiteSpace(overrideMethod) ? method : overrideMethod).ToUpperInvariant();

        if ((rel == "web" || rel == "site") && effective == "DELETE") return true;
        return rel == "web/deleteobject" || rel == "site/deleteobject";
    }

    /// <summary>
    /// /api/graph/users?...  ->  https://graph.microsoft.com/v1.0/users?...
    /// "me" is not available with an application identity, so /me/... is mapped to /users/{signed-in user}/...
    /// </summary>
    public static Uri BuildGraphUri(string rawRelative, string? meIdentity)
    {
        rawRelative = rawRelative.TrimStart('/');
        if (rawRelative.Contains('\\'))
            throw new BadHttpRequestException("Invalid path.");

        var q = rawRelative.IndexOf('?');
        var pathPart = q >= 0 ? rawRelative.Substring(0, q) : rawRelative;
        var queryPart = q >= 0 ? rawRelative.Substring(q) : string.Empty;

        var slash = pathPart.IndexOf('/');
        var first = slash >= 0 ? pathPart.Substring(0, slash) : pathPart;
        var rest = slash >= 0 ? pathPart.Substring(slash) : string.Empty;

        if (first.Equals("me", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(meIdentity))
                throw new BadHttpRequestException("Your portal account has no e-mail address, so your Microsoft 365 profile cannot be looked up.");

            rawRelative = "users/" + Uri.EscapeDataString(meIdentity) + rest + queryPart;
        }

        Uri uri;
        try { uri = new Uri(GraphBase + rawRelative); }
        catch (UriFormatException) { throw new BadHttpRequestException("Invalid path."); }

        if (!uri.Host.Equals("graph.microsoft.com", StringComparison.OrdinalIgnoreCase)
            || !AllowedGraphRoots.Contains(GraphRoot(uri)))
        {
            throw new BadHttpRequestException("This Graph endpoint is not exposed by the proxy.");
        }

        return uri;
    }

    public static string GraphRoot(Uri graphUri) =>
        graphUri.AbsolutePath.StartsWith("/v1.0/", StringComparison.Ordinal)
            ? graphUri.AbsolutePath.Substring("/v1.0/".Length).Split('/')[0]
            : string.Empty;

    public static string ExtractRawRelative(HttpContext http, string marker)
    {
        var raw = http.Features.Get<IHttpRequestFeature>()?.RawTarget
                  ?? (http.Request.Path.Value + http.Request.QueryString.Value);
        var i = raw.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (i < 0) throw new BadHttpRequestException("Bad proxy path.");
        return raw.Substring(i + marker.Length);
    }

    // ------------------------------------------------------------------ forwarding

    /// <summary>Streams the incoming request to <paramref name="target"/> and streams the response back.</summary>
    public async Task ForwardAsync(HttpContext http, Uri target, string accessToken, CancellationToken ct)
    {
        using var req = new HttpRequestMessage(new HttpMethod(http.Request.Method), target);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        req.Headers.TryAddWithoutValidation("User-Agent", "NONISV|SSL|DmsPortal/1.0");

        foreach (var header in http.Request.Headers)
        {
            if (ForwardRequestHeaders.Contains(header.Key))
                req.Headers.TryAddWithoutValidation(header.Key, header.Value.ToArray());
        }

        var hasBody = http.Request.ContentLength > 0 || http.Request.Headers.ContainsKey("Transfer-Encoding");
        if (hasBody)
        {
            req.Content = new StreamContent(http.Request.Body);
            if (!string.IsNullOrEmpty(http.Request.ContentType))
                req.Content.Headers.TryAddWithoutValidation("Content-Type", http.Request.ContentType);
            if (http.Request.ContentLength.HasValue)
                req.Content.Headers.ContentLength = http.Request.ContentLength;
        }

        using var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);

        http.Response.StatusCode = (int)resp.StatusCode;
        foreach (var header in resp.Headers.Concat(resp.Content.Headers))
        {
            if (ForwardResponseHeaders.Contains(header.Key))
                http.Response.Headers[header.Key] = header.Value.ToArray();
        }

        // Responses can contain user data: never let a browser or proxy cache them.
        http.Response.Headers["Cache-Control"] = "no-store";

        await resp.Content.CopyToAsync(http.Response.Body, ct);
    }

    /// <summary>GET with the app token; the caller owns (disposes) the response. Used for file streaming.</summary>
    public async Task<HttpResponseMessage> GetSharePointAsync(Uri target, CancellationToken ct)
    {
        var token = await GetSharePointTokenAsync(ct);
        using var req = new HttpRequestMessage(HttpMethod.Get, target);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.TryAddWithoutValidation("User-Agent", "NONISV|SSL|DmsPortal/1.0");
        return await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
    }

    /// <summary>GET against Graph with the app token; the caller owns (disposes) the response.</summary>
    public async Task<HttpResponseMessage> GetGraphAsync(Uri target, CancellationToken ct)
    {
        var token = await GetGraphTokenAsync(ct);
        using var req = new HttpRequestMessage(HttpMethod.Get, target);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        return await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
    }

    /// <summary>Simple JSON GET against SharePoint REST (used by /api/config).</summary>
    public async Task<JsonDocument> GetSharePointJsonAsync(string relative, CancellationToken ct)
    {
        var target = await BuildSharePointUriAsync(relative, ct);
        var token = await GetSharePointTokenAsync(ct);

        using var req = new HttpRequestMessage(HttpMethod.Get, target);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.Accept.ParseAdd("application/json;odata=nometadata");
        req.Headers.TryAddWithoutValidation("User-Agent", "NONISV|SSL|DmsPortal/1.0");

        using var resp = await _http.SendAsync(req, ct);
        if (!resp.IsSuccessStatusCode)
        {
            var body = await resp.Content.ReadAsStringAsync(ct);
            throw new DmsUpstreamException((int)resp.StatusCode, body);
        }

        await using var stream = await resp.Content.ReadAsStreamAsync(ct);
        return await JsonDocument.ParseAsync(stream, cancellationToken: ct);
    }
}
