using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Caching.Memory;

namespace SSL_Core_Web_App_Identify_Sample.DmsApp;

/// <summary>
/// SharePoint only sees the APPLICATION (token), so "Created By" / "Modified By" would always be "SharePoint App".
/// After every successful write that goes through the proxy, this class sets Author / Editor of the touched
/// list item or file to the REAL signed-in portal user (ValidateUpdateListItem with bNewDocumentUpdate = true:
/// it does not create an extra version and does not touch the Modified date).
/// Any failure here is only logged - it must never break the user's actual save/upload.
/// </summary>
public sealed class DmsUserStamper
{
    private const string ClaimsPrefix = "i:0#.f|membership|";

    // .../items(123)  or  .../ListItemAllFields   (end of path)
    private static readonly Regex ItemPath =
        new(@"^(?<base>.*/items\(\d+\))(/validateupdatelistitem)?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex FileItemPath =
        new(@"^(?<base>.*/listitemallfields)(/validateupdatelistitem)?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ItemAddPath =
        new(@"^(?<list>.*/lists(/getbytitle\('[^']*'\)|\(guid'[^']*'\)|/getbyid\('[^']*'\))?)/items$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex UploadPath =
        new(@"/files/(add|addusingpath)\b|/finishupload\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly HttpClient _http;
    private readonly DmsProxy _proxy;
    private readonly IMemoryCache _cache;
    private readonly ILogger<DmsUserStamper> _logger;

    public DmsUserStamper(HttpClient http, DmsProxy proxy, IMemoryCache cache, ILogger<DmsUserStamper> logger)
    {
        _http = http;
        _proxy = proxy;
        _cache = cache;
        _logger = logger;
    }

    /// <summary>True when this write may need stamping (so the proxy knows to buffer the response).</summary>
    public static bool NeedsStamp(Uri target, string method, string? overrideMethod) =>
        Classify(target, method, overrideMethod).Kind != WriteKind.None;

    public async Task StampAsync(
        Uri target, string method, string? overrideMethod, byte[] responseBody,
        string? userEmail, string? userName, string token, CancellationToken ct)
    {
        try
        {
            var identity = !string.IsNullOrWhiteSpace(userEmail) ? userEmail
                : (userName != null && userName.Contains('@') ? userName : null);
            if (identity is null) return; // local account without an e-mail: nothing to map to

            var siteUri = await _proxy.GetSiteUriAsync(ct);
            var apiRoot = siteUri.GetLeftPart(UriPartial.Authority);
            var info = Classify(target, method, overrideMethod);

            string? itemBase = null;      // absolute URL of the item, e.g. https://x/sites/a/_api/web/lists/.../items(5)
            var setAuthor = false;

            switch (info.Kind)
            {
                case WriteKind.ItemUpdate:
                    itemBase = apiRoot + info.BasePath;
                    // Safety net for uploads: right after every upload the app writes the metadata onto the new item.
                    // If Author is still the app identity (Author == Editor), fix Created By here as well.
                    setAuthor = await AuthorIsStillAppAsync(itemBase, token, ct);
                    break;

                case WriteKind.ItemAdd:
                {
                    var id = ReadInt(responseBody, "Id") ?? ReadInt(responseBody, "ID");
                    if (id is null) return;
                    itemBase = $"{apiRoot}{info.BasePath}/items({id})";
                    setAuthor = true;
                    break;
                }

                case WriteKind.Upload:
                {
                    var serverRelative = ReadString(responseBody, "ServerRelativeUrl");
                    if (string.IsNullOrEmpty(serverRelative))
                    {
                        _logger.LogInformation("Upload answer has no ServerRelativeUrl (chunk step?) - nothing to stamp yet.");
                        return;
                    }
                    var escaped = Uri.EscapeDataString(serverRelative.Replace("'", "''"));
                    itemBase = new Uri(siteUri,
                        $"_api/web/GetFileByServerRelativePath(decodedurl='{escaped}')/ListItemAllFields").ToString();
                    // Overwriting keeps the original Author in SharePoint. Only a brand-new file (Author == Editor == app) gets one.
                    setAuthor = await AuthorIsStillAppAsync(itemBase, token, ct);
                    break;
                }

                default:
                    return;
            }

            var login = ClaimsPrefix + identity;
            if (!await EnsureUserAsync(siteUri, login, token, ct)) return;

            var fields = new List<object>();
            if (setAuthor) fields.Add(new { FieldName = "Author", FieldValue = UserValue(login) });
            fields.Add(new { FieldName = "Editor", FieldValue = UserValue(login) });

            var body = JsonSerializer.Serialize(new { formValues = fields, bNewDocumentUpdate = true });
            using var req = NewRequest(HttpMethod.Post, itemBase + "/ValidateUpdateListItem", token);
            req.Content = new StringContent(body, Encoding.UTF8, "application/json");
            req.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=nometadata");

            using var resp = await _http.SendAsync(req, ct);
            var text = await resp.Content.ReadAsStringAsync(ct);

            if (resp.IsSuccessStatusCode && !text.Contains("\"HasException\":true", StringComparison.OrdinalIgnoreCase))
                _logger.LogInformation("Created/Modified By set to {Login} on {Item} (author: {Author}).", login, itemBase, setAuthor);
            else
                _logger.LogWarning("Could not set Created/Modified By for {Item}: HTTP {Status} {Body}",
                    itemBase, (int)resp.StatusCode, text.Length > 300 ? text[..300] : text);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            _logger.LogWarning(ex, "Setting Created/Modified By to the signed-in user failed (the save itself succeeded).");
        }
    }

    // ------------------------------------------------------------------ helpers

    private static string UserValue(string login) =>
        JsonSerializer.Serialize(new[] { new { Key = login } });

    private async Task<bool> EnsureUserAsync(Uri siteUri, string login, string token, CancellationToken ct)
    {
        var key = "dms:spuser:" + login.ToLowerInvariant();
        if (_cache.TryGetValue(key, out bool known) && known) return true;

        using var req = NewRequest(HttpMethod.Post, new Uri(siteUri, "_api/web/ensureuser").ToString(), token);
        req.Content = new StringContent(JsonSerializer.Serialize(new { logonName = login }), Encoding.UTF8, "application/json");
        req.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=nometadata");

        using var resp = await _http.SendAsync(req, ct);
        if (!resp.IsSuccessStatusCode)
        {
            _logger.LogWarning("ensureuser failed for {Login}: HTTP {Status}. Created/Modified By stays 'SharePoint App'.",
                login, (int)resp.StatusCode);
            return false;
        }

        _cache.Set(key, true, TimeSpan.FromHours(6));
        return true;
    }

    private async Task<bool> AuthorIsStillAppAsync(string itemBase, string token, CancellationToken ct)
    {
        using var req = NewRequest(HttpMethod.Get, itemBase + "?$select=AuthorId,EditorId", token);
        using var resp = await _http.SendAsync(req, ct);
        if (!resp.IsSuccessStatusCode) return false;

        var bytes = await resp.Content.ReadAsByteArrayAsync(ct);
        var a = ReadInt(bytes, "AuthorId");
        var e = ReadInt(bytes, "EditorId");
        return a is not null && a == e;
    }

    private static HttpRequestMessage NewRequest(HttpMethod method, string url, string token)
    {
        var req = new HttpRequestMessage(method, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Headers.Accept.ParseAdd("application/json;odata=nometadata");
        req.Headers.TryAddWithoutValidation("User-Agent", "NONISV|SSL|DmsPortal/1.0");
        return req;
    }

    private static int? ReadInt(byte[] json, string name)
    {
        var el = Find(json, name);
        if (el is null) return null;
        if (el.Value.ValueKind == JsonValueKind.Number && el.Value.TryGetInt32(out var n)) return n;
        return null;
    }

    private static string? ReadString(byte[] json, string name)
    {
        var el = Find(json, name);
        return el is { ValueKind: JsonValueKind.String } ? el.Value.GetString() : null;
    }

    /// <summary>Finds a property in either nometadata ({"Id":1}) or verbose ({"d":{"Id":1}}) JSON.</summary>
    private static JsonElement? Find(byte[] json, string name)
    {
        if (json.Length == 0) return null;
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object) return null;
            if (root.TryGetProperty("d", out var d) && d.ValueKind == JsonValueKind.Object) root = d;

            foreach (var p in root.EnumerateObject())
                if (p.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    return p.Value.Clone();
        }
        catch (JsonException) { }
        return null;
    }

    private enum WriteKind { None, ItemUpdate, ItemAdd, Upload }

    private static (WriteKind Kind, string BasePath) Classify(Uri target, string method, string? overrideMethod)
    {
        if (HttpMethods.IsGet(method) || HttpMethods.IsHead(method) || HttpMethods.IsOptions(method))
            return (WriteKind.None, "");

        var effective = (string.IsNullOrWhiteSpace(overrideMethod) ? method : overrideMethod).ToUpperInvariant();
        if (effective == "DELETE") return (WriteKind.None, "");

        var path = Uri.UnescapeDataString(target.AbsolutePath).TrimEnd('/');

        // Order matters: ".../files/add(...)" must win over generic matches.
        if (UploadPath.IsMatch(path)) return (WriteKind.Upload, "");

        var m = ItemPath.Match(path);
        if (m.Success) return (WriteKind.ItemUpdate, target.AbsolutePath.Substring(0, FindBaseLength(target.AbsolutePath, m.Groups["base"].Value)));

        m = FileItemPath.Match(path);
        if (m.Success) return (WriteKind.ItemUpdate, target.AbsolutePath.Substring(0, FindBaseLength(target.AbsolutePath, m.Groups["base"].Value)));

        m = ItemAddPath.Match(path);
        if (m.Success && effective == "POST")
            return (WriteKind.ItemAdd, target.AbsolutePath.TrimEnd('/')[..^"/items".Length]);

        return (WriteKind.None, "");
    }

    // The unescaped path and the raw path can differ in length (%20 etc.); cut the RAW path at the same logical end.
    private static int FindBaseLength(string rawPath, string unescapedBase)
    {
        var raw = rawPath.TrimEnd('/');
        for (var i = raw.Length; i > 0; i--)
        {
            if (Uri.UnescapeDataString(raw[..i]).Length == unescapedBase.Length
                && Uri.UnescapeDataString(raw[..i]).Equals(unescapedBase, StringComparison.OrdinalIgnoreCase))
                return i;
        }
        return raw.Length;
    }
}
