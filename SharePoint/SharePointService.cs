using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace SSL_Core_Web_App_Identify_Sample.SharePoint;

internal sealed class SharePointService : ISharePointService
{
    private const string SiteCacheKey = "sp:site";
    private const string LibrariesCacheKey = "sp:libraries";

    private readonly GraphServiceClient _graph;
    private readonly SharePointOptions _options;
    private readonly IMemoryCache _cache;

    public SharePointService(
        GraphServiceClient graph,
        IOptions<SharePointOptions> options,
        IMemoryCache cache)
    {
        _graph = graph;
        _options = options.Value;
        _cache = cache;
    }

    public async Task<SiteInfo> GetSiteAsync(CancellationToken ct = default)
    {
        var site = await _cache.GetOrCreateAsync(SiteCacheKey, async entry =>
        {
            entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromHours(1);

            var key =
                !string.IsNullOrWhiteSpace(_options.SiteId) ? _options.SiteId
                : !string.IsNullOrWhiteSpace(_options.SitePath) && !string.IsNullOrWhiteSpace(_options.SiteHostname)
                    ? $"{_options.SiteHostname}:/{_options.SitePath.Trim('/')}"
                    : "root";

            var s = await _graph.Sites[key].GetAsync(cancellationToken: ct);

            return new SiteInfo(
                s?.Id ?? throw new SharePointException("SharePoint site could not be resolved.", 404),
                s.DisplayName,
                s.WebUrl);
        });

        return site!;
    }

    public async Task<IReadOnlyList<SharePointListInfo>> GetListsAsync(CancellationToken ct = default)
    {
        var site = await GetSiteAsync(ct);
        var lists = await _graph.Sites[site.Id].Lists.GetAsync(cancellationToken: ct);

        return lists?.Value?
            .Select(l => new SharePointListInfo(l.Id ?? "", l.Name, l.DisplayName))
            .OrderBy(l => l.DisplayName, StringComparer.CurrentCultureIgnoreCase)
            .ToList()
            ?? new List<SharePointListInfo>();
    }

    public async Task<IReadOnlyList<LibraryInfo>> GetLibrariesAsync(CancellationToken ct = default)
    {
        var libraries = await _cache.GetOrCreateAsync(LibrariesCacheKey, async entry =>
        {
            entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(5);

            var site = await GetSiteAsync(ct);
            var drives = await _graph.Sites[site.Id].Drives.GetAsync(cancellationToken: ct);

            return (IReadOnlyList<LibraryInfo>)(drives?.Value?
                .Where(d => !string.IsNullOrEmpty(d.Id))
                .Select(d => new LibraryInfo(d.Id!, d.Name))
                .OrderBy(d => d.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList()
                ?? new List<LibraryInfo>());
        });

        return libraries ?? Array.Empty<LibraryInfo>();
    }

    public async Task<FolderListing> GetFolderAsync(string driveId, string? folderId, CancellationToken ct = default)
    {
        await EnsureLibraryBelongsToSiteAsync(driveId, ct);

        try
        {
            var isRoot = string.IsNullOrEmpty(folderId);
            var parentId = isRoot ? await GetRootIdAsync(driveId, ct) : folderId!;

            var page = await _graph.Drives[driveId].Items[parentId].Children.GetAsync(cancellationToken: ct);

            var items = page?.Value?
                .Select(i => new FileEntry(
                    i.Id ?? "",
                    i.Name ?? "(unnamed)",
                    i.Folder is not null,
                    i.Size,
                    i.LastModifiedDateTime))
                .OrderByDescending(i => i.IsFolder)
                .ThenBy(i => i.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList()
                ?? new List<FileEntry>();

            return new FolderListing(driveId, parentId, isRoot, items);
        }
        catch (ODataError e) when (e.ResponseStatusCode == 404)
        {
            throw new SharePointException("Folder not found.", 404, e);
        }
    }

    public async Task<FileDownload> DownloadAsync(string driveId, string itemId, CancellationToken ct = default)
    {
        await EnsureLibraryBelongsToSiteAsync(driveId, ct);

        try
        {
            var item = await _graph.Drives[driveId].Items[itemId].GetAsync(cancellationToken: ct);

            if (item is null || item.Folder is not null)
                throw new SharePointException("File not found.", 404);

            var stream = await _graph.Drives[driveId].Items[itemId].Content.GetAsync(cancellationToken: ct)
                         ?? throw new SharePointException("File has no content.", 404);

            return new FileDownload(
                stream,
                item.Name ?? "download",
                item.File?.MimeType ?? "application/octet-stream");
        }
        catch (ODataError e) when (e.ResponseStatusCode == 404)
        {
            throw new SharePointException("File not found.", 404, e);
        }
    }

    public async Task CreateFolderAsync(string driveId, string? parentFolderId, string folderName, CancellationToken ct = default)
    {
        await EnsureLibraryBelongsToSiteAsync(driveId, ct);

        try
        {
            var parentId = string.IsNullOrEmpty(parentFolderId)
                ? await GetRootIdAsync(driveId, ct)
                : parentFolderId;

            var folder = new DriveItem
            {
                Name = folderName,
                Folder = new Folder(),
                AdditionalData = new Dictionary<string, object>
                {
                    ["@microsoft.graph.conflictBehavior"] = "fail"
                }
            };

            await _graph.Drives[driveId].Items[parentId].Children.PostAsync(folder, cancellationToken: ct);
        }
        catch (ODataError e) when (e.ResponseStatusCode == 409)
        {
            throw new SharePointException($"A folder or file named '{folderName}' already exists.", 409, e);
        }
        catch (ODataError e) when (e.ResponseStatusCode == 404)
        {
            throw new SharePointException("Parent folder not found.", 404, e);
        }
    }

    private async Task<string> GetRootIdAsync(string driveId, CancellationToken ct)
    {
        var root = await _graph.Drives[driveId].Root.GetAsync(cancellationToken: ct);
        return root?.Id ?? throw new SharePointException("Library root not found.", 404);
    }

    /// <summary>
    /// SECURITY: the app's Graph permission may reach more than this one site (e.g. Sites.ReadWrite.All).
    /// Never trust a driveId from the browser; only accept libraries that belong to the configured site.
    /// </summary>
    private async Task EnsureLibraryBelongsToSiteAsync(string driveId, CancellationToken ct)
    {
        var libraries = await GetLibrariesAsync(ct);

        if (!libraries.Any(l => string.Equals(l.Id, driveId, StringComparison.Ordinal)))
            throw new SharePointException("Library not found.", 404);
    }
}
