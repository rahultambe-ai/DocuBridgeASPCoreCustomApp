namespace SSL_Core_Web_App_Identify_Sample.SharePoint;

public sealed record SiteInfo(string Id, string? DisplayName, string? WebUrl);

public sealed record SharePointListInfo(string Id, string? Name, string? DisplayName);

public sealed record LibraryInfo(string Id, string? Name);

public sealed record FileEntry(string Id, string Name, bool IsFolder, long? Size, DateTimeOffset? LastModified);

public sealed record FolderListing(string DriveId, string FolderId, bool IsRoot, IReadOnlyList<FileEntry> Items);

public sealed record FileDownload(Stream Content, string FileName, string ContentType);

/// <summary>Translated Graph failure the UI can act on (404 = not found, 409 = already exists).</summary>
public sealed class SharePointException : Exception
{
    public int StatusCode { get; }

    public SharePointException(string message, int statusCode, Exception? inner = null)
        : base(message, inner) => StatusCode = statusCode;
}
