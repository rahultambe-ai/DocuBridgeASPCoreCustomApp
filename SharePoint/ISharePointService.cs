namespace SSL_Core_Web_App_Identify_Sample.SharePoint;

/// <summary>Everything the app does with SharePoint. Knows nothing about users or sign-in:
/// authorization is the caller's job (see SharePointController).</summary>
public interface ISharePointService
{
    Task<SiteInfo> GetSiteAsync(CancellationToken ct = default);
    Task<IReadOnlyList<SharePointListInfo>> GetListsAsync(CancellationToken ct = default);
    Task<IReadOnlyList<LibraryInfo>> GetLibrariesAsync(CancellationToken ct = default);
    Task<FolderListing> GetFolderAsync(string driveId, string? folderId, CancellationToken ct = default);
    Task<FileDownload> DownloadAsync(string driveId, string itemId, CancellationToken ct = default);
    Task CreateFolderAsync(string driveId, string? parentFolderId, string folderName, CancellationToken ct = default);
}
