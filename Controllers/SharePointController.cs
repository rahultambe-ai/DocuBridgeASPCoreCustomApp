using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SSL_Core_Web_App_Identify_Sample.Auth;
using SSL_Core_Web_App_Identify_Sample.SharePoint;

namespace SSL_Core_Web_App_Identify_Sample.Controllers;

/// <summary>
/// SharePoint is accessed with the APPLICATION's identity, so SharePoint itself cannot tell users apart.
/// THIS controller is therefore the authorization boundary: roles decide who may do what,
/// and every action is logged with the real user for audit.
/// It depends on ICurrentUser only; it has no idea which identity provider is in use.
/// </summary>
[Authorize(Policy = AppPolicies.DmsUser)]
public class SharePointController : Controller
{
    private static readonly char[] InvalidFolderNameChars = "\\/:*?\"<>|#%".ToCharArray();

    private readonly ISharePointService _sharePoint;
    private readonly ICurrentUser _currentUser;
    private readonly ILogger<SharePointController> _logger;

    public SharePointController(
        ISharePointService sharePoint,
        ICurrentUser currentUser,
        ILogger<SharePointController> logger)
    {
        _sharePoint = sharePoint;
        _currentUser = currentUser;
        _logger = logger;
    }

    public async Task<IActionResult> Lists(CancellationToken ct)
        => View(await _sharePoint.GetListsAsync(ct));

    public async Task<IActionResult> Libraries(CancellationToken ct)
        => View(await _sharePoint.GetLibrariesAsync(ct));

    public async Task<IActionResult> Files(string driveId, string? folderId, CancellationToken ct)
    {
        if (!IsSafeId(driveId) || (folderId is not null && !IsSafeId(folderId)))
            return BadRequest();

        try
        {
            return View(await _sharePoint.GetFolderAsync(driveId, folderId, ct));
        }
        catch (SharePointException ex) when (ex.StatusCode == 404)
        {
            return NotFound();
        }
    }

    /// <summary>Files are streamed THROUGH the app. Linking users to SharePoint's own URLs would
    /// make SharePoint ask them for a Microsoft sign-in, which is exactly what we are avoiding.</summary>
    public async Task<IActionResult> Download(string driveId, string itemId, CancellationToken ct)
    {
        if (!IsSafeId(driveId) || !IsSafeId(itemId))
            return BadRequest();

        try
        {
            var file = await _sharePoint.DownloadAsync(driveId, itemId, ct);

            _logger.LogInformation("User {UserId} ({UserName}) downloaded '{FileName}' (drive {DriveId}, item {ItemId}).",
                _currentUser.Id, _currentUser.UserName, file.FileName, driveId, itemId);

            Response.Headers.XContentTypeOptions = "nosniff";
            return File(file.Content, file.ContentType, file.FileName);   // always an attachment
        }
        catch (SharePointException ex) when (ex.StatusCode == 404)
        {
            return NotFound();
        }
    }

    [HttpPost]
    [Authorize(Policy = AppPolicies.DmsAdmin)]
    public async Task<IActionResult> CreateFolder(string driveId, string? parentFolderId, string folderName, CancellationToken ct)
    {
        if (!IsSafeId(driveId) || (!string.IsNullOrEmpty(parentFolderId) && !IsSafeId(parentFolderId)))
            return BadRequest();

        var returnTo = RedirectToAction(nameof(Files), new { driveId, folderId = parentFolderId });

        if (!IsValidFolderName(folderName))
        {
            TempData["Error"] = "Folder names cannot be empty, end with a dot or space, or contain \\ / : * ? \" < > | # %";
            return returnTo;
        }

        try
        {
            await _sharePoint.CreateFolderAsync(driveId, parentFolderId, folderName, ct);

            _logger.LogInformation("User {UserId} ({UserName}) created folder '{Folder}' in drive {DriveId}.",
                _currentUser.Id, _currentUser.UserName, folderName, driveId);
        }
        catch (SharePointException ex) when (ex.StatusCode is 404 or 409)
        {
            TempData["Error"] = ex.Message;
        }

        return returnTo;
    }

    /// <summary>First thing to check when SharePoint access fails: prints the resolved site id.</summary>
    [Authorize(Policy = AppPolicies.DmsAdmin)]
    public async Task<IActionResult> Diagnostics(CancellationToken ct)
        => Json(await _sharePoint.GetSiteAsync(ct));

    private static bool IsSafeId(string? id) =>
        !string.IsNullOrWhiteSpace(id) && id.Length <= 256 && id.IndexOfAny(['/', '\\', '?', '#']) < 0;

    private static bool IsValidFolderName(string? name) =>
        !string.IsNullOrWhiteSpace(name)
        && name.Length <= 255
        && name == name.Trim()
        && !name.EndsWith('.')
        && name.IndexOfAny(InvalidFolderNameChars) < 0
        && !name.Any(char.IsControl);
}
