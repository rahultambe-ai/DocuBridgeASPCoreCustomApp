using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SSL_Core_Web_App_Identify_Sample.Auth;
using SSL_Core_Web_App_Identify_Sample.DmsApp;

namespace SSL_Core_Web_App_Identify_Sample.Controllers;

/// <summary>GET /api/photo?email=... : a person's Microsoft 365 profile picture for &lt;img src&gt; (Graph, application permission).</summary>
[ApiController]
[Authorize(Policy = AppPolicies.DmsUser, AuthenticationSchemes = AppAuthDefaults.CookieScheme)]
[IgnoreAntiforgeryToken]
[Route("api/photo")]
public sealed class DmsPhotoController : ControllerBase
{
    private readonly DmsProxy _proxy;
    private readonly ILogger<DmsPhotoController> _logger;

    public DmsPhotoController(DmsProxy proxy, ILogger<DmsPhotoController> logger)
    {
        _proxy = proxy;
        _logger = logger;
    }

    [HttpGet]
    public async Task<IActionResult> Get([FromQuery] string email, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(email) || email.Length > 256 || !email.Contains('@')
            || email.IndexOfAny(new[] { '/', '\\', '?', '#', ' ' }) >= 0)
        {
            return NotFound();
        }

        try
        {
            var target = new Uri("https://graph.microsoft.com/v1.0/users/" + Uri.EscapeDataString(email) + "/photo/$value");
            using var resp = await _proxy.GetGraphAsync(target, ct);
            if (!resp.IsSuccessStatusCode)
                return NotFound();

            var bytes = await resp.Content.ReadAsByteArrayAsync(ct);
            var type = resp.Content.Headers.ContentType?.MediaType;
            if (string.IsNullOrEmpty(type) || !type.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                type = "image/jpeg";

            Response.Headers["Cache-Control"] = "private, max-age=3600";
            Response.Headers["X-Content-Type-Options"] = "nosniff";
            return File(bytes, type);
        }
        catch (Exception ex) when (ex is HttpRequestException or Azure.Identity.AuthenticationFailedException)
        {
            _logger.LogWarning(ex, "Could not load the profile picture.");
            return NotFound();
        }
    }
}
