using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using SSL_Core_Web_App_Identify_Sample.Auth;
using SSL_Core_Web_App_Identify_Sample.DmsApp;

namespace SSL_Core_Web_App_Identify_Sample.Controllers;

/// <summary>
/// Server-side hop to the licence service. The URL stays on the server and there is no CORS to worry about.
/// </summary>
[ApiController]
[Authorize(Policy = AppPolicies.DmsUser, AuthenticationSchemes = AppAuthDefaults.CookieScheme)]
[IgnoreAntiforgeryToken]
[RequireAjaxHeader]
[Route("api/license")]
public sealed class DmsLicenseController : ControllerBase
{
    private readonly IHttpClientFactory _factory;
    private readonly DmsOptions _options;

    public DmsLicenseController(IHttpClientFactory factory, IOptions<DmsOptions> options)
    {
        _factory = factory;
        _options = options.Value;
    }

    [HttpGet]
    public async Task<IActionResult> Get([FromQuery] string tenantId, [FromQuery] string siteId, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(_options.License.ApiUrl))
            return StatusCode(503, new { error = "Dms:License:ApiUrl is not configured." });

        var url = $"{_options.License.ApiUrl}?tenantId={Uri.EscapeDataString(tenantId)}&siteId={Uri.EscapeDataString(siteId)}";
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.TryAddWithoutValidation("ngrok-skip-browser-warning", "true"); // only matters while the licence API sits on a free ngrok URL

        try
        {
            using var resp = await _factory.CreateClient("license").SendAsync(req, ct);
            var body = await resp.Content.ReadAsStringAsync(ct);
            return new ContentResult { StatusCode = (int)resp.StatusCode, Content = body, ContentType = "application/json" };
        }
        catch (HttpRequestException ex)
        {
            return StatusCode(502, new { error = ex.Message });
        }
        catch (TaskCanceledException ex) when (!ct.IsCancellationRequested)
        {
            return StatusCode(504, new { error = "The licence service did not answer in time: " + ex.Message });
        }
    }
}
