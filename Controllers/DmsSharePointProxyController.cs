using Azure.Identity;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SSL_Core_Web_App_Identify_Sample.Auth;
using SSL_Core_Web_App_Identify_Sample.DmsApp;

namespace SSL_Core_Web_App_Identify_Sample.Controllers;

/// <summary>
/// /api/sharepoint/_api/...  ->  {SiteUrl}/_api/...  with the application's identity.
/// The React client's PnPjs and spHttpClient shims send every SharePoint call here.
/// SharePoint cannot tell users apart any more, so THIS endpoint is the authorization boundary:
/// only signed-in users with a DMS role get through, and every write is logged with the real user.
/// </summary>
[ApiController]
[Authorize(Policy = AppPolicies.DmsUser, AuthenticationSchemes = AppAuthDefaults.CookieScheme)]
[IgnoreAntiforgeryToken]
[RequireAjaxHeader]
[Route("api/sharepoint/{**path}")]
public sealed class DmsSharePointProxyController : ControllerBase
{
    private readonly DmsProxy _proxy;
    private readonly ICurrentUser _user;
    private readonly ILogger<DmsSharePointProxyController> _logger;

    public DmsSharePointProxyController(DmsProxy proxy, ICurrentUser user, ILogger<DmsSharePointProxyController> logger)
    {
        _proxy = proxy;
        _user = user;
        _logger = logger;
    }

    [AcceptVerbs("GET", "POST", "PUT", "PATCH", "DELETE")]
    [DisableRequestSizeLimit]
    public async Task Forward(string? path, CancellationToken ct)
    {
        try
        {
            var raw = DmsProxy.ExtractRawRelative(HttpContext, "/api/sharepoint/");
            var target = await _proxy.BuildSharePointUriAsync(raw, ct);

            var overrideMethod = Request.Headers["X-HTTP-Method"].ToString();
            if (DmsProxy.IsBlockedSharePointCall(target, Request.Method, overrideMethod))
            {
                await WriteError(StatusCodes.Status403Forbidden, "blocked", "This SharePoint call is not allowed.");
                return;
            }

            if (!HttpMethods.IsGet(Request.Method))
            {
                _logger.LogInformation("DMS write by {UserId} ({UserName}): {Method} {Path}",
                    _user.Id, _user.UserName, Request.Method, target.AbsolutePath);
            }

            var token = await _proxy.GetSharePointTokenAsync(ct);
            await _proxy.ForwardAsync(HttpContext, target, token, ct);
        }
        catch (BadHttpRequestException ex)
        {
            await WriteError(StatusCodes.Status400BadRequest, "bad_request", ex.Message);
        }
        catch (AuthenticationFailedException ex)
        {
            _logger.LogError(ex, "The app could not obtain an access token for SharePoint.");
            await WriteError(StatusCodes.Status502BadGateway, "app_token_failed",
                "The application could not sign in to Microsoft Entra ID (see server log).");
        }
        catch (HttpRequestException ex)
        {
            await WriteError(StatusCodes.Status502BadGateway, "upstream_error", ex.Message);
        }
        catch (InvalidOperationException ex)
        {
            await WriteError(StatusCodes.Status502BadGateway, "site_unknown", ex.Message);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            _logger.LogError(ex, "DMS SharePoint proxy failed.");
            await WriteError(StatusCodes.Status502BadGateway, "proxy_failed", "SharePoint could not be reached (see server log).");
        }
    }

    private async Task WriteError(int status, string code, string message)
    {
        if (Response.HasStarted) return;
        Response.StatusCode = status;
        await Response.WriteAsJsonAsync(new { error = new { code, message } });
    }
}
