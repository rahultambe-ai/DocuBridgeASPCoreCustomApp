using Azure.Identity;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SSL_Core_Web_App_Identify_Sample.Auth;
using SSL_Core_Web_App_Identify_Sample.DmsApp;

namespace SSL_Core_Web_App_Identify_Sample.Controllers;

/// <summary>
/// /api/graph/users?...  ->  https://graph.microsoft.com/v1.0/users?...  with the application's identity.
/// /api/graph/me/...     ->  /users/{signed-in user's e-mail}/...   ("me" does not exist for an app).
/// </summary>
[ApiController]
[Authorize(Policy = AppPolicies.DmsUser, AuthenticationSchemes = AppAuthDefaults.CookieScheme)]
[IgnoreAntiforgeryToken]
[RequireAjaxHeader]
[Route("api/graph/{**path}")]
public sealed class DmsGraphProxyController : ControllerBase
{
    private readonly DmsProxy _proxy;
    private readonly ICurrentUser _user;
    private readonly ILogger<DmsGraphProxyController> _logger;

    public DmsGraphProxyController(DmsProxy proxy, ICurrentUser user, ILogger<DmsGraphProxyController> logger)
    {
        _proxy = proxy;
        _user = user;
        _logger = logger;
    }

    [AcceptVerbs("GET", "POST")]
    public async Task Forward(string? path, CancellationToken ct)
    {
        try
        {
            var raw = DmsProxy.ExtractRawRelative(HttpContext, "/api/graph/");

            var meIdentity = !string.IsNullOrWhiteSpace(_user.Email)
                ? _user.Email
                : (_user.UserName != null && _user.UserName.Contains('@') ? _user.UserName : null);

            var target = DmsProxy.BuildGraphUri(raw, meIdentity);

            // Reading is open to DMS users; the only thing the app writes through Graph is a Teams chat message.
            if (!HttpMethods.IsGet(Request.Method)
                && !DmsProxy.GraphRoot(target).Equals("chats", StringComparison.OrdinalIgnoreCase))
            {
                await WriteError(StatusCodes.Status405MethodNotAllowed, "method_not_allowed",
                    "Only read access is exposed for this Graph endpoint.");
                return;
            }

            var token = await _proxy.GetGraphTokenAsync(ct);
            await _proxy.ForwardAsync(HttpContext, target, token, ct);
        }
        catch (BadHttpRequestException ex)
        {
            await WriteError(StatusCodes.Status400BadRequest, "bad_request", ex.Message);
        }
        catch (AuthenticationFailedException ex)
        {
            _logger.LogError(ex, "The app could not obtain an access token for Microsoft Graph.");
            await WriteError(StatusCodes.Status502BadGateway, "app_token_failed",
                "The application could not sign in to Microsoft Entra ID (see server log).");
        }
        catch (HttpRequestException ex)
        {
            await WriteError(StatusCodes.Status502BadGateway, "upstream_error", ex.Message);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            _logger.LogError(ex, "DMS Graph proxy failed.");
            await WriteError(StatusCodes.Status502BadGateway, "proxy_failed", "Microsoft Graph could not be reached (see server log).");
        }
    }

    private async Task WriteError(int status, string code, string message)
    {
        if (Response.HasStarted) return;
        Response.StatusCode = status;
        await Response.WriteAsJsonAsync(new { error = new { code, message } });
    }
}
