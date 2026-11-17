using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using SSL_Core_Web_App_Identify_Sample.Auth;
using SSL_Core_Web_App_Identify_Sample.Models;

namespace SSL_Core_Web_App_Identify_Sample.Controllers;

/// <summary>
/// Provider-agnostic sign-in/out. It talks only to ICredentialAuthenticator (form providers)
/// or the OIDC scheme (redirect providers). It never references ASP.NET Identity.
/// </summary>
[AllowAnonymous]
public class AccountController : Controller
{
    private readonly ICredentialAuthenticator _authenticator;
    private readonly AppAuthOptions _options;
    private readonly ILogger<AccountController> _logger;

    public AccountController(
        ICredentialAuthenticator authenticator,
        IOptions<AppAuthOptions> options,
        ILogger<AccountController> logger)
    {
        _authenticator = authenticator;
        _options = options.Value;
        _logger = logger;
    }

    private bool IsRedirectProvider => _options.Provider == AuthProvider.Oidc;

    [HttpGet]
    public IActionResult Login(string? returnUrl = null)
    {
        if (User.Identity?.IsAuthenticated == true)
            return RedirectToLocal(returnUrl);

        if (IsRedirectProvider)
        {
            return Challenge(
                new AuthenticationProperties { RedirectUri = SafeReturnUrl(returnUrl) },
                AppAuthDefaults.OidcScheme);
        }

        return View(new LoginViewModel { ReturnUrl = returnUrl });
    }

    [HttpPost]
    public async Task<IActionResult> Login(LoginViewModel model)
    {
        if (IsRedirectProvider)
            return RedirectToAction(nameof(Login), new { returnUrl = model.ReturnUrl });

        if (!ModelState.IsValid)
            return View(model);

        var result = await _authenticator.AuthenticateAsync(
            model.UserName.Trim(), model.Password, HttpContext.RequestAborted);

        if (!result.Succeeded || result.User is null)
        {
            _logger.LogWarning("Failed sign-in attempt. Reason: {Reason}", result.Failure);

            ModelState.AddModelError(string.Empty, result.Failure switch
            {
                CredentialFailure.LockedOut => "Too many failed attempts. Please try again later.",
                CredentialFailure.Unavailable => "Sign-in is currently unavailable. Please contact your administrator.",
                _ => "Invalid user name or password."
            });

            return View(model);
        }

        var principal = AppPrincipalFactory.Create(result.User, _options.Provider.ToString());

        await HttpContext.SignInAsync(
            AppAuthDefaults.CookieScheme,
            principal,
            new AuthenticationProperties { IsPersistent = model.RememberMe });

        _logger.LogInformation("User {UserId} signed in via {Provider}.", result.User.SubjectId, _options.Provider);

        return RedirectToLocal(model.ReturnUrl);
    }

    [HttpPost]
    public async Task<IActionResult> Logout()
    {
        if (IsRedirectProvider)
        {
            // Clears the app cookie, then ends the session at the IdP (RP-initiated logout).
            return SignOut(
                new AuthenticationProperties { RedirectUri = Url.Action(nameof(LoggedOut)) ?? "/" },
                AppAuthDefaults.CookieScheme,
                AppAuthDefaults.OidcScheme);
        }

        await HttpContext.SignOutAsync(AppAuthDefaults.CookieScheme);
        return RedirectToAction(nameof(LoggedOut));
    }

    [HttpGet] public IActionResult LoggedOut() => View();
    [HttpGet] public IActionResult AccessDenied() => View();
    [HttpGet] public IActionResult LoginFailed() => View();

    private string SafeReturnUrl(string? returnUrl) =>
        !string.IsNullOrEmpty(returnUrl) && Url.IsLocalUrl(returnUrl)
            ? returnUrl
            : Url.Action("Index", "Home") ?? "/";

    private IActionResult RedirectToLocal(string? returnUrl) =>
        !string.IsNullOrEmpty(returnUrl) && Url.IsLocalUrl(returnUrl)
            ? LocalRedirect(returnUrl)
            : RedirectToAction("Index", "Home");
}
