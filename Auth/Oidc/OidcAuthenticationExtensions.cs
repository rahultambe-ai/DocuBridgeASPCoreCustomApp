using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;

namespace SSL_Core_Web_App_Identify_Sample.Auth.Oidc;

internal static class OidcAuthenticationExtensions
{
    /// <summary>Generic OpenID Connect sign-in against the CUSTOMER's identity provider.
    /// Point Authority at their IdP. (Do not point it at Microsoft Entra if users must never see a Microsoft sign-in.)</summary>
    public static AuthenticationBuilder AddOidcAuthentication(this AuthenticationBuilder builder, AppAuthOptions options)
    {
        var oidc = options.Oidc;

        if (string.IsNullOrWhiteSpace(oidc.Authority) || string.IsNullOrWhiteSpace(oidc.ClientId))
            throw new InvalidOperationException(
                "Authentication:Oidc:Authority and Authentication:Oidc:ClientId are required when Authentication:Provider is 'Oidc'.");

        return builder.AddOpenIdConnect(AppAuthDefaults.OidcScheme, o =>
        {
            o.SignInScheme = AppAuthDefaults.CookieScheme;   // the SAME app cookie as every other provider

            o.Authority = oidc.Authority;
            o.ClientId = oidc.ClientId;
            o.ClientSecret = oidc.ClientSecret;
            o.RequireHttpsMetadata = oidc.RequireHttpsMetadata;
            o.CallbackPath = oidc.CallbackPath;
            o.SignedOutCallbackPath = oidc.SignedOutCallbackPath;

            o.ResponseType = OpenIdConnectResponseType.Code;   // authorization-code flow
            o.UsePkce = true;
            o.GetClaimsFromUserInfoEndpoint = true;
            o.SaveTokens = false;          // the app never needs the user's tokens
            o.MapInboundClaims = false;    // keep the IdP's raw claim names so ClaimMapping config is predictable

            o.Scope.Clear();
            foreach (var scope in oidc.Scope.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                o.Scope.Add(scope);

            o.Events = new OpenIdConnectEvents
            {
                // Last hook before the cookie is issued: replace the IdP principal with the app's neutral principal.
                OnTicketReceived = ctx =>
                {
                    var mapper = ctx.HttpContext.RequestServices.GetRequiredService<IExternalClaimsMapper>();
                    var user = ctx.Principal is null ? null : mapper.Map(ctx.Principal);

                    if (user is null)
                    {
                        ctx.Fail("The identity provider did not return the configured subject claim.");
                        return Task.CompletedTask;
                    }

                    ctx.Principal = AppPrincipalFactory.Create(user, AppAuthDefaults.OidcScheme);
                    return Task.CompletedTask;
                },

                OnRemoteFailure = ctx =>
                {
                    var logger = ctx.HttpContext.RequestServices
                        .GetRequiredService<ILoggerFactory>().CreateLogger("Auth.Oidc");
                    logger.LogWarning(ctx.Failure, "OIDC sign-in failed.");

                    ctx.Response.Redirect("/Account/LoginFailed");
                    ctx.HandleResponse();
                    return Task.CompletedTask;
                }
            };
        });
    }
}
