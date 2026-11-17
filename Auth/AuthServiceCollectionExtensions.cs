using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.DependencyInjection.Extensions;
using SSL_Core_Web_App_Identify_Sample.Auth.Local;
using SSL_Core_Web_App_Identify_Sample.Auth.Oidc;

namespace SSL_Core_Web_App_Identify_Sample.Auth;

public static class AuthServiceCollectionExtensions
{
    /// <summary>
    /// One call wires authentication + authorization for whichever provider is configured.
    /// This is the ONLY place that knows which identity implementation is in use.
    /// </summary>
    public static IServiceCollection AddAppAuthentication(
        this IServiceCollection services, IConfiguration configuration)
    {
        var section = configuration.GetSection(AppAuthOptions.SectionName);
        var options = section.Get<AppAuthOptions>() ?? new AppAuthOptions();

        services.Configure<AppAuthOptions>(section);
        services.AddHttpContextAccessor();
        services.AddScoped<ICurrentUser, CurrentUser>();
        services.AddSingleton<IExternalClaimsMapper, ExternalClaimsMapper>();

        var auth = services.AddAuthentication(o =>
            {
                o.DefaultScheme = AppAuthDefaults.CookieScheme;
                // With a redirect-based IdP, unauthenticated requests go straight to it.
                o.DefaultChallengeScheme = options.Provider == AuthProvider.Oidc
                    ? AppAuthDefaults.OidcScheme
                    : AppAuthDefaults.CookieScheme;
            })
            .AddCookie(AppAuthDefaults.CookieScheme, c =>
            {
                c.LoginPath = "/Account/Login";
                c.AccessDeniedPath = "/Account/AccessDenied";
                c.Cookie.Name = "DmsAuth";
                c.Cookie.HttpOnly = true;
                c.Cookie.SecurePolicy = CookieSecurePolicy.Always;   // run with the https launch profile
                c.Cookie.SameSite = SameSiteMode.Lax;
                c.ExpireTimeSpan = TimeSpan.FromHours(options.SessionHours);
                c.SlidingExpiration = true;

                // The DMS React app calls /api/* with fetch(). A redirect to the HTML login page would be
                // followed silently and look like a successful (but wrong) answer, so API calls get a plain status code.
                c.Events.OnRedirectToLogin = ctx =>
                {
                    if (ctx.Request.Path.StartsWithSegments("/api"))
                        ctx.Response.StatusCode = StatusCodes.Status401Unauthorized;
                    else
                        ctx.Response.Redirect(ctx.RedirectUri);
                    return Task.CompletedTask;
                };
                c.Events.OnRedirectToAccessDenied = ctx =>
                {
                    if (ctx.Request.Path.StartsWithSegments("/api"))
                        ctx.Response.StatusCode = StatusCodes.Status403Forbidden;
                    else
                        ctx.Response.Redirect(ctx.RedirectUri);
                    return Task.CompletedTask;
                };
            });

        switch (options.Provider)
        {
            case AuthProvider.LocalIdentity:
                services.AddLocalIdentityAuthentication(configuration);
                break;

            case AuthProvider.Oidc:
                auth.AddOidcAuthentication(options);
                services.TryAddScoped<ICredentialAuthenticator, DisabledCredentialAuthenticator>();
                break;

            case AuthProvider.Custom:
                // The host registers its own ICredentialAuthenticator AFTER calling this method:
                //   builder.Services.AddScoped<ICredentialAuthenticator, MyCustomerAuthenticator>();
                // Until it does, sign-in reports "unavailable" instead of crashing.
                services.TryAddScoped<ICredentialAuthenticator, DisabledCredentialAuthenticator>();
                break;

            default:
                throw new InvalidOperationException($"Unsupported Authentication:Provider '{options.Provider}'.");
        }

        // Secure by default: every endpoint requires a signed-in user unless it has [AllowAnonymous].
        services.AddAuthorizationBuilder()
            .SetFallbackPolicy(new AuthorizationPolicyBuilder().RequireAuthenticatedUser().Build())
            .AddPolicy(AppPolicies.DmsUser, p => p.RequireRole(AppRoles.DmsUser, AppRoles.DmsAdmin))
            .AddPolicy(AppPolicies.DmsAdmin, p => p.RequireRole(AppRoles.DmsAdmin));

        return services;
    }
}
