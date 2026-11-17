using System.Security.Claims;

namespace SSL_Core_Web_App_Identify_Sample.Auth;

/// <summary>Single place that turns an <see cref="AuthenticatedUser"/> into the cookie principal.
/// Because every provider goes through here, the claims the app sees are always identical.</summary>
public static class AppPrincipalFactory
{
    public static ClaimsPrincipal Create(AuthenticatedUser user, string providerName)
    {
        var claims = new List<Claim>
        {
            new(AppClaimTypes.Subject, user.SubjectId),
            new(AppClaimTypes.UserName, user.UserName),
            new(AppClaimTypes.Provider, providerName)
        };

        if (!string.IsNullOrWhiteSpace(user.Email))
            claims.Add(new Claim(AppClaimTypes.Email, user.Email));

        if (!string.IsNullOrWhiteSpace(user.DisplayName))
            claims.Add(new Claim(AppClaimTypes.DisplayName, user.DisplayName));

        claims.AddRange(user.Roles
            .Where(r => !string.IsNullOrWhiteSpace(r))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(r => new Claim(AppClaimTypes.Role, r)));

        var identity = new ClaimsIdentity(
            claims,
            authenticationType: providerName,
            nameType: AppClaimTypes.UserName,
            roleType: AppClaimTypes.Role);

        return new ClaimsPrincipal(identity);
    }
}
