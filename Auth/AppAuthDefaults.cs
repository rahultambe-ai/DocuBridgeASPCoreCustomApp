using System.Security.Claims;

namespace SSL_Core_Web_App_Identify_Sample.Auth;

/// <summary>Scheme names used by the app. The app only ever signs in to ONE cookie scheme,
/// regardless of which identity provider authenticated the user.</summary>
public static class AppAuthDefaults
{
    public const string CookieScheme = "AppCookie";
    public const string OidcScheme = "Oidc";
}

/// <summary>Roles THIS APPLICATION understands. Identity providers are mapped onto these
/// (see Authentication:RoleMap) so the rest of the code never sees provider-specific roles.</summary>
public static class AppRoles
{
    public const string DmsUser = "DmsUser";     // can browse + download
    public const string DmsAdmin = "DmsAdmin";   // can also create folders / run diagnostics

    public static readonly string[] All = [DmsUser, DmsAdmin];
}

public static class AppPolicies
{
    public const string DmsUser = "DmsUser";     // DmsUser OR DmsAdmin
    public const string DmsAdmin = "DmsAdmin";
}

/// <summary>The only claim types application code is allowed to rely on.</summary>
public static class AppClaimTypes
{
    public const string Subject = ClaimTypes.NameIdentifier;
    public const string UserName = ClaimTypes.Name;
    public const string Email = ClaimTypes.Email;
    public const string Role = ClaimTypes.Role;
    public const string DisplayName = "display_name";
    public const string Provider = "auth_provider";
}
