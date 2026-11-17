namespace SSL_Core_Web_App_Identify_Sample.Auth;

public enum AuthProvider
{
    /// <summary>ASP.NET Core Identity + SQL Server (the sample identity). Form login.</summary>
    LocalIdentity,

    /// <summary>Customer's OpenID Connect provider (Keycloak, Okta, Auth0, ADFS, Ping ...). Redirect login.</summary>
    Oidc,

    /// <summary>You register your own ICredentialAuthenticator (LDAP, customer API, legacy DB ...). Form login.</summary>
    Custom
}

public sealed class AppAuthOptions
{
    public const string SectionName = "Authentication";

    public AuthProvider Provider { get; set; } = AuthProvider.LocalIdentity;
    public int SessionHours { get; set; } = 8;

    /// <summary>External role/group name -> app role. Empty = pass roles through unchanged.</summary>
    public Dictionary<string, string> RoleMap { get; set; } = new();

    public OidcAuthOptions Oidc { get; set; } = new();
    public LocalAuthOptions Local { get; set; } = new();
}

public sealed class OidcAuthOptions
{
    public string Authority { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string? ClientSecret { get; set; }
    public string Scope { get; set; } = "openid profile email";
    public string CallbackPath { get; set; } = "/signin-oidc";
    public string SignedOutCallbackPath { get; set; } = "/signout-callback-oidc";
    public bool RequireHttpsMetadata { get; set; } = true;
    public ClaimMappingOptions ClaimMapping { get; set; } = new();
}

/// <summary>Which claim in the customer's token carries what. Change values, not code.</summary>
public sealed class ClaimMappingOptions
{
    public string SubjectClaim { get; set; } = "sub";
    public string UserNameClaim { get; set; } = "preferred_username";
    public string EmailClaim { get; set; } = "email";
    public string DisplayNameClaim { get; set; } = "name";
    public string RolesClaim { get; set; } = "roles";
}

public sealed class LocalAuthOptions
{
    /// <summary>Development only. See appsettings.Development.json.</summary>
    public List<LocalSeedUser> SeedUsers { get; set; } = new();
}

public sealed class LocalSeedUser
{
    public string UserName { get; set; } = "";
    public string? Email { get; set; }
    public string? DisplayName { get; set; }
    public string Password { get; set; } = "";
    public List<string> Roles { get; set; } = new();
}
