namespace SSL_Core_Web_App_Identify_Sample.Auth;

/// <summary>
/// Provider-neutral description of a signed-in person. Every identity source
/// (ASP.NET Identity, OIDC, LDAP, a customer API ...) must produce exactly this.
/// </summary>
public sealed record AuthenticatedUser(
    string SubjectId,
    string UserName,
    string? Email,
    string? DisplayName,
    IReadOnlyCollection<string> Roles);

public enum CredentialFailure
{
    None = 0,
    InvalidCredentials,
    LockedOut,
    Unavailable
}

public sealed class CredentialAuthenticationResult
{
    public bool Succeeded { get; private init; }
    public AuthenticatedUser? User { get; private init; }
    public CredentialFailure Failure { get; private init; }

    public static CredentialAuthenticationResult Success(AuthenticatedUser user) =>
        new() { Succeeded = true, User = user, Failure = CredentialFailure.None };

    public static CredentialAuthenticationResult Fail(CredentialFailure failure) =>
        new() { Succeeded = false, Failure = failure };
}
