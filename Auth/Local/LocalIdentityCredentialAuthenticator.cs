using Microsoft.AspNetCore.Identity;

namespace SSL_Core_Web_App_Identify_Sample.Auth.Local;

/// <summary>Validates a user name + password against ASP.NET Core Identity (SQL Server),
/// including lockout, and returns the provider-neutral <see cref="AuthenticatedUser"/>.</summary>
internal sealed class LocalIdentityCredentialAuthenticator : ICredentialAuthenticator
{
    private readonly UserManager<ApplicationUser> _users;

    public LocalIdentityCredentialAuthenticator(UserManager<ApplicationUser> users) => _users = users;

    public async Task<CredentialAuthenticationResult> AuthenticateAsync(
        string userName, string password, CancellationToken cancellationToken = default)
    {
        var user = await _users.FindByNameAsync(userName)
                   ?? await _users.FindByEmailAsync(userName);

        // Same failure for "no such user" and "wrong password" -> no user enumeration.
        if (user is null)
            return CredentialAuthenticationResult.Fail(CredentialFailure.InvalidCredentials);

        if (await _users.IsLockedOutAsync(user))
            return CredentialAuthenticationResult.Fail(CredentialFailure.LockedOut);

        if (!await _users.CheckPasswordAsync(user, password))
        {
            // CheckPasswordAsync alone does NOT count failures; without this there is no brute-force protection.
            await _users.AccessFailedAsync(user);
            return CredentialAuthenticationResult.Fail(CredentialFailure.InvalidCredentials);
        }

        await _users.ResetAccessFailedCountAsync(user);

        var roles = await _users.GetRolesAsync(user);

        return CredentialAuthenticationResult.Success(new AuthenticatedUser(
            SubjectId: user.Id,
            UserName: user.UserName ?? userName,
            Email: user.Email,
            DisplayName: user.DisplayName ?? user.UserName,
            Roles: roles.ToArray()));
    }
}
