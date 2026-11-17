namespace SSL_Core_Web_App_Identify_Sample.Auth;

/// <summary>
/// *** THE SWAP POINT for form-based (user name + password) sign-in. ***
/// Implement this against the customer's directory / user store / API and register it
/// (Authentication:Provider = Custom). Nothing else in the application changes.
/// Redirect-based identity providers (OIDC) do not use this; see Auth/Oidc.
/// </summary>
public interface ICredentialAuthenticator
{
    Task<CredentialAuthenticationResult> AuthenticateAsync(
        string userName, string password, CancellationToken cancellationToken = default);
}

/// <summary>Registered when the active provider is redirect-based (or Custom with nothing registered yet).</summary>
internal sealed class DisabledCredentialAuthenticator : ICredentialAuthenticator
{
    public Task<CredentialAuthenticationResult> AuthenticateAsync(
        string userName, string password, CancellationToken cancellationToken = default) =>
        Task.FromResult(CredentialAuthenticationResult.Fail(CredentialFailure.Unavailable));
}
