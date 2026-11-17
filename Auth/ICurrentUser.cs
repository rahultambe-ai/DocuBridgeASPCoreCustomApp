using System.Security.Claims;

namespace SSL_Core_Web_App_Identify_Sample.Auth;

/// <summary>What controllers/services/views use to learn who is signed in.
/// They must NOT touch UserManager, ApplicationUser, or provider-specific claims.</summary>
public interface ICurrentUser
{
    bool IsAuthenticated { get; }
    string? Id { get; }
    string? UserName { get; }
    string? Email { get; }
    string? DisplayName { get; }
    IReadOnlyCollection<string> Roles { get; }
    bool IsInRole(string role);
}

internal sealed class CurrentUser : ICurrentUser
{
    private readonly IHttpContextAccessor _accessor;

    public CurrentUser(IHttpContextAccessor accessor) => _accessor = accessor;

    private ClaimsPrincipal? Principal => _accessor.HttpContext?.User;

    public bool IsAuthenticated => Principal?.Identity?.IsAuthenticated == true;
    public string? Id => Principal?.FindFirstValue(AppClaimTypes.Subject);
    public string? UserName => Principal?.FindFirstValue(AppClaimTypes.UserName);
    public string? Email => Principal?.FindFirstValue(AppClaimTypes.Email);
    public string? DisplayName => Principal?.FindFirstValue(AppClaimTypes.DisplayName) ?? UserName;

    public IReadOnlyCollection<string> Roles =>
        Principal?.FindAll(AppClaimTypes.Role).Select(c => c.Value).ToArray() ?? Array.Empty<string>();

    public bool IsInRole(string role) => Principal?.IsInRole(role) == true;
}
