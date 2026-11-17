using Microsoft.Extensions.Options;
using System.Security.Claims;

namespace SSL_Core_Web_App_Identify_Sample.Auth;

/// <summary>Converts a principal issued by a redirect-based provider (OIDC today; SAML/WS-Fed later)
/// into the provider-neutral <see cref="AuthenticatedUser"/>, driven entirely by configuration.</summary>
public interface IExternalClaimsMapper
{
    AuthenticatedUser? Map(ClaimsPrincipal external);
}

internal sealed class ExternalClaimsMapper : IExternalClaimsMapper
{
    private readonly AppAuthOptions _options;

    public ExternalClaimsMapper(IOptions<AppAuthOptions> options) => _options = options.Value;

    public AuthenticatedUser? Map(ClaimsPrincipal external)
    {
        var m = _options.Oidc.ClaimMapping;

        string? Get(string type) => external.FindFirst(type)?.Value;

        var subject = Get(m.SubjectClaim);
        if (string.IsNullOrWhiteSpace(subject))
            return null;

        var email = Get(m.EmailClaim);
        var userName = Get(m.UserNameClaim) ?? email ?? subject;

        IEnumerable<string> roles = external.FindAll(m.RolesClaim).Select(c => c.Value);

        if (_options.RoleMap.Count > 0)
        {
            var map = new Dictionary<string, string>(_options.RoleMap, StringComparer.OrdinalIgnoreCase);
            roles = roles.Where(map.ContainsKey).Select(r => map[r]);
        }

        return new AuthenticatedUser(
            subject,
            userName,
            email,
            Get(m.DisplayNameClaim) ?? userName,
            roles.ToArray());
    }
}
