using Microsoft.AspNetCore.Identity;

namespace SSL_Core_Web_App_Identify_Sample.Auth.Local;

/// <summary>Only the Local provider knows this type. Application code uses ICurrentUser instead.</summary>
public class ApplicationUser : IdentityUser
{
    public string? DisplayName { get; set; }
    public string? EmployeeCode { get; set; }
    public string? Department { get; set; }
}
