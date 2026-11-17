using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;

namespace SSL_Core_Web_App_Identify_Sample.Auth.Local;

internal static class LocalIdentityServiceCollectionExtensions
{
    /// <summary>Everything the Local provider needs. Removing the customer-replaced identity = deleting this call + this folder.</summary>
    public static IServiceCollection AddLocalIdentityAuthentication(
        this IServiceCollection services, IConfiguration configuration)
    {
        services.AddDbContext<ApplicationDbContext>(options =>
            options.UseSqlServer(configuration.GetConnectionString("DefaultConnection")));

        // AddIdentityCore (not AddIdentity): we only want users/roles/password hashing.
        // The cookie + sign-in is owned by the app, not by Identity, so it works the same for every provider.
        services.AddIdentityCore<ApplicationUser>(options =>
            {
                options.Password.RequiredLength = 10;
                options.Password.RequireDigit = true;
                options.Password.RequireLowercase = true;
                options.Password.RequireUppercase = true;
                options.Password.RequireNonAlphanumeric = false;

                options.Lockout.MaxFailedAccessAttempts = 5;
                options.Lockout.DefaultLockoutTimeSpan = TimeSpan.FromMinutes(15);

                options.User.RequireUniqueEmail = true;
            })
            .AddRoles<IdentityRole>()
            .AddEntityFrameworkStores<ApplicationDbContext>();

        services.AddScoped<ICredentialAuthenticator, LocalIdentityCredentialAuthenticator>();
        services.AddHostedService<LocalIdentityStartupTask>();

        return services;
    }
}
