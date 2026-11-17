using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;

namespace SSL_Core_Web_App_Identify_Sample.Auth.Local;

/// <summary>Runs once at startup (Local provider only): creates the DB in Development,
/// ensures the app roles exist, and seeds development users.</summary>
internal sealed class LocalIdentityStartupTask : IHostedService
{
    private readonly IServiceProvider _services;
    private readonly IHostEnvironment _env;
    private readonly AppAuthOptions _options;
    private readonly ILogger<LocalIdentityStartupTask> _logger;

    public LocalIdentityStartupTask(
        IServiceProvider services,
        IHostEnvironment env,
        IOptions<AppAuthOptions> options,
        ILogger<LocalIdentityStartupTask> logger)
    {
        _services = services;
        _env = env;
        _options = options.Value;
        _logger = logger;
    }

    public async Task StartAsync(CancellationToken cancellationToken)
    {
        using var scope = _services.CreateScope();
        var sp = scope.ServiceProvider;

        // Production schemas should come from EF migrations, not EnsureCreated.
        if (_env.IsDevelopment())
        {
            var db = sp.GetRequiredService<ApplicationDbContext>();
            await db.Database.EnsureCreatedAsync(cancellationToken);
        }

        var roles = sp.GetRequiredService<RoleManager<IdentityRole>>();
        foreach (var role in AppRoles.All)
        {
            if (!await roles.RoleExistsAsync(role))
                await roles.CreateAsync(new IdentityRole(role));
        }

        if (!_env.IsDevelopment())
            return;

        var users = sp.GetRequiredService<UserManager<ApplicationUser>>();

        foreach (var seed in _options.Local.SeedUsers)
        {
            if (string.IsNullOrWhiteSpace(seed.UserName) || string.IsNullOrWhiteSpace(seed.Password))
                continue;

            var user = await users.FindByNameAsync(seed.UserName);

            if (user is null)
            {
                user = new ApplicationUser
                {
                    UserName = seed.UserName,
                    Email = seed.Email,
                    DisplayName = seed.DisplayName,
                    EmailConfirmed = true
                };

                var created = await users.CreateAsync(user, seed.Password);
                if (!created.Succeeded)
                {
                    _logger.LogWarning("Could not seed user {UserName}: {Errors}",
                        seed.UserName, string.Join("; ", created.Errors.Select(e => e.Description)));
                    continue;
                }
            }

            // Also repairs users created by the earlier sample, which had no roles.
            foreach (var role in seed.Roles.Where(r => AppRoles.All.Contains(r)))
            {
                if (!await users.IsInRoleAsync(user, role))
                    await users.AddToRoleAsync(user, role);
            }
        }
    }

    public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;
}
