using SSL_Core_Web_App_Identify_Sample.Auth;
using SSL_Core_Web_App_Identify_Sample.DmsApp;
using SSL_Core_Web_App_Identify_Sample.SharePoint;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllersWithViews(options =>
{
    // CSRF protection on every POST/PUT/DELETE without having to remember the attribute.
    options.Filters.Add(new Microsoft.AspNetCore.Mvc.AutoValidateAntiforgeryTokenAttribute());
});

// ---- Identity: the ONLY line that selects how users sign in (Authentication:Provider in appsettings). ----
builder.Services.AddAppAuthentication(builder.Configuration);

// Authentication:Provider = Custom? Register the customer's authenticator here:
// builder.Services.AddScoped<ICredentialAuthenticator, MyCustomerAuthenticator>();

// ---- SharePoint: app-only access, independent of how users sign in. ----
builder.Services.AddSharePointIntegration(builder.Configuration);

// ---- DMS: the React app (ClientApp -> wwwroot/dms) is served at /Dms after sign-in; its SharePoint / Graph calls
// ---- are forwarded by the server with the app's own identity (see Controllers/Dms*.cs).
builder.Services.AddDmsPortal(builder.Configuration);

builder.WebHost.ConfigureKestrel(o => o.Limits.MaxRequestBodySize = 500L * 1024 * 1024); // 500 MB document uploads

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();

// The built React shell is served by DmsController (sign-in required); keep the raw file out of the static-file path.
app.Use(async (context, next) =>
{
    if (context.Request.Path.Equals("/dms/index.html", StringComparison.OrdinalIgnoreCase))
    {
        context.Response.Redirect("/Dms");
        return;
    }
    await next();
});

app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
