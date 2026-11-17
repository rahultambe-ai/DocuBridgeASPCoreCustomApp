using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace SSL_Core_Web_App_Identify_Sample.Controllers;

public class HomeController : Controller
{
    // Requires sign-in via the app-wide fallback policy.
    public IActionResult Index() => View();

    [AllowAnonymous]
    [ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
    public IActionResult Error() => View();
}
