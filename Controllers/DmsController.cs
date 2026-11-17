using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SSL_Core_Web_App_Identify_Sample.Auth;

namespace SSL_Core_Web_App_Identify_Sample.Controllers;

/// <summary>
/// Serves the DMS React app (built into wwwroot/dms) to signed-in users only.
/// The user signed in on the portal's own login page; the React app never shows a Microsoft sign-in.
/// </summary>
[Authorize(Policy = AppPolicies.DmsUser)]
public class DmsController : Controller
{
    private readonly IWebHostEnvironment _env;

    public DmsController(IWebHostEnvironment env) => _env = env;

    public IActionResult Index()
    {
        var index = _env.WebRootFileProvider.GetFileInfo("dms/index.html");
        if (!index.Exists || string.IsNullOrEmpty(index.PhysicalPath))
            return View("NotBuilt");

        Response.Headers["Cache-Control"] = "no-store";
        return PhysicalFile(index.PhysicalPath, "text/html; charset=utf-8");
    }
}
