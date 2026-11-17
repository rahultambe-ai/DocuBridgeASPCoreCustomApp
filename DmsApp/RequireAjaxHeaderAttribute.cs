using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;

namespace SSL_Core_Web_App_Identify_Sample.DmsApp;

/// <summary>
/// CSRF defence for the DMS JSON/proxy endpoints: every request that changes something must carry the custom
/// header "X-Requested-With: XMLHttpRequest". A foreign web site cannot add a custom header to a cross-site
/// request without a CORS pre-flight, and this app never allows one.
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Method)]
public sealed class RequireAjaxHeaderAttribute : ActionFilterAttribute
{
    public override void OnActionExecuting(ActionExecutingContext context)
    {
        var method = context.HttpContext.Request.Method;
        if (HttpMethods.IsGet(method) || HttpMethods.IsHead(method) || HttpMethods.IsOptions(method))
            return;

        var header = context.HttpContext.Request.Headers["X-Requested-With"].ToString();
        if (!string.Equals(header, "XMLHttpRequest", StringComparison.Ordinal))
        {
            context.Result = new BadRequestObjectResult(new
            {
                error = new { code = "missing_header", message = "The X-Requested-With header is required." }
            });
        }
    }
}
