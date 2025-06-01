using System.Web;

public async Task InvokeAsync(HttpContext context)
{
    context.Response.Headers.ContentSecurityPolicy = "script-src 'self' 'unsafe-inline';"; // Noncompliant
}