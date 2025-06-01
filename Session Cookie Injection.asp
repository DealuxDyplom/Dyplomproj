using System.Web;
using System.Web.Mvc;

public class ExampleController : Controller
{
    [HttpGet]
    public IActionResult CheckCookie(string cookie)
    {
        if (Request.Cookies["ASP.NET_SessionId"] == null)
        {
            Response.Cookies.Append("ASP.NET_SessionId", cookie);
        }

        return View("Welcome");
    }
}