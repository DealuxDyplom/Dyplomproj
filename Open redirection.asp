using System.Web;
using System.Web.Mvc;

public class ExampleController : Controller
{
    [HttpGet]
    public void Redirect(string url)
    {
        Response.Redirect(url);
    }
}