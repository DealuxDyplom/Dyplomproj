public class ExampleController : Controller
{
    [HttpGet]
    public IActionResult Authenticate(string user, string pass)
    {
        XmlDocument doc = new XmlDocument();

        String expression = "/users/user[@name='" + user + "' and @pass='" + pass + "']";

        return Json(doc.SelectSingleNode(expression) != null);
    }
}