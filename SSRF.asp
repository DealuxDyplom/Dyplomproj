using System.Web;
using System.Web.Mvc;

public class ExampleController: Controller
{
    [HttpGet]
    public IActionResult GetUser(string id)
    {
        string url = "http://example.com/api/user/" + id;
        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url); // Noncompliant

        return Ok();
    }
}