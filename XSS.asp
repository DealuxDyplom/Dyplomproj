using System.Web;
using System.Web.Mvc;

public class HelloController : Controller
{
    [HttpGet]
    public void Hello(string name, HttpResponse response)
    {
        string html = "<h1>Hello"+ name +"</h1>"
        response.Write(html);
    }
}