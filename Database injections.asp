public class ExampleController : Controller
{
    private readonly UserAccountContext Context;

    public IActionResult Authenticate(string user, string pass)
    {
        var query = "SELECT * FROM users WHERE user = '" + user + "' AND pass = '" + pass + "'";

        var queryResults = Context
            .Database
            .FromSqlRaw(query);

        if (queryResults == 0)
        {
            return Unauthorized();
        }

        return Ok();
    }
}