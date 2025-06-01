public class ExampleController : Controller
{
    public IActionResult Validate(string regex, string input)
    {
        bool match = Regex.IsMatch(input, regex);

        return Json(match);
    }
}