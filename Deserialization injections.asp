public class Example : Controller
{
    [HttpPost]
    public ActionResult Deserialize(HttpPostedFileBase inputFile)
    {
        ExpectedType expectedObject = null;
        var formatter               = new BinaryFormatter();
        expectedObject              = (ExpectedType)formatter.Deserialize(inputFile.InputStream);
    }
}