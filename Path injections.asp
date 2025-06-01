public class ExampleController : Controller
{
    private static string TargetDirectory = "/path/to/target/directory/";

    public void Example(string filename)
    {
        string path = Path.Combine(TargetDirectory, filename);
        System.IO.File.Delete(path);
    }
}