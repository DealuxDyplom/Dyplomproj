public class ExampleController : Controller
{
    public void Run(string binary)
    {
        Process p = new Process();
        p.StartInfo.FileName = binary;
        p.Start();
    }
}