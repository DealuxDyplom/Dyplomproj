public class ExampleController : Controller
{
    public void Run(string args)
    {
        Process p             = new Process();
        p.StartInfo.FileName  = "/usr/bin/find";
        p.StartInfo.Arguments = "/some/folder -iname " + args;
        p.Start();
    }
}