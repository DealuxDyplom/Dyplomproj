using System.Diagnostics;

public class ExampleController : Controller
{
    public void Example(string name, string value)
    {
        Process proc = new Process();
        proc.StartInfo.FileName = "path/to/executable";
        proc.StartInfo.EnvironmentVariables.Add(name, value); // Noncompliant
        proc.Start();
    }
}