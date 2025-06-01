using System.IO;

public void Example()
{
    var tempPath = Path.GetTempFileName();  // Noncompliant

    using (var writer = new StreamWriter(tempPath))
    {
        writer.WriteLine("content");
    }
}