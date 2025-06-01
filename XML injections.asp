using System.Xml;

public class ExampleController : Controller
{
    public async void Example(string username)
    {
        XmlWriter writer = XmlWriter.Create("data.xml");
        await writer.WriteRawAsync(
            $@"<user>
                <username>{username}</username> <!-- Noncompliant -->
                <role>user</role>
            </user>"
        );
        await writer.DisposeAsync();
    }
}