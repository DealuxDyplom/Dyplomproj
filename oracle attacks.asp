using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("Example")]
public class ExampleController : ControllerBase
{
    private const string TargetDirectory = "/path/to/target/directory";

    public IActionResult FileExists()
    {
        string file = Request.Query["file"];
        string path = TargetDirectory + file;
        if (!System.IO.File.Exists(path)) { // Noncompliant
            return NotFound("File not found");
        }
        return Ok("File found");
    }
}