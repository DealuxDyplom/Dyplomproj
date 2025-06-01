[ApiController]
[Route("/")]
public class StacktraceController : ControllerBase
{
    [HttpGet("Exception")]
    public string ExceptionEndpoint()
    {
        try {
            throw new InvalidOperationException(ExceptionMessage);
        }
        catch (Exception ex) {
            return ex.StackTrace; // Noncompliant
        }
        return "Ok";
    }
}