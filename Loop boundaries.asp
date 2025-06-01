public class ExampleController : Controller
{
    public IActionResult Compute(int data)
    {
        for (int i = 0; i < data; i++) // Noncompliant
        {
            Console.WriteLine("Hello");
        }

        Enumerable
            .Range(1, data) // Noncompliant
            .ToList()
            .ForEach(i => Console.WriteLine("World"));

        return Ok();
    }
}