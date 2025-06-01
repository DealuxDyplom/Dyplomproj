using MongoDB.Driver;
using MongoDB.Bson;

[ApiController]
[Route("Example")]
public class ExampleController: ControllerBase
{
    private string connectionString;

    [Route("Example")]
    public async Task<string> Example()
    {
        var client     = new MongoClient(connectionString);
        var database   = client.GetDatabase("example");
        var collection = database.GetCollection<Message>("messages");

        var filterDefinition = Request.Query["filterDefinition"];

        await collection.FindAsync(filter)
    }
}