public class ExampleController : Controller
{
    private static string TargetDirectory = "/example/directory/";

    public void ExtractEntry(IEnumerator<ZipArchiveEntry> entriesEnumerator)
    {
        ZipArchiveEntry entry = entriesEnumerator.Current;
        string destinationPath = Path.Combine(TargetDirectory, entry.FullName);

        entry.ExtractToFile(destinationPath);
    }
}