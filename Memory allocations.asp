[Route("NonCompliantArrayList")]
public string NonCompliantArrayList()
{
    int size;
    try
    {
        size = int.Parse(Request.Query["size"]);
    }
    catch (FormatException)
    {
        return "Number format exception while reading size";
    }
    ArrayList arrayList = new ArrayList(size); // Noncompliant
    return size + " bytes were allocated.";
}