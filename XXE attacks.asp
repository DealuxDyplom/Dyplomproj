using System.Xml;

public static void decode()
{
    XmlDocument parser = new XmlDocument();
    parser.XmlResolver = new XmlUrlResolver(); // Noncompliant
    parser.LoadXml("xxe.xml");
}