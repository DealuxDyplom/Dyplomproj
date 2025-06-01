XmlDocument xmlDoc = new()
{
    PreserveWhitespace = true
};
xmlDoc.Load("/data/login.xml");
SignedXml signedXml = new(xmlDoc);
XmlNodeList nodeList = xmlDoc.GetElementsByTagName("Signature");
signedXml.LoadXml((XmlElement?)nodeList[0]);
if (signedXml.CheckSignature()) {
    // Process the XML content
} else {
    // Raise an error
}