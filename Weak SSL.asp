using System.Net;

public void encrypt()
{
    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls; // Noncompliant
}

using System.Net.Http;
using System.Security.Authentication;

public void encrypt()
{
    new HttpClientHandler
    {
        SslProtocols = SslProtocols.Tls // Noncompliant
    };
}