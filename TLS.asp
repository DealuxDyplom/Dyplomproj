using System.Net;
using System.Net.Http;

public static void connect()
{
    ServicePointManager.ServerCertificateValidationCallback +=
	 (sender, certificate, chain, errors) => {
	     return true; // Noncompliant
    };

    HttpClient httpClient = new HttpClient();
    HttpResponseMessage response = httpClient.GetAsync("https://example.com").Result;
}