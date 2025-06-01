using System.Security.Cryptography;

public void encrypt()
{
    var simpleDES = new DESCryptoServiceProvider(); // Noncompliant
}