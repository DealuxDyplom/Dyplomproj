using System.IO;
using System.Security.Cryptography;

public void Encrypt(byte[] key, byte[] dataToEncrypt, MemoryStream target)
{
    var aes = new AesCryptoServiceProvider();

    byte[] iv     = new byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 };
    var encryptor = aes.CreateEncryptor(key, iv); // Noncompliant

    var cryptoStream = new CryptoStream(target, encryptor, CryptoStreamMode.Write);
    var swEncrypt    = new StreamWriter(cryptoStream);

    swEncrypt.Write(dataToEncrypt);
}