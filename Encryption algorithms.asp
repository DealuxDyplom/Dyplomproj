using System.Security.Cryptography;

public void encrypt()
{
    AesManaged aes = new AesManaged
    {
        keysize = 128,
        blocksize = 128,
        mode = ciphermode.ecb,        // Noncompliant
        padding = paddingmode.pkcs7
    };
}