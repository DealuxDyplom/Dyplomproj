using System;
using System.Security.Cryptography;

public void encrypt()
{
    var RsaCsp = new RSACryptoServiceProvider(); // Noncompliant
}

using System;
using System.Security.Cryptography;

public void encrypt()
{
    var DsaCsp = new DSACryptoServiceProvider(); // Noncompliant
}

using System;
using System.Security.Cryptography;

public void encrypt()
{
    ECDsa ecdsa = ECDsa.Create(ECCurve.NamedCurves.brainpoolP160t1); // Noncompliant
}

