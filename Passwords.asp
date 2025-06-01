using System.Security.Cryptography;

RNGCryptoServiceProvider rngCsp = new RNGCryptoServiceProvider();
byte[] salt = new byte[32];
rngCsp.GetBytes(salt);
Rfc2898DeriveBytes kdf = new Rfc2898DeriveBytes(password, salt, 100, HashAlgorithmName.SHA256); // Noncompliant
string hashed = Convert.ToBase64String(kdf.GetBytes(256 / 8));

using Microsoft.AspNet.Identity;

string password = "NotSoSecure";
PasswordHasher hasher = new PasswordHasher();
string hash = hasher.HashPassword(password); // Noncompliant