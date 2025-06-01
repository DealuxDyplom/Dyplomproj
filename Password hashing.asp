using System.Security.Cryptography;

public static void hash(string password)
{
    var salt = Encoding.UTF8.GetBytes("salty");
    var hashed = new Rfc2898DeriveBytes(password, salt); // Noncompliant
}