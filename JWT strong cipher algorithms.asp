using JWT;

public static void decode(IJwtDecoder decoder)
{
    decoder.Decode(token, secret, verify: false); // Noncompliant
}
using JWT;

public static void decode()
{
    var jwt = new JwtBuilder()
        .WithSecret(secret)
        .Decode(token); // Noncompliant
}