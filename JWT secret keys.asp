public class LoginExampleController : ApiController
{
    private const string key = "SecretSecretSecretSecretSecretSecretSecretSecret";

    public IHttpActionResult Post([FromBody] LoginModel login)
    {
        // Code to validate the login information is omitted

        var securityKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(key));
        var credentials = new SigningCredentials(securityKey, SecurityAlgorithms.HmacSha256);

        var secToken = new JwtSecurityToken(
            "example.com",
            "example.com",
            null,
            expires: DateTime.Now.AddMinutes(120),
            signingCredentials: credentials
        );

        var token = new JwtSecurityTokenHandler().WriteToken(secToken);
        return Ok(token);
    }
}