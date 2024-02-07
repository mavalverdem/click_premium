

using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using ClickPremium.Application.Common.Interfaces.Authentication;
using Microsoft.IdentityModel.Tokens;

namespace ClickPremium.Infrastructure.Authentication {
    public class JwtTokenGenerator : IJwtTokenGenerator
    {
        public string GenerateToken(Guid userId, string firstName, string lastName)
        {
            var signingCredentials = new SigningCredentials(
                new SymmetricSecurityKey(Encoding.UTF8.GetBytes("Cuando me mira y me toca, todo lo que quiero es tocarla.")),
                SecurityAlgorithms.HmacSha256);
            var claims = new List<Claim>
            {
                new Claim(JwtRegisteredClaimNames.Sub, userId.ToString()),
                new Claim(JwtRegisteredClaimNames.GivenName, firstName),
                new Claim(JwtRegisteredClaimNames.FamilyName, lastName),
                new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString())
            };
            var SecurityToken = new JwtSecurityToken(
                issuer: "ClickPremium",
                expires: DateTime.Now.AddDays(1),
                claims: claims,
                signingCredentials: signingCredentials
            );

            return new JwtSecurityTokenHandler().WriteToken(SecurityToken);
        }
    }
}