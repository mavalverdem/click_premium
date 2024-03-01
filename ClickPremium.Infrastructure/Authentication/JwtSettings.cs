using System.Reflection.Metadata;

namespace ClickPremium.Application.Common.Interfaces.Authentication
{
    public class JwtSettings
    {
        public const string SectionName = "JwtSettings";
        public string? Secret { get; init; }

        public int ExpiryMinutes { get; set; }
        public string? Issuer { get; init; }
        public string? Audience { get; init; }



    }
}