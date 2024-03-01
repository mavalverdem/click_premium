using ClickPremium.Domain.Entities;

namespace ClickPremium.Application.Common.Interfaces.Authentication
{
    public interface IJwtTokenGenerator
    {
        string GenerateToken(User user);
    }
}