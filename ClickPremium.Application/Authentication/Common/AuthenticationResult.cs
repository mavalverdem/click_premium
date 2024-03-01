using ClickPremium.Domain.Entities;

namespace ClickPremium.Application.Authentication.Common {
    public record AuthenticationResult(
        User User,
        string Token
    );
}
