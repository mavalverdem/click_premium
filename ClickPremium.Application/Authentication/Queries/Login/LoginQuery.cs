
using ClickPremium.Application.Authentication.Common;
using ErrorOr;
using MediatR;

namespace ClickPremium.Application.Authentication.Queries.Login
{
    public record LoginQuery(string Email, string Password) : IRequest<ErrorOr<AuthenticationResult>>;
}
