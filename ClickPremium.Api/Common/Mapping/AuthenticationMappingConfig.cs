using ClickPremium.Application.Authentication.Commands.Register;
using ClickPremium.Application.Authentication.Common;
using ClickPremium.Application.Authentication.Queries.Login;
using ClickPremium.Contracts.Authentication;
using Mapster;

namespace ClickPremium.Api.Common.Mapping;

public class AuthenticationMappingConfig : IRegister
{
    public void Register(TypeAdapterConfig config)
    {
        config.NewConfig<RegisterRequest, RegisterCommand>();
        config.NewConfig<LoginRequest, LoginQuery>();
        config.NewConfig<AuthenticationResult, AuthenticationResponse>()
            .Map(dest => dest, src => src.User);
    }
}
