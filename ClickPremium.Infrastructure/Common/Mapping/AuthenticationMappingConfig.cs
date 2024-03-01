using ClickPremium.Application.Authentication.Commands.Register;

using ClickPremium.Domain.Entities;
using ClickPremium.Infrastructure.Models;
using Mapster;

namespace ClickPremium.Infrastructure.Common.Mapping;

public class AuthenticationMappingConfig : IRegister
{
    public void Register(TypeAdapterConfig config)
    {
        config.NewConfig<Sgusr, User>()
            .Map(
                dest => dest.Email,
                src => src.Codusr
            )
            .Map(
                dest => dest.Password,
                src => src.Clausr
            )
            .Map(
                dest => dest.FirstName,
                src => src.Nomusr
            );
    }
}
