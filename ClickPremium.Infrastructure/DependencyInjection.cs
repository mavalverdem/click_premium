


using ClickPremium.Application.Common.Interfaces.Authentication;
using ClickPremium.Infrastructure.Authentication;
using Microsoft.Extensions.DependencyInjection;

namespace ClickPremium.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        services.AddSingleton<IJwtTokenGenerator, JwtTokenGenerator>();
        return services;
    }
}
