


using ClickPremium.Api.Common.Errors;
using ClickPremium.Api.Common.Mapping;
using MediatR;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Microsoft.Extensions.DependencyInjection;

namespace ClickPremium.Application;

public static class DependencyInjection
{
    public static IServiceCollection AddPresentation(this IServiceCollection services)
    {
        services.AddControllers();
        services.AddSingleton<ProblemDetailsFactory, ClickPremiumProblemDetailsFactory>();
        services.AddMapping();


        return services;
    }
}
