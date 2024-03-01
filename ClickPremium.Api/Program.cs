using ClickPremium.Api.Common.Errors;
using ClickPremium.Application;
using ClickPremium.Infrastructure;
using ClickPremium.Infrastructure.Models.cfg;
using ClickPremium.Infrastructure.Models.data;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

var builder = WebApplication.CreateBuilder(args);
{
    //builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
    builder.Services
        .AddPresentation()
        .AddApplication()
        .AddInfrastructure(builder.Configuration);
    //builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
    
    builder.Services.AddDbContext<ClickpremiumcfgContext>(options => options.UseSqlServer("Server=.;Database=CLICKPREMIUMCFG;User Id=sa;Password=P@ssw0rd@DB1;MultipleActiveResultSets=true;Trusted_Connection=True;Integrated security=False;Encrypt=False"));
    builder.Services.AddDbContext<ClickpremSysmaplaContext>(options => options.UseSqlServer(builder.Configuration.GetConnectionString("conStringData")));
 
}

var app = builder.Build();
{
    app.UseExceptionHandler("/error");

    app.UseHttpsRedirection();
    app.UseAuthentication();
    app.MapControllers();
    app.Run();
}

public class AppSettings 
{
    private readonly IConfiguration _configuration;

    public AppSettings(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    public string ConnectionString => _configuration.GetConnectionString("Clickpremiumcfg") ?? string.Empty;
}
