using Microsoft.AspNetCore.Mvc;

namespace ClickPremium.Api.Controllers;
[Route("[controller]")]
public class CompaniasController : ApiController
{
    [HttpGet]
    public IActionResult ListarCompanias()
    {
        return Ok(Array.Empty<string>());
    }
}