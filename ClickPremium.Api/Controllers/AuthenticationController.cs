using ClickPremium.Application.Services.Authentication;
using ClickPremium.Contracts.Authentication;
using Microsoft.AspNetCore.Mvc;

namespace  ClickPremium.Api.Controllers;

[ApiController]
[Route("auth")]
public class AuthenticactionController: ControllerBase {

    private readonly IAuthenticationService _authenticationService;

    public AuthenticactionController(IAuthenticationService authenticationService) {
        _authenticationService = authenticationService;
    }

    [HttpPost("register")]
    public IActionResult Register(RegisterRequest request) {
        var authResult = _authenticationService.Register(request.FirstName,
                                                         request.LastName,
                                                         request.Email,
                                                         request.Password);
        var response = new AuthenticationResponse(authResult.Id,
                                                  authResult.FirstName,
                                                  authResult.LastName,
                                                  authResult.Email,
                                                  authResult.Token);
        return Ok(response);
    }

    [HttpPost("login")]
    public IActionResult Login(LoginRequest request) {
        var authResult = _authenticationService.Login(request.Email, request.Password);
        var response = new AuthenticationResponse(authResult.Id,
                                                  authResult.FirstName,
                                                  authResult.LastName,
                                                  authResult.Email,
                                                  authResult.Token);
        return Ok(response);
    }
}

