using System.Net;
using System.Runtime.CompilerServices;
using ClickPremium.Application.Authentication.Commands.Register;
using ClickPremium.Application.Authentication.Common;
using ClickPremium.Application.Authentication.Queries.Login;
using ClickPremium.Application.Common.Errors;
using ClickPremium.Contracts.Authentication;
using ErrorOr;
using MapsterMapper;
using MediatR;
using Microsoft.AspNetCore.Mvc;


namespace  ClickPremium.Api.Controllers;


[Route("auth")]
public class AuthenticactionController: ApiController {

    private readonly ISender _mediator;
    private readonly IMapper _mapper;   

    public AuthenticactionController(ISender mediator, IMapper mapper)
    {
        _mediator = mediator;
        _mapper = mapper;
    }
  

    [HttpPost("register")]
    public async Task<IActionResult> Register(RegisterRequest request) {
        var command = _mapper.Map<RegisterCommand>(request);

        ErrorOr<AuthenticationResult> registerResult = await _mediator.Send(command);
        return registerResult.Match(
            authResult => Ok(_mapper.Map<AuthenticationResponse>(authResult)),
            Problem);
    }

    [HttpPost("login")]
    public async Task<IActionResult> Login(LoginRequest request) {

        var query = _mapper.Map<LoginQuery>(request);
        ErrorOr<AuthenticationResult> loginResult = await _mediator.Send(query);

        return loginResult.Match(
            authResult => Ok(_mapper.Map<AuthenticationResponse>(authResult)),
            errors  => Problem(errors));
    }
}