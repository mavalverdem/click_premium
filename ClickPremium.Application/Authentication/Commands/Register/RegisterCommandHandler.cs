using ClickPremium.Application.Common.Interfaces.Authentication;
using ClickPremium.Application.Common.Interfaces.Persistence;

using ClickPremium.Domain.Entities;
using ErrorOr;
using MediatR;
using ClickPremium.Domain.Common.Errors;
using ClickPremium.Application.Authentication.Common;

namespace ClickPremium.Application.Authentication.Commands.Register;

public class RegisterCommandHandler : IRequestHandler<RegisterCommand, ErrorOr<AuthenticationResult>>
{
    private readonly IJwtTokenGenerator _jwtTokenGenerator;
    private readonly IUserRepository _userRepository;

    public RegisterCommandHandler(IJwtTokenGenerator jwtTokenGenerator, IUserRepository userRepository)
    {
        _jwtTokenGenerator = jwtTokenGenerator;
        _userRepository = userRepository;
    }
    public async Task<ErrorOr<AuthenticationResult>> Handle(RegisterCommand command, CancellationToken cancellationToken)
    {
        await Task.CompletedTask;

        // 1. Validate the user doesn't exist
        if (_userRepository.GetUserByEmail(command.Email) != null)
        {
            return Errors.User.DuplicateEmail;
        }
        // 2. Create user
        var user = new User
        {
            
            FirstName = command.FirstName,
            LastName = command.LastName,
            Email = command.Email,
            Password = command.Password
        };
        await _userRepository.Add(user);

        // 3. Create JWT token
        var token = _jwtTokenGenerator.GenerateToken(user);

        return new AuthenticationResult(
            user,
            token);
    }
}
