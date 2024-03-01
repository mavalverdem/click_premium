using FluentValidation;

namespace ClickPremium.Application.Authentication.Queries.Login
{
    public class LoginQueryValidator : AbstractValidator<LoginQuery>
    {
        public LoginQueryValidator()
        {
            // RuleFor(v => v.Email)
            //     .NotEmpty().WithMessage("Email is required")
            //     .EmailAddress().WithMessage("Email is not valid");

            RuleFor(v => v.Password)
                .NotEmpty().WithMessage("Password is required");
        }
    }
}