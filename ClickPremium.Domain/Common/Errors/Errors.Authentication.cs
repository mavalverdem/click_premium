using ErrorOr;

namespace ClickPremium.Domain.Common.Errors
{
    public static partial class Errors
    {
        public static class Authentication
        {
            public static Error InvalidCredentials => Error.Validation(code:"Auth.InvalidadCrecentials", description: "Invalid Credentials");

        }
    }
}