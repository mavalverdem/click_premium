using System.Net;

namespace ClickPremium.Application.Common.Errors
{
    public class DuplicateEmailException : Exception, IServiceException
    {
        public HttpStatusCode StatusCode => HttpStatusCode.Conflict;

        public string ErrorMessage => "Email already exists in the system.";
    }
}
