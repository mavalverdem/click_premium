using Bubber.Core.Application.Common.Interfaces.Services;

namespace Bubber.Core.Infrastructure.Services
{
    public class DateTimeProvider : IDateTimeProvider
    {
        public DateTime UtcNow => DateTime.UtcNow;
    }
}