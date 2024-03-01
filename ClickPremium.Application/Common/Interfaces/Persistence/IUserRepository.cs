using ClickPremium.Domain.Entities;

namespace ClickPremium.Application.Common.Interfaces.Persistence
{
    public interface IUserRepository
    {
        User? GetUserByEmail(string email);
        Task<User> Add(User user);
    }
}
