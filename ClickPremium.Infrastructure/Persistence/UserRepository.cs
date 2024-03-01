using ClickPremium.Application.Common.Interfaces.Persistence;
using ClickPremium.Domain.Entities;
using ClickPremium.Infrastructure.Models;
using ClickPremium.Infrastructure.Models.cfg;
using MapsterMapper;


namespace ClickPremium.Infrastructure.Persistence
{
    public class UserRepository : IUserRepository
    {
        private readonly ClickpremiumcfgContext _dbContext;
        private readonly IMapper _mapper;

        public UserRepository(ClickpremiumcfgContext dbContext, IMapper mapper)
        {
            _dbContext = dbContext;
            _mapper = mapper;
        }   
        
        public User GetUserByEmail(string email)
        {
            var x = _dbContext.Sgusrs.ToList();
            Sgusr? sgusr = _dbContext.Sgusrs.FirstOrDefault(u => u.Codusr == email);
            var User = _mapper.Map<User>(sgusr ?? new Sgusr());
            return User;
        }
        public async Task<User> Add(User user)
        {
            Sgusr sgusr = _mapper.Map<Sgusr>(user);
            var entity =  _dbContext.Sgusrs.Add(sgusr);
            await _dbContext.SaveChangesAsync();
            return _mapper.Map<User>(entity.Entity); 
        }
    }
}
