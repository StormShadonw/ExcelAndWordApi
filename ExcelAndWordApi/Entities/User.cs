using Microsoft.AspNetCore.Identity;

namespace ExcelAndWordApi.Entities
{
    public class User: IdentityUser
    {
        public string UserId { get; set; } = default!;
    }
}
