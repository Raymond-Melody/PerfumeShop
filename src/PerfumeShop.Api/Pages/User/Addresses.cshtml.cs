using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

public class AddressesModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public AddressesModel(PerfumeShopContext db) => _db = db;

    public List<UserAddress> Addresses { get; set; } = new();
    public bool IsEmpty => Addresses.Count == 0;

    public async Task OnGetAsync()
    {
        var userId = 1;
        Addresses = await _db.UserAddresses
            .Where(a => a.UserId == userId)
            .OrderByDescending(a => a.IsDefault)
            .ThenByDescending(a => a.CreatedAt)
            .ToListAsync();
    }
}
