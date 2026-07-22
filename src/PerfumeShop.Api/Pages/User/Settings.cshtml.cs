using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

public class SettingsModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public SettingsModel(PerfumeShopContext db) => _db = db;

    public string Email { get; set; } = "";
    public string FullName { get; set; } = "";
    public string Phone { get; set; } = "";
    public string? AvatarUrl { get; set; }
    public bool EmailSubscribed { get; set; } = true;

    public async Task OnGetAsync()
    {
        var userId = 1;
        var user = await _db.Users.FindAsync(userId);
        if (user != null)
        {
            Email = user.Email;
            FullName = user.FullName ?? "";
            Phone = user.Phone ?? "";
        }

        // TODO: EmailSubscribed 可从 UserPreference 扩展字段获取
    }
}
