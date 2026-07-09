using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

public class LoginModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public LoginModel(PerfumeShopContext db) => _db = db;

    public string? Error { get; set; }

    public void OnGet() { }

    public async Task<IActionResult> OnPostAsync(string email, string password)
    {
        var user = await _db.Users.FirstOrDefaultAsync(u => u.Email == email);
        if (user == null || user.Password != password) // 简化版，生产环境应用 hash
        {
            Error = "邮箱或密码错误";
            return Page();
        }
        // 简单 cookie 认证 (实际应用应用 JWT/Claims)
        Response.Cookies.Append("UserId", user.UserId.ToString(), new CookieOptions { HttpOnly = true, MaxAge = TimeSpan.FromDays(7) });
        return RedirectToPage("/Index");
    }
}
