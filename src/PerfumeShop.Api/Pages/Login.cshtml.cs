using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

/// <summary>
/// V19 登录页 — 对齐 V18 user/login.asp
/// 增强：Cookie Authentication + 匿名购物车合并 + 记住我
/// </summary>
public class LoginModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public LoginModel(PerfumeShopContext db) => _db = db;

    public string? Error { get; set; }
    public string? Success { get; set; }

    public void OnGet()
    {
        if (User.Identity?.IsAuthenticated == true)
        {
            Response.Redirect("/user");
            return;
        }

        var msg = Request.Query["msg"].FirstOrDefault();
        if (msg == "registered")
            Success = "注册成功！请登录您的账户。";
        else if (msg == "pwd_changed")
            Success = "密码修改成功！请使用新密码登录。";
    }

    public async Task<IActionResult> OnPostAsync(string? email, string? password, bool remember = false, string? returnUrl = null)
    {
        if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password))
        {
            Error = "请输入邮箱和密码";
            return Page();
        }

        var user = await _db.Users
            .FirstOrDefaultAsync(u => (u.Email == email || u.Username == email) && u.IsActive == true);

        if (user == null)
        {
            Error = "邮箱或密码错误";
            return Page();
        }

        // 密码验证 — 支持 V3$ 哈希格式
        if (!VerifyPassword(password, user.Password))
        {
            Error = "邮箱或密码错误";
            return Page();
        }

        // 登录成功 — Cookie Authentication (对齐 V18 Session 机制)
        var claims = new List<Claim>
        {
            new(ClaimTypes.NameIdentifier, user.UserId.ToString()),
            new(ClaimTypes.Name, user.Username),
            new(ClaimTypes.Email, user.Email),
            new("FullName", user.FullName ?? user.Username),
            new("CustomerTier", user.CustomerTier ?? "bronze")
        };

        var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
        var principal = new ClaimsPrincipal(identity);

        var authProperties = new AuthenticationProperties
        {
            IsPersistent = remember,
            ExpiresUtc = remember ? DateTimeOffset.UtcNow.AddDays(30) : null
        };

        await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal, authProperties);

        // 合并匿名购物车 — 对齐 V18 login.asp MergeAnonymousCart
        var sessionId = HttpContext.Session.Id;
        var anonCartItems = await _db.Carts
            .Where(c => c.SessionId == sessionId && c.UserId == null)
            .ToListAsync();

        if (anonCartItems.Any())
        {
            foreach (var item in anonCartItems)
            {
                item.UserId = user.UserId;
                item.SessionId = null;
            }
            await _db.SaveChangesAsync();
        }

        // 跳转
        if (!string.IsNullOrEmpty(returnUrl) && Url.IsLocalUrl(returnUrl))
            return Redirect(returnUrl);

        return RedirectToPage("/User/Index");
    }

    /// <summary>V3 密码验证 — 对齐 V18 VerifyPassword()</summary>
    private static bool VerifyPassword(string input, string stored)
    {
        if (string.IsNullOrEmpty(stored)) return false;

        if (stored.StartsWith("V3$"))
        {
            var parts = stored.Split('$');
            if (parts.Length != 3) return false;
            var salt = parts[1];
            var storedHash = parts[2];
            var saltBytes = Convert.FromBase64String(salt);
            using var pbkdf2 = new Rfc2898DeriveBytes(input, saltBytes, 10000, HashAlgorithmName.SHA256);
            var computedHash = Convert.ToBase64String(pbkdf2.GetBytes(32));
            return computedHash == storedHash;
        }

        // 向后兼容明文 (逐步淘汰)
        return input == stored;
    }
}
