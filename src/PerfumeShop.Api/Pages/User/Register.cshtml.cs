using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;
using PerfumeShop.Shared;
using PerfumeShop.Shared.Services;

namespace PerfumeShop.Api.Pages.User;

public class RegisterModel : PageModel
{
    private readonly PerfumeShopContext _db;
    private readonly ILocaleService _locale;
    private readonly IEmailService? _emailService;

    public RegisterModel(PerfumeShopContext db, ILocaleService locale, IEmailService? emailService = null)
    {
        _db = db;
        _locale = locale;
        _emailService = emailService;
    }

    public ILocaleService LocaleService => _locale;

    [BindProperty]
    public RegisterInput Input { get; set; } = new();

    public string? Error { get; set; }
    public string? Success { get; set; }

    public IActionResult OnGet()
    {
        // 已登录用户跳转个人中心
        if (User.Identity?.IsAuthenticated == true)
            return RedirectToPage("/User/Index");
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (User.Identity?.IsAuthenticated == true)
            return RedirectToPage("/User/Index");

        // 服务端校验
        if (string.IsNullOrWhiteSpace(Input.Username) || string.IsNullOrWhiteSpace(Input.Email) ||
            string.IsNullOrWhiteSpace(Input.Password))
        {
            Error = "用户名、邮箱和密码不能为空";
            return Page();
        }

        if (Input.Username.Length < 3 || Input.Username.Length > 20)
        {
            Error = "用户名长度需在3-20个字符之间";
            return Page();
        }

        if (Input.Password.Length < 8)
        {
            Error = "密码长度至少8个字符";
            return Page();
        }

        if (!Regex.IsMatch(Input.Password, @"[A-Z]") ||
            !Regex.IsMatch(Input.Password, @"[a-z]") ||
            !Regex.IsMatch(Input.Password, @"\d"))
        {
            Error = "密码必须包含大写字母、小写字母和数字";
            return Page();
        }

        if (Input.Password != Input.ConfirmPassword)
        {
            Error = "两次输入的密码不一致";
            return Page();
        }

        if (!Input.Email.Contains('@'))
        {
            Error = "邮箱格式不正确";
            return Page();
        }

        // 检查用户名/邮箱是否已存在
        var usernameExists = await _db.Users.AnyAsync(u => u.Username == Input.Username);
        if (usernameExists)
        {
            Error = "用户名已被使用";
            return Page();
        }

        var emailExists = await _db.Users.AnyAsync(u => u.Email == Input.Email);
        if (emailExists)
        {
            Error = "邮箱已被注册";
            return Page();
        }

        // 处理推荐码
        int? referrerUserId = null;
        if (!string.IsNullOrWhiteSpace(Input.ReferralCode))
        {
            var tokenHash = ComputeSha256Hash(Input.ReferralCode.Trim());
            var refToken = await _db.ReferralTokens
                .FirstOrDefaultAsync(t => t.TokenHash == tokenHash && t.IsActive == true && t.ExpiresAt > DateTime.Now);
            if (refToken != null)
            {
                referrerUserId = refToken.ReferrerUserId;
            }
        }

        // 创建用户
        var user = new PerfumeShop.Data.Models.User
        {
            Username = Input.Username.Trim(),
            Email = Input.Email.Trim(),
            Password = HashPasswordV3(Input.Password),
            FullName = Input.FullName?.Trim(),
            Phone = Input.Phone?.Trim(),
            ReferrerUserId = referrerUserId,
            IsActive = true,
            CustomerTier = "bronze",
            Points = 0,
            CreatedAt = DateTime.Now
        };

        _db.Users.Add(user);
        await _db.SaveChangesAsync();

        // 注册成功后自动登录
        var claims = new List<Claim>
        {
            new(ClaimTypes.NameIdentifier, user.UserId.ToString()),
            new(ClaimTypes.Name, user.Username),
            new(ClaimTypes.Email, user.Email)
        };
        var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
        var principal = new ClaimsPrincipal(identity);
        await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal,
            new AuthenticationProperties { IsPersistent = true });

        // 合并匿名购物车
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

        // 发送欢迎邮件 (best-effort)
        if (_emailService != null)
        {
            try
            {
                await _emailService.SendTemplateAsync("welcome", new EmailTemplateModel
                {
                    Variables = new Dictionary<string, string>
                    {
                        ["username"] = user.Username,
                        ["fullName"] = user.FullName ?? user.Username
                    }
                }, user.Email);
            }
            catch { /* 邮件发送失败不阻塞注册 */ }
        }

        return RedirectToPage("/User/Index");
    }

    // ========== 工具方法 ==========

    private static string HashPasswordV3(string password)
    {
        var saltBytes = RandomNumberGenerator.GetBytes(16);
        var salt = Convert.ToBase64String(saltBytes);
        using var pbkdf2 = new Rfc2898DeriveBytes(password, saltBytes, 10000, HashAlgorithmName.SHA256);
        var hash = Convert.ToBase64String(pbkdf2.GetBytes(32));
        return $"V3${salt}${hash}";
    }

    private static string ComputeSha256Hash(string input)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(input));
        return Convert.ToHexString(bytes).ToLower();
    }
}

public class RegisterInput
{
    public string Username { get; set; } = "";
    public string Email { get; set; } = "";
    public string Password { get; set; } = "";
    public string ConfirmPassword { get; set; } = "";
    public string? FullName { get; set; }
    public string? Phone { get; set; }
    public string? ReferralCode { get; set; }
}
