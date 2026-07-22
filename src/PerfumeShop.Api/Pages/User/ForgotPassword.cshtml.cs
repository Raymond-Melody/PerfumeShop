using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;
using PerfumeShop.Shared;
using PerfumeShop.Shared.Services;

namespace PerfumeShop.Api.Pages.User;

public class ForgotPasswordModel : PageModel
{
    private readonly PerfumeShopContext _db;
    private readonly ILocaleService _locale;
    private readonly IEmailService? _emailService;

    public ForgotPasswordModel(PerfumeShopContext db, ILocaleService locale, IEmailService? emailService = null)
    {
        _db = db;
        _locale = locale;
        _emailService = emailService;
    }

    public ILocaleService LocaleService => _locale;

    [BindProperty]
    public ForgotPasswordInput Input { get; set; } = new();

    public string? Error { get; set; }
    public string? Success { get; set; }

    public IActionResult OnGet()
    {
        if (User.Identity?.IsAuthenticated == true)
            return RedirectToPage("/User/Index");
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (User.Identity?.IsAuthenticated == true)
            return RedirectToPage("/User/Index");

        if (string.IsNullOrWhiteSpace(Input.Email))
        {
            Error = "请输入邮箱地址";
            return Page();
        }

        var user = await _db.Users.FirstOrDefaultAsync(u => u.Email == Input.Email.Trim());
        if (user == null)
        {
            Error = "未找到使用该邮箱的账户";
            return Page();
        }

        // 生成 64 位随机 Token
        var tokenBytes = RandomNumberGenerator.GetBytes(32);
        var token = Convert.ToHexString(tokenBytes).ToLower();
        var tokenHash = ComputeSha256Hash(token);

        var resetToken = new PasswordResetToken
        {
            UserId = user.UserId,
            Token = token,
            TokenHash = tokenHash,
            ExpiresAt = DateTime.Now.AddHours(1),
            IsUsed = false,
            CreatedAt = DateTime.Now
        };

        _db.PasswordResetTokens.Add(resetToken);
        await _db.SaveChangesAsync();

        // 发送重置邮件 (best-effort)
        if (_emailService != null)
        {
            try
            {
                var resetUrl = $"{Request.Scheme}://{Request.Host}/user/reset-password?token={token}";
                await _emailService.SendTemplateAsync("password-reset", new EmailTemplateModel
                {
                    Variables = new Dictionary<string, string>
                    {
                        ["username"] = user.Username,
                        ["fullName"] = user.FullName ?? user.Username,
                        ["resetUrl"] = resetUrl,
                        ["token"] = token
                    }
                }, user.Email);
            }
            catch { /* 邮件发送失败不阻塞 */ }
        }

        Success = "密码重置链接已发送到您的邮箱，请在1小时内完成操作。";
        return Page();
    }

    private static string ComputeSha256Hash(string input)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(input));
        return Convert.ToHexString(bytes).ToLower();
    }
}

public class ForgotPasswordInput
{
    public string Email { get; set; } = "";
}
