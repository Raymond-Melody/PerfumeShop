using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;
using PerfumeShop.Shared;
using PerfumeShop.Shared.Security;

namespace PerfumeShop.Api.Pages.User;

public class ResetPasswordModel : PageModel
{
    private readonly PerfumeShopContext _db;
    private readonly ILocaleService _locale;

    public ResetPasswordModel(PerfumeShopContext db, ILocaleService locale)
    {
        _db = db;
        _locale = locale;
    }

    public ILocaleService LocaleService => _locale;

    [BindProperty]
    public ResetPasswordInput Input { get; set; } = new();

    public string? Error { get; set; }
    public string? Success { get; set; }
    public bool TokenValid { get; set; }

    public async Task<IActionResult> OnGetAsync(string? token)
    {
        if (string.IsNullOrEmpty(token))
        {
            TokenValid = false;
            return Page();
        }

        var tokenHash = ComputeSha256Hash(token);
        var resetToken = await _db.PasswordResetTokens
            .FirstOrDefaultAsync(t => t.TokenHash == tokenHash && !t.IsUsed && t.ExpiresAt > DateTime.Now);

        if (resetToken == null)
        {
            TokenValid = false;
            return Page();
        }

        TokenValid = true;
        Input.Token = token;
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (string.IsNullOrEmpty(Input.Token))
        {
            TokenValid = false;
            return Page();
        }

        var tokenHash = ComputeSha256Hash(Input.Token);
        var resetToken = await _db.PasswordResetTokens
            .FirstOrDefaultAsync(t => t.TokenHash == tokenHash && !t.IsUsed && t.ExpiresAt > DateTime.Now);

        if (resetToken == null)
        {
            TokenValid = false;
            return Page();
        }

        TokenValid = true;

        if (string.IsNullOrWhiteSpace(Input.NewPassword) || Input.NewPassword.Length < 8)
        {
            Error = "密码长度至少8个字符";
            return Page();
        }

        if (!Regex.IsMatch(Input.NewPassword, @"[A-Z]") ||
            !Regex.IsMatch(Input.NewPassword, @"[a-z]") ||
            !Regex.IsMatch(Input.NewPassword, @"\d"))
        {
            Error = "密码必须包含大写字母、小写字母和数字";
            return Page();
        }

        if (Input.NewPassword != Input.ConfirmPassword)
        {
            Error = "两次输入的密码不一致";
            return Page();
        }

        var user = await _db.Users.FirstOrDefaultAsync(u => u.UserId == resetToken.UserId);
        if (user == null)
        {
            Error = "用户不存在";
            return Page();
        }

        // 更新密码
        user.Password = HashPasswordV3(Input.NewPassword);
        resetToken.IsUsed = true;
        resetToken.UsedAt = DateTime.Now;

        await _db.SaveChangesAsync();

        Success = "密码重置成功！请使用新密码登录。";
        return Page();
    }

    private static string HashPasswordV3(string password)
    {
        // V19 统一口令散列 — 委托共享 PasswordHasher（迭代SHA-256+pepper）
        return PasswordHasher.Hash(password);
    }

    private static string ComputeSha256Hash(string input)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(input));
        return Convert.ToHexString(bytes).ToLower();
    }
}

public class ResetPasswordInput
{
    public string Token { get; set; } = "";
    public string NewPassword { get; set; } = "";
    public string ConfirmPassword { get; set; } = "";
}
