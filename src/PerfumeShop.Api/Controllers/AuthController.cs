using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Shared.Security;
using PerfumeShop.Shared.Services;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 用户认证 API — V19 M3-A 扩展
/// 对齐 V18 user/register.asp, user/forgot.asp, reset_pwd.asp
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class AuthController : ControllerBase
{
    private readonly IUserRepository _userRepo;
    private readonly PerfumeShopContext _db;
    private readonly IEmailService? _emailService;

    public AuthController(IUserRepository userRepo, PerfumeShopContext db, IEmailService? emailService = null)
    {
        _userRepo = userRepo;
        _db = db;
        _emailService = emailService;
    }

    /// <summary>用户登录</summary>
    [HttpPost("login")]
    public async Task<IActionResult> Login([FromBody] LoginRequest request)
    {
        if (string.IsNullOrEmpty(request.Username) || string.IsNullOrEmpty(request.Password))
            return BadRequest(new { message = "用户名和密码不能为空" });

        // V19 统一口令校验：按用户名/邮箱查询后用 PasswordHasher 验证（不再明文比较）
        var user = await _db.Users.FirstOrDefaultAsync(u =>
            (u.Username == request.Username || u.Email == request.Username) && u.IsActive == true);
        if (user == null || !PasswordHasher.Verify(request.Password, user.Password).Success)
            return Unauthorized(new { message = "用户名或密码错误" });

        return Ok(new
        {
            userId = user.UserId,
            username = user.Username,
            email = user.Email,
            fullName = user.FullName,
            tier = user.CustomerTier,
            points = user.Points
        });
    }

    /// <summary>用户注册 — 对齐 V18 register.asp</summary>
    [HttpPost("register")]
    public async Task<IActionResult> Register([FromBody] RegisterDto request)
    {
        if (string.IsNullOrEmpty(request.Username) || string.IsNullOrEmpty(request.Password) ||
            string.IsNullOrEmpty(request.Email))
            return BadRequest(new { message = "用户名、密码和邮箱不能为空" });

        // V19 密码强度校验: ≥8字符 + 大小写 + 数字
        if (request.Password.Length < 8)
            return BadRequest(new { message = "密码长度至少8个字符" });
        if (!System.Text.RegularExpressions.Regex.IsMatch(request.Password, @"[A-Z]") ||
            !System.Text.RegularExpressions.Regex.IsMatch(request.Password, @"[a-z]") ||
            !System.Text.RegularExpressions.Regex.IsMatch(request.Password, @"\d"))
            return BadRequest(new { message = "密码必须包含大写字母、小写字母和数字" });

        if (await _userRepo.UsernameExistsAsync(request.Username))
            return Conflict(new { message = "用户名已存在" });

        if (await _userRepo.EmailExistsAsync(request.Email))
            return Conflict(new { message = "邮箱已被注册" });

        var user = new User
        {
            Username = request.Username,
            Password = HashPasswordV3(request.Password),
            Email = request.Email,
            FullName = request.FullName,
            Phone = request.Phone,
            ReferrerUserId = request.ReferrerUserId,
            IsActive = true,
            CustomerTier = "bronze",
            Points = 0,
            CreatedAt = DateTime.Now
        };

        await _userRepo.AddAsync(user);
        await _userRepo.SaveChangesAsync();

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

        return Ok(new { userId = user.UserId, username = user.Username, message = "注册成功" });
    }

    /// <summary>请求密码重置 — 对齐 V18 forgot.asp</summary>
    [HttpPost("forgot-password")]
    public async Task<IActionResult> RequestPasswordReset([FromBody] ForgotPasswordDto request)
    {
        if (string.IsNullOrEmpty(request.Email))
            return BadRequest(new { message = "邮箱不能为空" });

        var user = await _userRepo.GetByEmailAsync(request.Email);
        if (user == null)
            return NotFound(new { message = "未找到使用该邮箱的账户" });

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
            catch { /* 邮件发送失败不阻塞请求 */ }
        }

        return Ok(new { message = "密码重置链接已发送到您的邮箱" });
    }

    /// <summary>重置密码 — 对齐 V18 reset_pwd.asp</summary>
    [HttpPost("reset-password")]
    public async Task<IActionResult> ResetPassword([FromBody] ResetPasswordDto request)
    {
        if (string.IsNullOrEmpty(request.Token) || string.IsNullOrEmpty(request.NewPassword))
            return BadRequest(new { message = "重置令牌和新密码不能为空" });

        if (request.NewPassword.Length < 8)
            return BadRequest(new { message = "密码长度至少8个字符" });

        var tokenHash = ComputeSha256Hash(request.Token);
        var resetToken = await _db.PasswordResetTokens
            .FirstOrDefaultAsync(t => t.TokenHash == tokenHash && !t.IsUsed && t.ExpiresAt > DateTime.Now);

        if (resetToken == null)
            return BadRequest(new { message = "重置令牌无效或已过期" });

        var user = await _userRepo.GetByIdAsync(resetToken.UserId);
        if (user == null)
            return NotFound(new { message = "用户不存在" });

        user.Password = HashPasswordV3(request.NewPassword);
        _userRepo.Update(user);

        resetToken.IsUsed = true;
        resetToken.UsedAt = DateTime.Now;

        await _db.SaveChangesAsync();

        return Ok(new { message = "密码重置成功，请使用新密码登录" });
    }

    /// <summary>获取当前用户信息</summary>
    [HttpGet("profile/{userId}")]
    public async Task<IActionResult> GetProfile(int userId)
    {
        var user = await _userRepo.GetByIdAsync(userId);
        if (user == null)
            return NotFound(new { message = "用户不存在" });

        return Ok(new
        {
            userId = user.UserId,
            username = user.Username,
            email = user.Email,
            fullName = user.FullName,
            phone = user.Phone,
            address = user.Address,
            city = user.City,
            tier = user.CustomerTier,
            points = user.Points,
            totalSpent = user.TotalSpent,
            orderCount = user.OrderCount
        });
    }

    // ========== 工具方法 ==========

    /// <summary>V3 密码哈希 — 对齐 V18 HashPassword()</summary>
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

// ========== DTOs ==========

public class LoginRequest
{
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
}

public class RegisterDto
{
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
    public string Email { get; set; } = "";
    public string? FullName { get; set; }
    public string? Phone { get; set; }
    public int? ReferrerUserId { get; set; }
}

public class ForgotPasswordDto
{
    public string Email { get; set; } = "";
}

public class ResetPasswordDto
{
    public string Token { get; set; } = "";
    public string NewPassword { get; set; } = "";
}
