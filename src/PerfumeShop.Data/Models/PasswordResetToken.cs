using System;

namespace PerfumeShop.Data.Models;

/// <summary>
/// 密码重置令牌 — 对齐 V18 Users.ResetToken 字段独立化
/// </summary>
public class PasswordResetToken
{
    public int TokenId { get; set; }
    public int UserId { get; set; }
    public string Token { get; set; } = null!;
    public string TokenHash { get; set; } = null!;
    public DateTime ExpiresAt { get; set; }
    public bool IsUsed { get; set; }
    public DateTime CreatedAt { get; set; }
    public DateTime? UsedAt { get; set; }
}
