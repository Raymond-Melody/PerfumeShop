using System;

namespace PerfumeShop.Data.Models;

/// <summary>
/// 认证令牌实体 — 支持 ASP 与 ASP.NET Core 双系统 Session 互通
/// V18 侧登录成功后写入此表，V19 侧通过 AuthBridgeMiddleware 读取验证
/// </summary>
public partial class AuthToken
{
    /// <summary>主键</summary>
    public int TokenId { get; set; }

    /// <summary>关联用户 ID（FK → Users.UserID）</summary>
    public int UserId { get; set; }

    /// <summary>SHA-256 哈希后的 Token（64 位十六进制字符串）</summary>
    public string Token { get; set; } = null!;

    /// <summary>Token 创建时间（UTC）</summary>
    public DateTime CreatedAt { get; set; }

    /// <summary>Token 过期时间（UTC）</summary>
    public DateTime ExpiresAt { get; set; }

    /// <summary>来源：UserLogin / AdminLogin</summary>
    public string Source { get; set; } = null!;

    /// <summary>是否仍然有效（可主动撤销）</summary>
    public bool IsActive { get; set; } = true;

    /// <summary>客户端 IP 地址</summary>
    public string? IpAddress { get; set; }
}
