using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Admin.Services;

/// <summary>
/// 登录安全服务 — 对齐 V18 login_monitor.asp：
/// 记录登录成功/失败到 AdminLogs、生成分级 LoginAlerts，并对暴力破解自动封禁
/// （同一 IP 24h 内失败 >= 阈值 → 写入 IPBlacklist 封锁 7 天 + 严重告警）。
/// </summary>
public interface ILoginSecurityService
{
    Task RecordFailureAsync(string username, string ip, string? userAgent = null);
    Task RecordSuccessAsync(int adminId, string ip);
    /// <summary>手动"运行告警检查"，返回本次新封禁 IP 数。</summary>
    Task<int> RunAlertCheckAsync(int? operatorAdminId = null);
}

public class LoginSecurityService : ILoginSecurityService
{
    private readonly DbContextOptions<PerfumeShopContext> _options;

    // 对齐 V18：同 IP 24h 内失败 >= 5 次即封禁 7 天
    private const int FailThreshold = 5;
    private const int BlockDays = 7;

    public LoginSecurityService(DbContextOptions<PerfumeShopContext> options) => _options = options;

    public async Task RecordFailureAsync(string username, string ip, string? userAgent = null)
    {
        try
        {
            await using var db = new PerfumeShopContext(_options);
            db.AdminLogs.Add(new AdminLog
            {
                ActionType = "登录失败",
                ModuleCode = "auth",
                Notes = username,
                Ipaddress = ip,
                CreatedAt = DateTime.Now
            });
            await db.SaveChangesAsync();

            if (!string.IsNullOrEmpty(ip) && ip != "unknown")
                await EvaluateIpAsync(db, ip, null);
        }
        catch { /* 记录失败不阻断登录流程 */ }
    }

    public async Task RecordSuccessAsync(int adminId, string ip)
    {
        try
        {
            await using var db = new PerfumeShopContext(_options);
            db.AdminLogs.Add(new AdminLog
            {
                ActionType = "登录成功",
                ModuleCode = "auth",
                AdminId = adminId,
                Ipaddress = ip,
                CreatedAt = DateTime.Now
            });
            await db.SaveChangesAsync();
        }
        catch { }
    }

    public async Task<int> RunAlertCheckAsync(int? operatorAdminId = null)
    {
        var blocked = 0;
        try
        {
            await using var db = new PerfumeShopContext(_options);
            var since = DateTime.Now.AddHours(-24);
            var offenders = await db.AdminLogs.AsNoTracking()
                .Where(l => l.ActionType == "登录失败" && l.CreatedAt >= since
                            && l.Ipaddress != null && l.Ipaddress != "")
                .GroupBy(l => l.Ipaddress!)
                .Select(g => new { Ip = g.Key, Cnt = g.Count() })
                .Where(x => x.Cnt >= FailThreshold)
                .ToListAsync();

            foreach (var o in offenders)
                if (await EvaluateIpAsync(db, o.Ip, operatorAdminId, o.Cnt)) blocked++;
        }
        catch { }
        return blocked;
    }

    /// <summary>评估单个 IP 是否需封禁；返回是否新增封禁。</summary>
    private async Task<bool> EvaluateIpAsync(PerfumeShopContext db, string ip, int? operatorAdminId, int? knownCount = null)
    {
        var since = DateTime.Now.AddHours(-24);
        var count = knownCount ?? await db.AdminLogs.CountAsync(l =>
            l.ActionType == "登录失败" && l.Ipaddress == ip && l.CreatedAt >= since);
        if (count < FailThreshold) return false;

        var now = DateTime.Now;
        var alreadyBlocked = await db.Ipblacklists.AnyAsync(b =>
            b.Ipaddress == ip && b.IsActive == true && (b.ExpiresAt == null || b.ExpiresAt > now));

        if (!alreadyBlocked)
        {
            db.Ipblacklists.Add(new Ipblacklist
            {
                Ipaddress = ip,
                Reason = $"自动封禁：24h内{count}次登录失败",
                IsActive = true,
                BlockedBy = operatorAdminId,
                BlockedAt = now,
                ExpiresAt = now.AddDays(BlockDays),
                HitCount = 0
            });
            db.LoginAlerts.Add(new LoginAlert
            {
                AlertType = "auto_block",
                AlertLevel = "high",
                AlertMessage = $"自动封禁IP: {ip}（24h内{count}次登录失败）",
                Ipaddress = ip,
                IsRead = false,
                CreatedAt = now
            });
            await db.SaveChangesAsync();
            return true;
        }

        db.LoginAlerts.Add(new LoginAlert
        {
            AlertType = "repeat_attack",
            AlertLevel = "critical",
            AlertMessage = $"已封禁IP再次尝试: {ip}（24h内{count}次登录失败）",
            Ipaddress = ip,
            IsRead = false,
            CreatedAt = now
        });
        await db.SaveChangesAsync();
        return false;
    }
}
