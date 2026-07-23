using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Admin.Services;

/// <summary>
/// 管理员操作审计器 — 对齐 V18 role_auth.asp 的 LogAdminAction。
/// 写入 AdminAuditLog（与 SysAdmin/AuditLog.razor 及 V18 audit_logs.asp 读取的是同一张表，
/// 从而消除此前 AuditService 写 AuditLogs / 页面读 AdminAuditLogs 的表名分歧）。
/// 使用独立 DbContext 实例，避免把调用方页面里跟踪中的实体一并提交。
/// </summary>
public interface IAdminAuditor
{
    // actor（adminId/adminName）由调用方从 AuthenticationState 提供，
    // 以保证 Blazor 交互态（无 HttpContext）下也能正确记录操作人。
    Task LogAsync(int adminId, string adminName, string actionType, string? targetType = null,
        int? targetId = null, string? targetName = null, string? details = null);
}

public class AdminAuditor : IAdminAuditor
{
    private readonly DbContextOptions<PerfumeShopContext> _options;
    private readonly ICurrentUserAccessor _user;

    public AdminAuditor(DbContextOptions<PerfumeShopContext> options, ICurrentUserAccessor user)
    {
        _options = options;
        _user = user;
    }

    public async Task LogAsync(int adminId, string adminName, string actionType, string? targetType = null,
        int? targetId = null, string? targetName = null, string? details = null)
    {
        try
        {
            await using var db = new PerfumeShopContext(_options);
            db.AdminAuditLogs.Add(new AdminAuditLog
            {
                AdminId = adminId,
                AdminName = adminName,
                ActionType = actionType,
                TargetType = targetType,
                TargetId = targetId,
                TargetName = targetName,
                Details = details,
                Ipaddress = _user.IpAddress,
                UserAgent = _user.UserAgent,
                CreatedAt = DateTime.Now
            });
            await db.SaveChangesAsync();
        }
        catch
        {
            // 审计失败不阻断主流程 — 对齐 V18 On Error Resume Next
        }
    }
}
