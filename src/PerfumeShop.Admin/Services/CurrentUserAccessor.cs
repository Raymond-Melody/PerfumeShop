using System.Security.Claims;
using Microsoft.AspNetCore.Http;

namespace PerfumeShop.Admin.Services;

/// <summary>
/// 统一当前登录管理员访问器 — 读取 Login.razor 下发的 Claims（AdminId/UserName/RoleCode）
/// 及请求 IP/UserAgent。供审计埋点、RBAC 判定、登录监控复用。
/// </summary>
public interface ICurrentUserAccessor
{
    int AdminId { get; }
    string UserName { get; }
    string RoleCode { get; }
    string IpAddress { get; }
    string UserAgent { get; }
    bool IsAuthenticated { get; }
    bool IsSuperAdmin { get; }
}

public class CurrentUserAccessor : ICurrentUserAccessor
{
    private readonly IHttpContextAccessor _http;
    public CurrentUserAccessor(IHttpContextAccessor http) => _http = http;

    private ClaimsPrincipal? User => _http.HttpContext?.User;

    public int AdminId => int.TryParse(User?.FindFirst(ClaimTypes.NameIdentifier)?.Value, out var id) ? id : 0;
    public string UserName => User?.FindFirst(ClaimTypes.Name)?.Value ?? "";
    public string RoleCode => User?.FindFirst("RoleCode")?.Value ?? "";
    public bool IsAuthenticated => User?.Identity?.IsAuthenticated == true;
    public bool IsSuperAdmin => RoleCode == "SUPER_ADMIN";

    public string IpAddress
    {
        get
        {
            var ctx = _http.HttpContext;
            if (ctx == null) return "";
            var fwd = ctx.Request.Headers["X-Forwarded-For"].FirstOrDefault();
            if (!string.IsNullOrEmpty(fwd)) return fwd.Split(',')[0].Trim();
            return ctx.Connection.RemoteIpAddress?.ToString() ?? "";
        }
    }

    public string UserAgent => _http.HttpContext?.Request.Headers.UserAgent.ToString() ?? "";
}
