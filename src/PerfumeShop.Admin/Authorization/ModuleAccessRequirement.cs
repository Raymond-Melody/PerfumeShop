using Microsoft.AspNetCore.Authorization;
using PerfumeShop.Admin.Services;

namespace PerfumeShop.Admin.Authorization;

/// <summary>模块访问授权要求 — 对齐 V18 VerifyModuleAccess(moduleCode, level)。</summary>
public class ModuleAccessRequirement : IAuthorizationRequirement
{
    public string Module { get; }
    public ModuleAccessRequirement(string module) => Module = module;
}

/// <summary>
/// 模块访问授权处理器：从 ClaimsPrincipal 读取 RoleCode，经 IPermissionService 判定模块访问权。
/// 使用 IServiceScopeFactory 解析作用域内的 IPermissionService（DbContext），
/// 从 context.User 取身份，兼容 HTTP 请求与 Blazor 交互态（无 HttpContext）。
/// </summary>
public class ModuleAccessHandler : AuthorizationHandler<ModuleAccessRequirement>
{
    private readonly IServiceScopeFactory _scopeFactory;
    public ModuleAccessHandler(IServiceScopeFactory scopeFactory) => _scopeFactory = scopeFactory;

    protected override async Task HandleRequirementAsync(
        AuthorizationHandlerContext context, ModuleAccessRequirement requirement)
    {
        if (context.User?.Identity?.IsAuthenticated != true) return;

        var roleCode = context.User.FindFirst("RoleCode")?.Value;

        using var scope = _scopeFactory.CreateScope();
        var perms = scope.ServiceProvider.GetRequiredService<IPermissionService>();
        if (await perms.CanAccessAsync(roleCode, requirement.Module))
            context.Succeed(requirement);
    }
}

/// <summary>模块授权策略常量与注册辅助。</summary>
public static class ModulePolicies
{
    public static readonly string[] Modules =
    {
        "operation", "finance", "purchase", "techcenter", "system",
        "prodcenter", "semifinished", "logistics", "inventory", "analytics"
    };

    /// <summary>策略名：Module.{code}（如 Module.system）。</summary>
    public static string PolicyName(string module) => $"Module.{module}";
}
